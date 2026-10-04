"""New Outlook for Mac local store (HxStore.hxd), read from an in-memory snapshot.

A Python port of the field walk in securized/hxstore-reverse-engineering (hxprobe,
record.rs + schema.rs, MIT). The file is a series of CRC-checked blocks holding LZ4
payloads. Mail records anchor on the UTF-16LE ItemClass "IPM.Note", meeting items on their
own ItemClass; around it sit NUL-terminated UTF-16LE fields whose byte offsets drift but
whose order is stable, so fields are claimed by position in the run sequence plus a type
check, never by offset.

A decompressed block is a run of typed objects, each with a local object id (after the record
model of mitchell-johnson/hxstore-decode, MIT; the layout and type codes were found on this
store). A message is a header object (sender, subject, preview,
Message-ID, delivery time) and a body object (the ItemClass, then the HTML) that carry the
same id; that id is the only link between them, and every revision and copy of the message
carries it too. collect() merges records by that id, and each field walk stays inside the
record's own object. A body that no header claims by id falls back to the text match (the
adjacent header, or the one whose preview opens the body), never onto a header that has a
body object of its own. Most anchors outside header objects are conversation summaries; one
that repeats a cached message is folded into it.

Mail is only ever a header, body or summary object (0x0430). The ItemClass also sits in other
objects, such as the action backups and search caches (~/ActionsBackup/...: search cards,
people suggestions, image names), and those anchors are skipped, as is every anchor after the
first in one object: a summary names the class of each kind of item it covers, with names packed
behind the later ones. A summary holds its From address first and then the address its list
shows, which for the user's own mail is the recipient; the addresses behind a summary are the
next summaries' senders, not its recipients. A summary that carries a message's own local id is
the one Outlook writes while the user composes that message (on this store they share one
parent id at object start +37); its fields are left over from earlier saves, so it is skipped.
Display names (one name stored twice, or names packed behind a length each) are never a
preview, Outlook's signature line never a subject or a name, and the subject of the user's own
reply never a name. In a header object a lone subject is the topic stored right behind its own copy, not the first run
that looks like one (that is the preview, or the recipients' names when the subject is empty).
A subject still being typed when a draft was first saved can stay behind the names as the topic;
a longer subject that extends it and is stored twice whole is the one the mail went out with.

Every header, body and summary object carries a change number (u64 at object start +117) that
rises with every save of the item; the delivery time is no order, as a later save of a draft
can carry an earlier one. Sender, name and subject come from the copies of the last save, the
preview and time from the last save that has them, and of a message's own saves only the last
keeps its preview searchable. Next to a draft's text Outlook keeps a second body object with only
the quoted mail, or nothing, and that one can carry the higher change number (on one forward its
EFMData file is the older of the two), so the body is the one with the highest change number that
holds the preview of the header's last save, else the one with the highest change number.

A body too large for its object is kept whole in a gzip file of HTML under Files/S*/*/EFMData/
in the profile folder; the body object then names the file (~/Files/...) and holds only the
opening part of the HTML, or none of it. The file is read into memory and replaces the opening
part when its text holds the start of that part.

A meeting item (a meeting request, its cancellation, a reply to it) is stored like a message,
a header object and a body object under one id, with its own ItemClass. Two more objects carry
that id: a time object (the end, then the start, as UTC ticks at fixed offsets) and, for a
request, a place object (the organizer's name and address, then the location). Each meeting
item is a doc of its own, never merged with or folded into mail; its conversation summaries are
skipped, as each one seen repeats a cached item. A report that a meeting request was not
delivered is mail.

The store is a cache. Almost every message whose header object is here has its body object
too; what keeps only a ~255 character preview is mostly a conversation summary whose message
is not cached. Whether Outlook keeps the body of mail never opened in the client is not
established.
"""
import bisect
import gzip
import hashlib
import html
import re
import struct
import unicodedata
import zlib
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path

import lz4.block

from .index import INDEX_PATH

PROFILES = Path.home() / "Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles"
TEXT_CACHE = INDEX_PATH.parent / "attachment_text"
ATTACHMENT_NAME = re.compile(r"^[^/\\<>@]{1,200}\.(pdf|docx|xlsx|pptx|doc|xls|ppt|csv|txt|eml|msg|zip)$", re.I)
BLOCK_MAGIC = bytes.fromhex("056a703b6445025d")
ANCHOR = "IPM.Note".encode("utf-16-le")
# meeting items: a meeting request, its cancellation and the replies to it, each with the line its doc starts with
MEETING_CLASSES = {"IPM.Schedule.Meeting.Request": "Meeting request",
                   "IPM.Schedule.Meeting.Canceled": "Meeting canceled",
                   "IPM.Schedule.Meeting.Resp.Pos": "Meeting accepted",
                   "IPM.Schedule.Meeting.Resp.Neg": "Meeting declined",
                   "IPM.Schedule.Meeting.Resp.Tent": "Meeting tentatively accepted"}
# a delivery report of a meeting request is a report, so mail; anchored whole so its class run is not read as a field
NDR_CLASS = "REPORT.IPM.Schedule.Meeting.Request.NDR"
CLASS_PREFIXES = tuple(c.lower() for c in sorted([*MEETING_CLASSES, NDR_CLASS, "IPM.Note"], key=len, reverse=True))
ANCHORS = re.compile(b"|".join(re.escape(c.encode("utf-16-le")) for c in
                               sorted([*MEETING_CLASSES, NDR_CLASS, "IPM.Note"], key=len, reverse=True)))

UNIX_EPOCH_TICKS = 621_355_968_000_000_000
TICK_LO = UNIX_EPOCH_TICKS + int(datetime(2015, 1, 1).timestamp()) * 10_000_000
TICK_HI = UNIX_EPOCH_TICKS + int((datetime.now() + timedelta(days=1)).timestamp()) * 10_000_000

PAIR_WINDOW = 24 * 3600 * 10_000_000  # one day in ticks: send -> delivery
SENDER_MAX_DIST = 320     # the sender pair sits within ~120 bytes; beyond this is the previous record
VOTED = ("sender", "name", "subject", "preview", "ts")  # fields a message takes from its last save
FILLED = ("preview", "ts")  # of these, from the last save that has one; an earlier save of a draft has stray names
SUBJECT_MAX_REL = 4096
SUBJECT_SOLO_MAX_REL = 2500
RECORD_MAX = 10_000          # the last record of a block otherwise runs to the block end, over unrelated objects
BODY_MATCH_CHARS = 1200      # a header's preview is the start of its body text (after any banner)
HTML_START = (b"<html", b"<HTML", b"<!DOC", b"<body", b"<div", b"<tab", b"<met", b"<head")

# A decompressed block is a run of objects: flag byte, u32 size, 05 00, u16 type, the size again, 00 00, then the
# first field: u16 tag, 8 zero bytes, u64 object id. The next object starts OBJ_TAIL bytes after the size.
# The pattern starts at the fixed 05 00, five bytes in: a literal start is found far faster than the flag byte.
OBJECT = re.compile(rb"\x05\x00(.{2})(.{4})\x00\x00", re.S)
OBJ_TAIL = 15
HEADER_OBJ = 0x060F  # MessageHeader: sender, subject, preview, Message-ID; its id is the message's local id
BODY_OBJ = 0x0740    # MessageData: the ItemClass, then the HTML body; same id as its header object
CONV_OBJ = 0x0430    # conversation summary: latest sender and preview, no body, no id of one message
RECEIVED_AT = 293    # .NET ticks of the delivery time (ReceivedTime) in a header object
CN_AT = 117          # u64 change number in a header, body or summary object: every save of an item raises it
TIME_OBJ = 0x014E    # a meeting item's times, same id as its header object: UTC ticks of the end, then the start
END_AT, START_AT = 133, 157
PLACE_OBJ = 0x015E   # a meeting request's organizer name, address and location, same id; the type also holds contacts
# a body too large for its object is kept whole in a gzip file the object names, ~ being the profile folder;
# the object then holds only its opening part, or nothing
BODY_FILE = re.compile(rb"~\x00/\x00F\x00i\x00l\x00e\x00s\x00/\x00(?:[\x21-\x7e]\x00)+?\.\x00d\x00a\x00t\x00\x00\x00")

# A run starts on a printable ASCII code unit and continues through ASCII, other BMP
# characters and surrogate pairs; a NUL code unit ends it.
_RUN = re.compile(
    rb"(?:[\x20-\x7e]\x00|[\xa0-\xff]\x00|[\x00-\xff][\x01]"        # start: ASCII, Latin-1, Latin Ext-A
    rb")(?:[\x20-\x7e]\x00|[\xa0-\xff]\x00|[\x00-\xff][\x01-\xd7]"  # BMP up to the surrogates
    rb"|[\x00-\xff][\xe0-\xfe]|[\x00-\xfc]\xff"                      # BMP above the surrogates
    rb"|[\x00-\xff][\xd8-\xdb][\x00-\xff][\xdc-\xdf])*",             # surrogate pair
    re.S,
)
ENUM_VALUES = {v.lower() for v in ("IPM.Note", "Anonymous", "Internal", "External", "Normal", "None", "Unknown",
                                   "Focused", "Other", "Mailbox", "OneOff", "SMTP", "Personal", "Confidential",
                                   *MEETING_CLASSES, NDR_CLASS)}
FILE_EXT = {"csv", "pdf", "dat", "png", "jpg", "jpeg", "gif", "xlsx", "docx", "zip", "eml"}
REPLY_PREFIXES = ("re:", "fw:", "fwd:", "re :", "aw:", "tr:", "vs:", "vá:", "továbbítás:")
BARE_NAME = re.compile(r"[\w.+-]*\.[A-Za-z]{2,12}")  # a host or file name, or the tail of an address: no preview
# Outlook's own signature line: a body, never a subject or a name
SIGNATURE = re.compile(r"(?:Sent from|Get) Outlook for (?:Mac|iOS|Android|Windows)", re.I)


def stores() -> list[Path]:
    return sorted(PROFILES.glob("*/HxStore.hxd"))


# -- blocks ---------------------------------------------------------------------

def _blocks(data: bytes, stats: Counter):
    i = data.find(BLOCK_MAGIC)
    while i != -1:
        b = i - 8
        if b >= 0:
            hdr_crc, pay_crc = struct.unpack_from("<II", data, b)
            plen, ilen = struct.unpack_from("<II", data, b + 0x14)
            end = b + 0x28 + plen
            if end <= len(data) and zlib.crc32(data[b + 4:b + 0x20]) == hdr_crc \
                    and zlib.crc32(data[b + 8:end]) == pay_crc:
                try:
                    yield lz4.block.decompress(data[b + 0x28:end], uncompressed_size=ilen)
                    stats["blocks_ok"] += 1
                except Exception:
                    stats["blocks_bad_lz4"] += 1
            else:
                stats["blocks_bad_crc"] += 1  # a few, while Outlook writes the file under us
        i = data.find(BLOCK_MAGIC, i + 1)


def _objects(block: bytes) -> list[tuple[int, int, int, int]]:
    """(start, end, type, id) of the objects in a decompressed block, in order."""
    out, taken = [], 0
    for m in OBJECT.finditer(block, 5):
        s = m.start() - 5
        if s < taken or block[s] > 1 or s + 33 > len(block):  # in the 15-byte prefix just found, no flag, no id
            continue
        size = struct.unpack_from("<I", block, s + 1)[0]
        if size == struct.unpack("<I", m[2])[0] and 16 < size <= len(block) - s and block[s + 17:s + 25] == bytes(8):
            out.append((s, min(s + size + OBJ_TAIL, len(block)), struct.unpack("<H", m[1])[0],
                        struct.unpack_from("<Q", block, s + 25)[0]))
            taken = s + OBJ_TAIL
    return out


def _owner(objects: list[tuple[int, int, int, int]], anchor: int) -> tuple[int, int, int, int] | None:
    """The header, body or conversation object an anchor sits in; None for other or unparsed objects."""
    k = bisect.bisect_left(objects, (anchor,)) - 1
    if k >= 0 and objects[k][0] < anchor < objects[k][1] and objects[k][2] in (HEADER_OBJ, BODY_OBJ, CONV_OBJ):
        return objects[k]
    return None


# -- field classifiers (record.rs) ------------------------------------------------

def _latin(u: int) -> bool:
    return (u < 0x250 or 0x2000 <= u < 0x2070 or 0x20A0 <= u < 0x20D0 or u == 0x2122 or 0x2190 <= u <= 0x21FF)


def _is_binary_run(s: str) -> bool:
    off = sum(1 for c in s if not (_latin(ord(c)) or 0x2600 <= ord(c) < 0x27C0 or 0x1F000 <= ord(c) <= 0x1FAFF
                                   or ord(c) in (0xFE0F, 0x200D, 0x203C, 0x2049)))
    return len(s) >= 3 and off * 5 >= len(s)


def _is_email(s: str) -> bool:
    at = s.find("@")
    if at <= 0 or at + 1 >= len(s) or " " in s or len(s) > 254:
        return False
    if not all(c.isascii() and (c.isalnum() or c in "@._-+%") for c in s):
        return False
    domain = s[at + 1:]
    dot = domain.rfind(".")
    tld = domain[dot + 1:]
    return not domain.startswith(".") and dot > 0 and 2 <= len(tld) <= 12 and tld.isalpha()


def _is_guid(t: str) -> bool:
    g = t.strip("{}").split("-")
    return [len(x) for x in g] == [8, 4, 4, 4, 12] and all(c in "0123456789abcdefABCDEF-" for c in t.strip("{}"))


def _is_opaque(t: str) -> bool:
    if t.startswith("eyJ"):
        return True
    if "@" in t:
        local, domain = t.split("@", 1)
        if local.rsplit(".", 1)[-1] in ("png", "jpg", "jpeg", "gif", "bmp") and domain.startswith("01") \
                and all(c in "0123456789abcdefABCDEF." for c in domain):
            return True
    return len(t) >= 24 and all(c in "0123456789abcdefABCDEF" for c in t)


def _is_enum(t: str) -> bool:
    return (t.lower() in ENUM_VALUES or t.startswith(("image/", "application/", "text/", "multipart/"))
            or _is_guid(t) or _is_opaque(t))


def _is_filename(t: str) -> bool:
    return t.startswith(("~/", "/")) or "\\" in t or t.rsplit(".", 1)[-1].lower() in FILE_EXT


def _split_packed(t: str) -> list[str]:
    """Packed collections separate entries with length prefixes that decode off-script."""
    parts = "".join(c if _latin(ord(c)) else "\0" for c in t).split("\0")
    return [p.strip() for p in parts if len(p.strip()) >= 2]


def _undouble(t: str) -> str:
    n = len(t)
    return t[:n // 2] if n % 2 == 0 and n and t[:n // 2] == t[n // 2:] else t


def _is_person(t: str) -> bool:
    n = len(t)
    if not 2 <= n <= 48 or "@" in t or _is_enum(t) or ". " in t or " - http" in t:
        return False
    if "<" in t or ">" in t or t.endswith(",") or _is_filename(t):
        return False
    if SIGNATURE.fullmatch(t) or _strip_prefix(t) != t or BARE_NAME.fullmatch(t):  # a signature line, the subject
        return False  # of the user's own reply, the tail of an address
    return sum(c.isalpha() for c in t) * 2 >= n


def _is_name_list(t: str) -> bool:
    """A names field, not text: one name twice (the display name stored twice), or names packed behind a length
    each, which decodes as a character U+XX00."""
    parts = "".join("\0" if ord(c) & 0xFF == 0 and not _latin(ord(c)) else c for c in t).split("\0")
    parts = [_undouble(p.strip()) for p in parts if p.strip()]
    return bool(parts) and (len(parts) > 1 or _undouble(t.strip()) != t.strip()) and all(map(_is_person, parts))


def _is_subject_like(t: str) -> bool:
    n = len(t)
    if not 3 <= n <= 400 or _is_email(t) or _is_filename(t) or "http" in t or "mailto:" in t:
        return False
    if SIGNATURE.fullmatch(t) or t.startswith("<") and t.endswith(">"):  # a signature line, a Message-ID
        return False
    if t.count(". ") >= 2:
        return False
    return sum(c.isalpha() for c in t) * 2 >= n


def _subject_like(t: str) -> bool:
    return len(_split_packed(t)) <= 1 and _is_subject_like(t) and not _is_enum(t)


def _strip_prefix(s: str) -> str:
    t = s.strip()
    while True:
        low = t.lower()
        p = next((p for p in REPLY_PREFIXES if low.startswith(p)), None)
        if not p:
            return t
        t = t[len(p):].lstrip()


def _html_to_text(raw: str) -> str:
    raw = re.sub(r"(?is)<(script|style|head)\b.*?</\1\s*>", " ", raw)
    raw = re.sub(r"(?i)<(br|/p|/div|/tr|/li|/h[1-3]|/td)\b[^>]*>", " ", raw)
    text = html.unescape(re.sub(r"<[^>]*>", "", raw))
    return " ".join("".join(c for c in text if c.isspace() or ord(c) >= 32).split())


# -- per-record walk (schema.rs) --------------------------------------------------

def _runs(block: bytes) -> list[tuple[int, str]]:
    out = []
    for m in _RUN.finditer(block):
        raw = m.group()
        if len(raw) % 2:
            raw = raw[:-1]
        s = raw.decode("utf-16-le", "replace").rstrip("�")
        if len(s) >= 3 and not _is_binary_run(s):
            out.append((m.start(), s))
    return out


def _behind_copy(block: bytes, pos: int, text: str) -> bool:
    """Whether the run at pos directly follows a copy of its own text, as a subject and its topic do."""
    tail = text[-8:].encode("utf-16-le") + b"\x00\x00"
    return pos >= len(tail) and block[pos - len(tail):pos] == tail


def _html_span(region: bytes, pos: int = 0) -> tuple[int, int] | None:
    """Start and end of the first HTML document at or after pos. The UTF-8 HTML holds no NUL, and the field
    behind it starts with a length that does, so a document ends before the first NUL after its start."""
    marks = (b"<html", b"<!DOCTYPE", b"<body", b"<div", b"<table")
    starts = [p for p in (region.find(m, pos) for m in marks) if p != -1]
    if not starts:
        return None
    nul = region.find(b"\x00", min(starts))
    tail = region[min(starts):nul] if nul != -1 else region[min(starts):]
    ends = [p + len(m) for m in (b"</html>", b"</body>") if (p := tail.find(m)) != -1]
    end = min(ends) if ends else len(tail) if nul != -1 else _text_end(tail)
    last_close = tail.rfind(b"</", 0, end)
    if last_close != -1 and (gt := tail.find(b">", last_close, end)) != -1:
        end = gt + 1
    return min(starts), min(starts) + end


def _html_region(region: bytes) -> str:
    span = _html_span(region)
    return region[span[0]:span[1]].decode("utf-8", "replace") if span else ""


def _body_text(obj: bytes) -> str:
    """Text of every HTML document in a body object, each distinct one once. A reply or forward written in
    Outlook keeps an empty editor document, its own text and the quoted mail apart; many bodies are stored twice."""
    texts, pos = [], 0
    while span := _html_span(obj, pos):
        text = _html_to_text(obj[span[0]:span[1]].decode("utf-8", "replace"))
        if text and not any(text in t for t in texts):
            texts = [t for t in texts if t not in text] + [text]
        pos = max(span[1], span[0] + 1)
    return " ".join(texts)


def _text_end(b: bytes, win: int = 32) -> int:
    for i in range(0, len(b) - win + 1, win):
        if sum(1 for c in b[i:i + win] if 0x20 <= c < 0x7f or c in (9, 10, 13)) * 4 < win * 3:
            return i
    return len(b)


def _kind(block: bytes, anchor: int, owner: tuple | None = None, alen: int = len(ANCHOR)) -> str:
    """Three layouts sit on the same ItemClass anchor (alen bytes long).

    body:  the body object of a header object (no sender, no subject of its own); the HTML usually follows the
           anchor, but some bodies hold a card or a file path first, so the owning object type decides
    mid:   header with the Message-ID right after the anchor; sender pair before the anchor
    other: anything else, mostly conversation summaries, whose sender, preview and subject precede the anchor
    """
    nxt = block[anchor + alen + 2:anchor + alen + 18]
    if (owner[2] == BODY_OBJ) if owner else nxt.startswith(HTML_START):
        return "body"
    return "mid" if nxt[:2] == b"<\x00" else "other"


def _record(block: bytes, runs: list[tuple[int, str]], anchor: int, prev_end: int, end: int, kind: str,
            owner: tuple | None = None, item_class: str = "IPM.Note") -> dict:
    back = min(SENDER_MAX_DIST, anchor - prev_end)
    alen = 2 * len(item_class)
    rec = dict(sender=None, name=None, subject=None, preview=None, message_id=None, recipients=[], conv_ids=[],
               attachments=[], kind=kind, oid=None, received=None, meeting=MEETING_CLASSES.get(item_class),
               body_file=None, from_file=False, cn=0, summary_id=None)
    # attachment names can sit past RECORD_MAX, and in a body object; collect() hands the latter to the header
    lo = anchor if kind == "body" else anchor - back
    rec["attachments"] = sorted({t.strip() for pos, t in runs
                                 if lo <= pos < end and pos != anchor and ATTACHMENT_NAME.match(t.strip())})
    if owner and owner[0] + CN_AT + 8 <= owner[1]:
        rec["cn"] = struct.unpack_from("<Q", block, owner[0] + CN_AT)[0]  # orders the saved revisions of one item
    if owner and owner[2] in (HEADER_OBJ, BODY_OBJ):
        rec["oid"] = owner[3]  # the local id a body object and its header object share
    elif owner:
        rec["summary_id"] = owner[3]  # _raw_records drops a summary with a message's own id
    if kind == "body":  # only the body text and attachment names are used
        body = _body_text(block[anchor:owner[1]]) if owner else _html_to_text(_html_region(block[anchor:end]))
        rec["body"], rec["ticks"] = (body if len(body.split()) >= 5 else ""), set()
        f = BODY_FILE.search(block, owner[0], owner[1]) if owner else None
        path = f.group()[:-2].decode("utf-16-le") if f else ""
        rec["body_file"] = path if "/EFMData/" in path and ".." not in path else None  # _raw_records reads it
        return rec
    wide_end = min(end, anchor + RECORD_MAX)  # recipients sit in the child objects behind the header object;
    if owner and owner[2] == CONV_OBJ:  # behind a summary sit the next summaries, whose senders are no recipients
        wide_end = min(wide_end, owner[1])
    if owner:  # every other field comes from the record's own object, never from a neighbour's
        back = anchor - max(prev_end, owner[0])  # sender, subject and preview can sit ~900 bytes back
        end = min(end, owner[1])
        if owner[2] == HEADER_OBJ and owner[0] + RECEIVED_AT + 8 <= owner[1]:  # a cut-off object has no time
            t = struct.unpack_from("<Q", block, owner[0] + RECEIVED_AT)[0]
            rec["received"] = t if TICK_LO <= t < TICK_HI else None
    end = min(end, anchor + RECORD_MAX)
    fields = [(pos - anchor, t) for pos, t in runs if anchor - back <= pos < end and pos != anchor]
    before = [f for f in fields if f[0] < 0]
    after = [f for f in fields if f[0] > 0]

    # sender: nearest address walking back from the anchor, then forward; name sits beside it
    nearest = [f for f in reversed(before)] + after
    for i, (rel, text) in enumerate(nearest):
        t = text.strip()
        if _is_email(t):
            rec["sender"] = _undouble(t).lower()
            for k in (i + 1, i - 1):
                if 0 <= k < len(nearest) and abs(nearest[k][0] - rel) <= 200:
                    raw = _undouble(nearest[k][1].strip())
                    cand = (_split_packed(raw) or [raw])[0]
                    if not _is_email(cand) and _is_person(cand):
                        rec["name"] = cand
                        break
            break
    if owner and owner[2] == CONV_OBJ:  # a summary holds the From address first, then the one its list shows,
        first = next((_undouble(t.strip()).lower() for _, t in before if _is_email(t.strip())), None)
        if first and first != rec["sender"]:  # which for the user's own mail is the recipient, with its name
            rec["sender"], rec["name"] = first, None

    ids = []  # (rel, end_rel, id) of the <...@...> runs
    for rel, text in after:
        t = text.strip()
        if t.startswith("<") and t.endswith(">") and "@" in t and " " not in t:
            ids.append((rel, rel + 2 * len(text) + 2, t.strip("<>")))
    for pos, text in runs:
        t = text.strip()
        if anchor < pos < wide_end and _is_email(t):
            a = _undouble(t).lower()  # an address stored twice, or behind the item class of a summary's list
            a = next((a[len(c):] for c in CLASS_PREFIXES if a.startswith(c) and _is_email(a[len(c):])), a)
            if a != rec["sender"]:
                rec["recipients"].append(a)
    # a reply stores In-Reply-To first and its own Message-ID right behind it; a new message only its own.
    # Only the mid layout starts with an id; elsewhere the first id may be a neighbour's.
    if kind == "mid" and ids and ids[0][0] == alen + 2:
        adjacent = len(ids) > 1 and ids[1][0] - ids[0][1] <= 8
        rec["message_id"] = ids[1][2] if adjacent else ids[0][2]

    # subject: the value stored twice (NormalizedSubject + Topic), reply prefixes ignored
    ordered = [f for f in after if f[0] <= SUBJECT_MAX_REL] + list(reversed(before))
    key40 = lambda s: _strip_prefix(s)[:40]

    def same(a: str, b: str) -> bool:
        if key40(a) == key40(b):
            return True
        sa = _strip_prefix(a).strip()
        if len(sa) < 4:
            return False
        tb = (_split_packed(b) or [b])[-1]
        return tb.endswith(sa) or _strip_prefix(b).strip().endswith(sa) or key40(tb) == key40(a)

    undoubled = [(rel, _undouble(t.strip())) for rel, t in ordered]
    subject = next((t for _, t in undoubled if _subject_like(t) and sum(same(t, u) for _, u in undoubled) > 1), None)
    stale = None
    whole = Counter(t for _, t in undoubled)
    if subject and whole[subject] < 2 and owner and owner[2] == HEADER_OBJ:
        # a subject still being typed when the draft was first saved stays behind the names as the topic; the subject
        # the mail went out with extends it and is stored twice whole
        longer = next((t for t, n in whole.items() if n > 1 and len(t) > len(subject) and _subject_like(t)
                       and _strip_prefix(t).startswith(_strip_prefix(subject))), None)
        if longer:
            stale, subject = subject, longer
    if subject is None and owner and owner[2] == HEADER_OBJ:
        # the topic right behind its own copy, which the walk misses when that copy is glued to the field before it;
        # a lone run elsewhere is the preview, or the recipients' names of the user's own mail without a subject
        subject = next((t.strip() for rel, t in after
                        if _subject_like(t.strip()) and _behind_copy(block, anchor + rel, t)), None)
    elif subject is None:  # a lone subject-like run exactly where the schema puts it
        subject = next((t for rel, t in undoubled
                        if 0 <= rel <= SUBJECT_SOLO_MAX_REL and _subject_like(t) and len(_split_packed(t)) <= 1
                        and not t.endswith("Anonymous") and t[:1] and (t[0].isupper() or t[0].isdigit() or not t[0].isalpha())
                        and len(t.split()) >= 2), None)
    rec["subject"] = subject
    if rec["name"] and rec["name"] == subject:
        rec["name"] = None
    rec["conv_ids"] = [t.strip() for _, t in ordered
                       if len(t.strip()) >= 24 and all(c in "0123456789abcdefABCDEF-" for c in t.strip())]

    claimed = {rec["subject"], rec["name"], rec["sender"], stale}
    topic = _strip_prefix(subject or "")

    def preview_like(t: str) -> bool:  # not the subject with a reply prefix, not a file path, not a names field
        return (t not in claimed and not t.startswith(("<", "~/")) and not _is_enum(t) and not _is_email(t)
                and not (topic and _strip_prefix(_undouble(t)) == topic) and not _is_name_list(t)
                and not BARE_NAME.fullmatch(t))
    # in an object the preview can precede the anchor
    rec["preview"] = max((t.strip() for _, t in (fields if owner else after) if preview_like(t.strip())),
                         key=len, default=None)

    body = _html_to_text(_html_region(block[anchor:end]))
    rec["body"] = body if len(body.split()) >= 5 else ""

    # send time, when the object holds no delivery time: the earliest .NET tick in the record's own span
    span = b"" if rec["received"] else block[anchor - back:end]
    ticks = [t for (t,) in (struct.unpack_from("<Q", span, x) for x in range(max(0, len(span) - 7)))
             if TICK_LO <= t < TICK_HI]
    rec["ticks"] = set(ticks)
    return rec


def _norm_name(name: str) -> str:
    """Normalize Outlook's sanitized cache name: drop the [n] suffix and non-alphanumerics."""
    return re.sub(r"[^a-z0-9]", "", re.sub(r"\[\d+\]", "", name).lower())


def _norm_original(name: str) -> set[str]:
    """Both ways Outlook sanitizes an accented character: 'a' plus '_' (decomposed) or just '_'."""
    keep_base = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    drop_char = "".join(c for c in unicodedata.normalize("NFC", name) if c.isascii())
    return {_norm_name(keep_base), _norm_name(drop_char)}


def attachment_files() -> list[Path]:
    """Attachments New Outlook has downloaded, stored with a [n] suffix in the profile."""
    return sorted(p for p in PROFILES.glob("*/Files/S*/*/Attachments/*/*") if p.is_file())


def _cached_text(path: Path, digest: str) -> str:
    from .onedrive import _extract  # same extractors as OneDrive files
    TEXT_CACHE.mkdir(parents=True, exist_ok=True)
    cache = TEXT_CACHE / f"{digest}.txt"
    if cache.exists():
        return cache.read_text("utf-8")
    try:
        text = _extract(path)[0]
    except Exception:
        text = ""
    cache.write_text(text, "utf-8")
    return text


def _attachment_docs(merged: dict, stats: Counter) -> list[dict]:
    """One doc per distinct cached attachment, linked to the mail(s) that carried it."""
    carriers: dict[str, list[dict]] = {}
    originals: dict[str, str] = {}
    for m in merged.values():
        for a in m["attachments"]:
            for key in _norm_original(a):
                carriers.setdefault(key, []).append(m)
                originals.setdefault(key, a)
    docs, seen = [], set()
    for path in attachment_files():
        if not ATTACHMENT_NAME.match(re.sub(r"\[\d+\]", "", path.name)):
            continue  # inline images and nameless parts
        digest = hashlib.sha1(path.read_bytes()).hexdigest()
        if digest in seen:
            continue  # the same file attached to several messages
        seen.add(digest)
        key = _norm_name(path.name)
        mails = sorted(carriers.get(key, []), key=lambda m: m.get("ts") or "")
        original = originals.get(key) or re.sub(r"\[\d+\]", "", path.name)
        first = mails[0] if mails else {}
        who = (f'{first["name"]} <{first["sender"]}>' if first.get("name") and first.get("sender")
               else first.get("name") or first.get("sender"))
        in_mail = "; ".join(sorted({m.get("subject") or "(no subject)" for m in mails})) or "(mail not identified)"
        body = f"Attachment: {original}\nIn mail: {in_mail}\nFile: {path}\n\n{_cached_text(path, digest)}"
        docs.append(dict(ref=f"attachment:{digest[:16]}", ts=first.get("ts"), who=who,
                         title=f"[attachment] {original}", body=body))
        stats["attachments"] += 1
        stats["attachments_linked"] += bool(mails)
    return docs


def _send_tick(ticks: set[int]) -> int | None:
    """The earliest tick that is followed by its delivery stamp.

    hxprobe takes the plain minimum. On this store that sometimes lands on a stray
    older value (years before the mailbox existed), so require what a real send time
    always has: another stamp (delivery, receipt) shortly after it.
    """
    if not ticks:
        return None
    ordered = sorted(ticks)
    for i, t in enumerate(ordered):
        if any(0 < u - t <= PAIR_WINDOW for u in ordered[i + 1:i + 6]):
            return t
    return ordered[0] if len(ordered) == 1 else ordered[-1] if ordered[-1] - ordered[0] > 30 * PAIR_WINDOW else ordered[0]


def _norm_text(s: str | None) -> str:
    return " ".join(re.sub(r"[^\w]+", " ", (s or "").lower()).split())


def _body_file(store: Path, path: str, cache: dict, stats: Counter) -> str:
    """Text of the body file a body object names (gzip of the whole HTML); "" when it is gone or unreadable.
    Read into memory only, once per collect."""
    key = (str(store), path)
    if key not in cache:
        try:
            data = (store.parent / path.removeprefix("~/")).read_bytes()
            cache[key] = _body_text(gzip.decompress(data) if data[:2] == b"\x1f\x8b" else data)
            stats["body_files_read"] += 1
        except (OSError, EOFError, zlib.error):
            cache[key] = ""
            stats["body_files_missing"] += 1
    return cache[key]


def _meeting_info(times: tuple[int, int] | None, place: bytes | None) -> dict:
    """Start and end (local time) from a meeting item's time object; organizer and location from the place
    object of a request: a name and address pair, then the location, after an optional time zone name."""
    info = dict(start=None, end=None, organizer=None, organizer_name=None, location=None)
    if times and TICK_LO <= times[0] <= times[1] <= times[0] + 31 * PAIR_WINDOW:
        info["start"], info["end"] = (datetime.fromtimestamp((t - UNIX_EPOCH_TICKS) / 1e7).isoformat(timespec="minutes")
                                      for t in times)
    runs = [t.strip() for _, t in _runs(place or b"")]
    zone = lambda t: t.endswith(" Time") or t.startswith("tzone:")
    at = next((i for i, t in enumerate(runs) if _is_email(t)), None)
    if at is not None:
        info["organizer"] = runs[at].lower()
        if at and _is_person(runs[at - 1]) and not zone(runs[at - 1]):
            info["organizer_name"] = runs[at - 1]
        info["location"] = next((t for t in runs[at + 1:] if not _is_email(t) and not zone(t) and len(t) >= 2), None)
    return info


def _raw_records(stats: Counter) -> tuple[list[dict], set[tuple[str, int]], dict]:
    """The record of every header, body and summary object with an ItemClass anchor, the (store, id) of every body
    object, also one without an inline body, and the times and place of every meeting item by (store, id)."""
    raw, body_ids, times, places, files = [], set(), {}, {}, {}
    for store in stores():
        snapshot = store.read_bytes()  # Outlook keeps writing; parse a consistent copy in memory
        for bno, block in enumerate(_blocks(snapshot, stats)):
            objects = _objects(block)
            for s, e, typ, oid in objects:
                if typ == BODY_OBJ:
                    body_ids.add((str(store), oid))
                elif typ == TIME_OBJ and s + START_AT + 8 <= e:
                    times[(str(store), oid)] = (struct.unpack_from("<Q", block, s + START_AT)[0],
                                                struct.unpack_from("<Q", block, s + END_AT)[0])
                elif typ == PLACE_OBJ:  # parsed only for meeting items: most of these objects are contacts
                    places[(str(store), oid)] = block[s:e]
            # a meeting item counts only in its header or body object, so it always has its id: its conversation
            # summaries all repeat a cached item here, with the names of its attendees run together for a preview
            anchors = [(m.start(), m.group().decode("utf-16-le")) for m in ANCHORS.finditer(block)]
            anchors = [(a, c) for a, c in anchors
                       if c == "IPM.Note" or (o := _owner(objects, a)) and o[2] in (HEADER_OBJ, BODY_OBJ)]
            if not anchors:
                continue
            runs = _runs(block)
            done = set()
            for k, (a, item_class) in enumerate(anchors):
                prev_end = anchors[k - 1][0] + 2 * len(anchors[k - 1][1]) if k else 0
                end = anchors[k + 1][0] if k + 1 < len(anchors) else len(block)
                owner = _owner(objects, a)
                # mail is a header, body or summary object; the class also sits in action backups and search caches
                # (~/ActionsBackup/...), whose cards and names are no mail. A summary names the class of each kind
                # of item it covers, with their names packed behind the later ones: one record per object
                if owner is None or owner in done:
                    stats["anchors_not_mail" if owner is None else "anchors_repeated"] += 1
                    continue
                done.add(owner)
                r = _record(block, runs, a, prev_end, end, _kind(block, a, owner, 2 * len(item_class)), owner,
                            item_class)
                r["blk"] = (str(store), bno)
                r["link"] = (str(store), r["oid"]) if r["oid"] is not None else None
                if r["body_file"]:  # the whole body; what the object holds is its opening part, or nothing
                    text = _body_file(store, r["body_file"], files, stats)
                    if text and _norm_text(r["body"])[:60] in _norm_text(text):
                        if len(text) > len(r["body"]):
                            r["body"], r["from_file"] = text, True
                    elif text:
                        stats["body_files_mismatch"] += 1
                raw.append(r)
    # a summary that carries a message's own local id is the one Outlook writes while the user composes that
    # message: its fields are left over from earlier saves (a recipient where the sender goes, a name over the
    # signature), and the message's own header and body objects hold all of it
    own = {r["link"] for r in raw if r["link"] and r["kind"] != "body"}
    kept =[r for r in raw if not (r["summary_id"] is not None and (r["blk"][0], r["summary_id"]) in own)]
    stats["summaries_of_drafts"] += len(raw) - len(kept)
    raw = kept
    meetings = {k: _meeting_info(times.get(k), places.get(k))
                for k in {r["link"] for r in raw if r["meeting"] and r["link"]}}
    return raw, body_ids, meetings


def _attach_bodies(raw: list[dict], stats: Counter, body_ids: set[tuple[str, int]]) -> set[tuple[str, int]]:
    """A body object and its header object share one local object id, and collect() merges records by that id,
    so a body whose id a header record carries needs nothing here (body_by_id).

    Only a body that no header record claims by id falls back to the text match: the adjacent header first, then
    the one header anywhere whose preview opens this body. The fallback decides once per body object (all its
    revisions go to one header) and skips every header whose own id has a body object, so it never overrides
    the link. A meeting item's body only goes to a meeting item and a mail body only to mail. A body that fits
    neither stays an orphan without sender. Returns the ids the header records carry.
    """
    stats["body_by_id"] += 0  # always in the notes: index.replace_source spots the first build by local id by it
    headers = {r["link"] for r in raw if r["kind"] != "body" and r["link"]}
    free = lambda h: h["link"] not in body_ids  # a header with a body object of its own takes no other body
    kin = lambda h, b: bool(h["meeting"]) == bool(b["meeting"])
    index: dict[str, list[dict]] = {}
    for r in raw:
        p = _norm_text(r["preview"])
        if r["kind"] != "body" and len(p) >= 40 and free(r):
            index.setdefault(p[:40], []).append(r)
    unclaimed: dict = {}
    for i, r in enumerate(raw):
        if r["kind"] != "body" or not r["body"]:
            continue
        if r["link"] in headers:
            stats["body_by_id"] += 1
        else:
            unclaimed.setdefault(r["link"] or i, []).append(i)
    taken = {}  # header -> the body object it took: one header never takes two known body objects
    for revisions in unclaimed.values():
        i = max(revisions, key=lambda k: len(raw[k]["body"]))  # the longest revision speaks for the object
        r = raw[i]
        start = _norm_text(r["body"])[:BODY_MATCH_CHARS]
        j = i + 1  # the header follows its body object, sometimes after another revision of the body
        while j < len(raw) and raw[j]["blk"] == r["blk"] and raw[j]["kind"] == "body":
            j += 1
        nxt = raw[j] if j < len(raw) and raw[j]["blk"] == r["blk"] and free(raw[j]) and kin(raw[j], r) else None
        p = _norm_text(nxt["preview"])[:60] if nxt else ""
        target, how = None, "body_to_next_header"
        at = start.find(p) if len(p) >= 3 else -1
        if nxt is not None and at >= 0 and (len(p) >= 20 or at < 100):  # a short preview must open the body
            target = nxt
        else:  # the one header elsewhere whose preview opens this body
            hits = [(start.find(key), h) for key, hs in index.items() if key in start for h in hs if kin(h, r)]
            first = min((pos for pos, _ in hits), default=None)
            cands = [h for pos, h in hits if pos == first and h["sender"]]
            cands = [h for h in cands if h["kind"] == "mid"] or cands  # the mid layout's sender pair is structural
            if cands and len({h["sender"] for h in cands}) == 1:
                target, how = max(cands, key=lambda h: bool(h["message_id"])), "body_to_matching_header"
        if target is None or taken.setdefault(id(target), r["link"]) != r["link"]:
            stats["body_orphan"] += len(revisions)
            continue
        stats[how] += len(revisions)
        latest = raw[max(revisions, key=lambda k: (raw[k]["cn"], len(raw[k]["body"])))]["body"]  # the last save
        if len(latest) > len(target["body"]):
            target["body"] = latest
        for k in revisions:
            b = raw[k]
            target["attachments"] = sorted(set(target["attachments"]) | set(b["attachments"]))
            b["body"], b["attachments"] = "", []
    return headers


def _fold(t: dict, m: dict) -> None:
    """Merge entry m into entry t; t keeps its own values, and a body joined by id beats any other body."""
    for f in ("sender", "name", "subject", "preview", "message_id", "ts"):
        t[f] = t.get(f) or m.get(f)
    if m.get("body") and m["body_rank"] > t.get("body_rank", (False, 0)):
        t["body"], t["body_rank"] = m["body"], m["body_rank"]
    t["previews"] += [p for p in m["previews"] if p not in t["previews"]]
    t["recipients"] |= m["recipients"]
    t["conv_ids"] |= m["conv_ids"]
    t["attachments"] |= m["attachments"]


def _pick_body(m: dict) -> None:
    """The body of the last save: among the bodies joined by id (or else all of them), the one with the highest
    change number that holds the preview of the header's last save, else the one with the highest change number.
    Next to a draft's text Outlook keeps a second body object, under the same change number or the next one,
    that holds only the quoted mail or nothing; the header's preview tells the two apart."""
    bodies = m.pop("bodies")
    if not bodies:
        return
    joined = max(b[0] for b in bodies)
    cands = [b for b in bodies if b[0] == joined]
    p = _norm_text(m.get("preview"))[:60]
    best = max([b for b in cands if p and p in _norm_text(b[2])] or cands, key=lambda b: (b[1], len(b[2])))
    m["body"], m["from_file"], m["body_rank"] = best[2], best[3], (best[0], len(best[2]))  # for _fold


def _ref_key(key: tuple, m: dict) -> tuple:
    """A message without a Message-ID keeps the ref it had before local ids: (sender, time) when both exist."""
    if key[0] != "oid":
        return key
    return (m["sender"], m["ts"]) if m.get("sender") and m.get("ts") else ("oid", key[2])


def _uid(m: dict) -> str:
    """Stable short hash of what tells two copies of one Message-ID apart."""
    key = f'{m.get("sender") or ""}|{m.get("ts") or ""}|{(m.get("preview") or "")[:80]}'
    return hashlib.sha1(key.encode()).hexdigest()[:8]


def same_mail(d: dict) -> tuple:
    """Preview-only copies of one mail (its summaries, or two local copies) look alike: list them once."""
    return d.get("sender"), _norm_text(_strip_prefix(d.get("subject") or "")), _norm_text(d.get("preview"))


def _unique_refs(docs: list[dict]) -> None:
    """Two local copies of one mail, and replies without their own id, repeat a Message-ID: suffix every copy."""
    counts = Counter(d["ref"] for d in docs)
    for d in docs:
        if counts[d["ref"]] > 1:
            d["ref"] += "#" + d["uid"]
    n = Counter()
    for d in docs:  # nothing left to tell them apart: number them
        ref = d["ref"]
        n[ref] += 1
        if n[ref] > 1:
            d["ref"] = f"{ref}-{n[ref]}"


def _meeting_lines(m: dict, info: dict) -> list[str]:
    """The plain lines a meeting item's doc starts with, so its kind, times and place are searchable."""
    lines = [m["meeting"]]
    lines += [f"{label}: {info[f].replace('T', ' ')}" for label, f in (("Start", "start"), ("End", "end"))
              if info.get(f)]
    if info.get("location"):
        lines.append("Location: " + info["location"])
    org = info.get("organizer")
    if org and org != m.get("sender"):  # the sender is the organizer, unless the request was forwarded
        lines.append("Organizer: " + (f'{info["organizer_name"]} <{org}>' if info.get("organizer_name") else org))
    return lines


def collect() -> tuple[list[dict], str]:
    stats = Counter()
    raw, body_ids, meetings = _raw_records(stats)
    stats["records_raw"] = len(raw)

    for r in raw:  # a header object's delivery time; elsewhere the earliest tick with a delivery stamp
        t = r["received"] or _send_tick(r["ticks"])
        r["ts"] = datetime.fromtimestamp((t - UNIX_EPOCH_TICKS) / 1e7).isoformat(timespec="minutes") if t else None
    headers = _attach_bodies(raw, stats, body_ids)

    # one message is stored in several synced revisions and copies: merge them field by field, by local id
    merged: dict[tuple, dict] = {}
    for r in raw:
        if not (r["sender"] or r["preview"] or r["body"] or (r["attachments"] and r["link"] in headers)):
            stats["skipped_empty"] += 1
            continue
        if r["link"]:  # a header or body object: every revision and copy carries the message's local id
            key = ("oid",) + r["link"]
        elif r["sender"] and r["ts"]:
            key = (r["sender"], r["ts"])
        elif r["kind"] == "body":  # a body no header claimed
            key = ("orphan", hashlib.sha1(r["body"].encode()).hexdigest())
        else:
            key = ("id", r["message_id"] or (r["preview"] or "")[:80])
        m = merged.setdefault(key, dict(recipients=set(), conv_ids=set(), attachments=set(), previews=[], bodies=[],
                                        ids=Counter(), votes={f: {} for f in VOTED}, latest=0))
        m["meeting"] = m.get("meeting") or r["meeting"]
        m["attachments"].update(r["attachments"])
        # a draft is saved many times, each save under a higher change number, and only the copies of the last
        # save vote; the delivery time is no order, a later save of a draft can carry an earlier one
        if r["kind"] != "body":
            m["latest"] = max(m["latest"], r["cn"])
        for f, votes in m["votes"].items():
            if r[f]:
                votes.setdefault(r["cn"], Counter())[r[f]] += 1
        if r["message_id"]:
            m["ids"][r["message_id"]] += 1
        if r["preview"]:
            m["previews"].append((r["cn"], r["preview"]))
        if r["body"]:
            m["bodies"].append((r["kind"] == "body" and r["link"] in headers, r["cn"], r["body"], r["from_file"]))
        m["recipients"].update(r["recipients"])
        m["conv_ids"].update(r["conv_ids"])
    for key, m in merged.items():  # the value most copies agree on
        m["message_id"] = m["ids"].most_common(1)[0][0] if m["ids"] else None
        top = m.pop("latest")
        for f, votes in m.pop("votes").items():
            cns = [cn for cn in votes if cn == top or f in FILLED]
            m[f] = votes[max(cns)].most_common(1)[0][0] if cns else None
        # every preview stays searchable, but of a message's own saves only the last: an earlier save of a draft
        # holds earlier text, or the recipients' names where the preview goes
        keep = [p for cn, p in m["previews"] if key[0] != "oid" or cn == top or p == m.get("preview")]
        m["previews"] = list(dict.fromkeys(keep))
        _pick_body(m)

    # a record outside a header object is mostly a conversation summary, which repeats its latest message:
    # fold it into that message, found by the summary's preview among the same sender's messages (then the same
    # time), then inside their text (then the same time), then among all senders when it is long and found whole
    # (notifications of one template share their opening); one that matches no message is the only trace of it
    # and stays. Every preview of the summary is tried, the last save's first. A summary without a preview folds
    # into the one message of its sender at its time. A meeting item's summary folds only into a meeting item,
    # mail's only into mail
    entries = [(k, m.get("sender"), m.get("ts"), {_norm_text(t)[:60] for t in m["previews"]},
                " ".join(_norm_text(t) for t in [m.get("body")] + m["previews"]), bool(m["meeting"]))
               for k, m in merged.items() if k[0] == "oid"]
    for key in [k for k in merged if k[0] != "oid"]:
        m = merged[key]
        kin = [e for e in entries if e[5] == bool(m["meeting"])]
        same = [e for e in kin if e[1] == m.get("sender")]
        tiers = [[e for e in same if m.get("sender") and m.get("ts") and e[2] == m["ts"]]] if not m["previews"] else []
        # the last save's preview first: a torn save can hold a stray one, and an earlier save the real one
        for whole in dict.fromkeys(_norm_text(t) for t in [m.get("preview")] + m["previews"] if t):
            p = whole[:60]
            exact = [e for e in same if p in e[3]]
            inside = [e for e in same if len(p) >= 20 and p in e[4]]
            tiers += [exact, [e for e in exact if e[2] == m.get("ts")], inside,
                      [e for e in inside if e[2] == m.get("ts")],
                      [e for e in kin if len(p) >= 40 and p in e[3] and whole in e[4]]]
        hit = next((t[0][0] for t in tiers if len(t) == 1), None)
        if hit:
            _fold(merged[hit], merged.pop(key))
            stats["folded_summaries"] += 1

    # a revision without sender or time lands under an ("id", ...) key; fold it into the
    # dated copy of the same message when exactly one such copy exists
    by_mid: dict[str, list[tuple]] = {}
    for key, m in merged.items():
        if key[0] != "id" and m.get("message_id"):
            by_mid.setdefault(m["message_id"], []).append(key)
    for key in [k for k in merged if k[0] == "id"]:
        m = merged[key]
        targets = by_mid.get(m.get("message_id") or "", [])
        if len(targets) == 1:
            _fold(merged[targets[0]], merged.pop(key))
            stats["folded_undated"] += 1

    # Topic is written once per conversation: back-fill subjects across a thread
    by_conv = {c: m["subject"] for m in merged.values() if m.get("subject") for c in m["conv_ids"]}
    for m in merged.values():
        if not m.get("subject"):
            m["subject"] = next((by_conv[c] for c in m["conv_ids"] if c in by_conv), None)
            stats["subject_inherited"] += bool(m["subject"])

    docs = []
    for key, m in merged.items():
        who = f'{m["name"]} <{m["sender"]}>' if m.get("name") and m.get("sender") else m.get("name") or m.get("sender")
        text = m.get("body") or m.get("preview") or ""
        info = meetings.get(key[1:], {}) if key[0] == "oid" else {}
        body_parts, seen = [m.get("subject") or "", text], _norm_text(text)
        if m["meeting"]:
            body_parts[:0] = _meeting_lines(m, info)
        for p in m["previews"]:  # keep every preview searchable, also when the body does not contain it
            if _norm_text(p)[:60] not in seen:
                body_parts.append("Preview: " + p)
                seen += " " + _norm_text(p)
        if m["recipients"]:
            body_parts.append("To: " + ", ".join(sorted(m["recipients"])))
        if m["attachments"]:
            body_parts.append("Attachments: " + "; ".join(sorted(m["attachments"])))
        docs.append(dict(ref=f"outlook:{m.get('message_id') or _ref_key(key, m)}", ts=m.get("ts") or info.get("start"),
                         who=who, title=m.get("subject") or "(no subject recovered)", body="\n".join(body_parts),
                         sender=m.get("sender"), name=m.get("name"), subject=m.get("subject"),
                         preview=m.get("preview"), has_body=bool(m.get("body")), body_ref=None, uid=_uid(m),
                         meeting=m["meeting"], body_cached=key[0] == "oid" and key[1:] in body_ids))
        m["doc"] = docs[-1]
    _unique_refs(docs)

    # a copy of a Message-ID that has only the preview of a body another copy holds: point to it
    copies: dict[str, list[dict]] = {}
    for m in merged.values():
        if m.get("message_id"):
            copies.setdefault(m["message_id"], []).append(m)
    for group in copies.values():
        for m in group:
            p = _norm_text(m.get("preview"))[:60]
            if p and not m.get("body"):
                m["doc"]["body_ref"] = next((c["doc"]["ref"] for c in group if p in _norm_text(c.get("body"))), None)
    mails = [m for m in merged.values() if not m["meeting"]]
    stats["messages"] = len(mails)
    # always in the notes: index.replace_source spots the first build with meeting items by it
    stats["meetings"] = len(merged) - len(mails)
    stats["meetings_with_body"] = sum(1 for m in merged.values() if m["meeting"] and m.get("body"))
    stats["meetings_with_time"] = sum(1 for k, m in merged.items()
                                      if m["meeting"] and meetings.get(k[1:], {}).get("start"))
    stats["body_files_linked"] = sum(1 for m in merged.values() if m.get("from_file"))
    att_docs = _attachment_docs(merged, stats)
    docs += att_docs
    stats["with_subject"] = sum(1 for m in mails if m.get("subject"))
    stats["with_sender"] = sum(1 for m in mails if m.get("sender"))
    stats["with_time"] = sum(1 for m in mails if m.get("ts"))
    stats["with_full_body"] = sum(1 for m in mails if m.get("body"))
    stats["preview_copies"] = sum(1 for d in docs if d.get("body_ref"))
    stats["preview_only"] = len({same_mail(d) for d in docs if to_open(d)})
    return docs, ", ".join(f"{k}={v}" for k, v in sorted(stats.items()))


def to_open(d: dict) -> bool:
    """A mail to open in Outlook for its body: none here, none in another copy, and no body object cached either
    (one that is empty or too short to keep has nothing more to fetch). A record without a sender is not one:
    nothing places it in Outlook, so it stays searchable but is not listed."""
    return (d.get("has_body") is False and not d["body_ref"] and not d["meeting"] and not d["body_cached"]
            and bool(d["sender"]))


def last_change() -> float:
    return max((s.stat().st_mtime for s in stores()), default=0.0)
