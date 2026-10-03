"""New Outlook for Mac local store (HxStore.hxd), read from an in-memory snapshot.

A Python port of the field walk in securized/hxstore-reverse-engineering (hxprobe,
record.rs + schema.rs, MIT). The file is a series of CRC-checked blocks holding LZ4
payloads. Mail records anchor on the UTF-16LE ItemClass "IPM.Note"; around it sit
NUL-terminated UTF-16LE fields whose byte offsets drift but whose order is stable, so
fields are claimed by position in the run sequence plus a type check, never by offset.

The store is a cache: most messages keep a ~255 character preview, and only messages
opened in the client keep their full HTML body.
"""
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

PROFILES = Path.home() / "Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles"
TEXT_CACHE = Path.home() / "Library/Application Support/m365search/attachment_text"
ATTACHMENT_NAME = re.compile(r"^[^/\\<>@]{1,200}\.(pdf|docx|xlsx|pptx|doc|xls|ppt|csv|txt|eml|msg|zip)$", re.I)
BLOCK_MAGIC = bytes.fromhex("056a703b6445025d")
ANCHOR = "IPM.Note".encode("utf-16-le")

UNIX_EPOCH_TICKS = 621_355_968_000_000_000
TICK_LO = UNIX_EPOCH_TICKS + int(datetime(2015, 1, 1).timestamp()) * 10_000_000
TICK_HI = UNIX_EPOCH_TICKS + int((datetime.now() + timedelta(days=1)).timestamp()) * 10_000_000

PAIR_WINDOW = 24 * 3600 * 10_000_000  # one day in ticks: send -> delivery
SENDER_MAX_DIST = 320     # the sender pair sits within ~120 bytes; beyond this is the previous record
SUBJECT_MAX_REL = 4096
SUBJECT_SOLO_MAX_REL = 2500

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
                                   "Focused", "Other", "Mailbox", "OneOff", "SMTP", "Personal", "Confidential")}
FILE_EXT = {"csv", "pdf", "dat", "png", "jpg", "jpeg", "gif", "xlsx", "docx", "zip", "eml"}
REPLY_PREFIXES = ("re:", "fw:", "fwd:", "re :", "aw:", "tr:", "vs:", "vá:", "továbbítás:")


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
    return sum(c.isalpha() for c in t) * 2 >= n


def _is_subject_like(t: str) -> bool:
    n = len(t)
    if not 3 <= n <= 400 or _is_email(t) or _is_filename(t) or "http" in t or "mailto:" in t:
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


def _html_region(region: bytes) -> str:
    starts = [p for p in (region.find(m) for m in (b"<html", b"<!DOCTYPE", b"<body", b"<div", b"<table")) if p != -1]
    if not starts:
        return ""
    tail = region[min(starts):]
    ends = [p + len(m) for m in (b"</html>", b"</body>") if (p := tail.find(m)) != -1]
    end = min(ends) if ends else _text_end(tail)
    last_close = tail.rfind(b"</", 0, end)
    if last_close != -1 and (gt := tail.find(b">", last_close, end)) != -1:
        end = gt + 1
    return tail[:end].decode("utf-8", "replace")


def _text_end(b: bytes, win: int = 32) -> int:
    for i in range(0, len(b) - win + 1, win):
        if sum(1 for c in b[i:i + win] if 0x20 <= c < 0x7f or c in (9, 10, 13)) * 4 < win * 3:
            return i
    return len(b)


def _record(block: bytes, runs: list[tuple[int, str]], anchor: int, prev_end: int, end: int) -> dict:
    back = min(SENDER_MAX_DIST, anchor - prev_end)
    fields = [(pos - anchor, t) for pos, t in runs if anchor - back <= pos < end and pos != anchor]
    before = [f for f in fields if f[0] < 0]
    after = [f for f in fields if f[0] > 0]
    rec = dict(sender=None, name=None, subject=None, preview=None, message_id=None, recipients=[], conv_ids=[],
               attachments=[])

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

    for rel, text in after:
        t = text.strip()
        if t.startswith("<") and t.endswith(">") and "@" in t and " " not in t and not rec["message_id"]:
            rec["message_id"] = t.strip("<>")
        elif _is_email(t) and t.lower() != rec["sender"]:
            rec["recipients"].append(t.lower())

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
    if subject is None:  # a lone subject-like run exactly where the schema puts it
        subject = next((t for rel, t in undoubled
                        if 0 <= rel <= SUBJECT_SOLO_MAX_REL and _subject_like(t) and len(_split_packed(t)) <= 1
                        and not t.endswith("Anonymous") and t[:1] and (t[0].isupper() or t[0].isdigit() or not t[0].isalpha())
                        and len(t.split()) >= 2), None)
    rec["subject"] = subject
    if rec["name"] and rec["name"] == subject:
        rec["name"] = None
    rec["conv_ids"] = [t.strip() for _, t in ordered
                       if len(t.strip()) >= 24 and all(c in "0123456789abcdefABCDEF-" for c in t.strip())]

    claimed = {rec["subject"], rec["name"], rec["sender"]}
    rec["attachments"] = sorted({t.strip() for _, t in fields if ATTACHMENT_NAME.match(t.strip())})
    rec["preview"] = max((t.strip() for _, t in after if t.strip() not in claimed and not t.lstrip().startswith("<")),
                         key=len, default=None)

    body = _html_to_text(_html_region(block[anchor:end]))
    rec["body"] = body if len(body.split()) >= 5 else ""

    # send time: the earliest .NET tick in the record's own span
    span = block[anchor - back:end]
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


def collect() -> tuple[list[dict], str]:
    stats = Counter()
    raw = []
    for store in stores():
        snapshot = store.read_bytes()  # Outlook keeps writing; parse a consistent copy in memory
        for block in _blocks(snapshot, stats):
            anchors = [m.start() for m in re.finditer(re.escape(ANCHOR), block)]
            if not anchors:
                continue
            runs = _runs(block)
            for k, a in enumerate(anchors):
                prev_end = anchors[k - 1] + len(ANCHOR) if k else 0
                end = anchors[k + 1] if k + 1 < len(anchors) else len(block)
                raw.append(_record(block, runs, a, prev_end, end))
    stats["records_raw"] = len(raw)

    for r in raw:
        t = _send_tick(r["ticks"])
        r["ts"] = datetime.fromtimestamp((t - UNIX_EPOCH_TICKS) / 1e7).isoformat(timespec="minutes") if t else None

    # one message is stored in several synced revisions: merge them field by field
    merged: dict[tuple, dict] = {}
    for r in raw:
        if not (r["sender"] or r["preview"] or r["body"]):
            stats["skipped_empty"] += 1
            continue
        key = (r["sender"], r["ts"]) if r["sender"] and r["ts"] else ("id", r["message_id"] or (r["preview"] or "")[:80])
        m = merged.setdefault(key, dict(recipients=set(), conv_ids=set(), attachments=set()))
        m["attachments"].update(r["attachments"])
        for f in ("sender", "name", "subject", "preview", "message_id", "ts"):
            m[f] = m.get(f) or r[f]
        if len(r["body"]) > len(m.get("body") or ""):
            m["body"] = r["body"]
        m["recipients"].update(r["recipients"])
        m["conv_ids"].update(r["conv_ids"])

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
            t = merged[targets[0]]
            for f in ("sender", "name", "subject", "preview", "ts"):
                t[f] = t.get(f) or m.get(f)
            if len(m.get("body") or "") > len(t.get("body") or ""):
                t["body"] = m["body"]
            t["recipients"] |= m["recipients"]
            t["conv_ids"] |= m["conv_ids"]
            t["attachments"] |= m["attachments"]
            del merged[key]
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
        body_parts = [m.get("subject") or "", m.get("body") or m.get("preview") or ""]
        if m["recipients"]:
            body_parts.append("To: " + ", ".join(sorted(m["recipients"])))
        if m["attachments"]:
            body_parts.append("Attachments: " + "; ".join(sorted(m["attachments"])))
        docs.append(dict(ref=f"outlook:{m.get('message_id') or key}", ts=m.get("ts"), who=who,
                         title=m.get("subject") or "(no subject recovered)", body="\n".join(body_parts)))
    stats["messages"] = len(docs)
    att_docs = _attachment_docs(merged, stats)
    docs += att_docs
    stats["with_subject"] = sum(1 for m in merged.values() if m.get("subject"))
    stats["with_sender"] = sum(1 for m in merged.values() if m.get("sender"))
    stats["with_time"] = sum(1 for m in merged.values() if m.get("ts"))
    stats["with_full_body"] = sum(1 for m in merged.values() if m.get("body"))
    return docs, ", ".join(f"{k}={v}" for k, v in sorted(stats.items()))


def last_change() -> float:
    return max((s.stat().st_mtime for s in stores()), default=0.0)
