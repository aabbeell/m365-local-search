"""Teams meeting transcripts.

Two kinds of documents:
- Transcript files downloaded from Teams (Recap > Transcript > Download) into TRANSCRIPTS_DIR
  (default ~/Documents/auchan-transcripts): .vtt, .docx, .txt, .md. Full text, speaker per line.
- "[transcript available]" markers from the Teams cache: one per recorded meeting that has a
  transcript in the organizer's OneDrive, so it is visible which ones are still to be downloaded.
"""
import html
import os
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote

from . import onedrive, teams

TRANSCRIPTS_DIR = Path(os.environ.get("M365SEARCH_TRANSCRIPTS", Path.home() / "Documents/auchan-transcripts"))
TYPES = {".vtt", ".docx", ".txt", ".md"}

CUE_TIME = re.compile(r"^(\d{2}:)?\d{2}:\d{2}[.,]\d{3} --> ")
VOICE = re.compile(r"<v\s+([^>]+)>(.*?)(?:</v>|$)", re.S)
DURATION = re.compile(r'duration="(\d+):(\d{2}):(\d{2})')
ORGANIZER = re.compile(r"/personal/([a-z0-9_]+?)_(?:partner_)?(?:auchan_hu|p2m_hu|bpdata_com)/", re.I)


def _vtt(text: str) -> str:
    """Speaker-per-line text from a Teams WebVTT; consecutive lines of one speaker are merged."""
    lines, last = [], None
    for block in re.split(r"\n\s*\n", text.replace("\r", "")):
        rows = [r for r in block.split("\n") if r.strip()]
        if not rows or rows[0].startswith("WEBVTT") or not any(CUE_TIME.match(r) for r in rows):
            continue
        cue = " ".join(r for r in rows if not CUE_TIME.match(r) and not re.fullmatch(r"[\w-]+/\d+-\d+|\d+", r.strip()))
        m = VOICE.search(cue)
        who, said = (m.group(1).strip(), m.group(2)) if m else (None, cue)
        said = html.unescape(re.sub(r"<[^>]+>", "", said)).strip()
        if not said:
            continue
        if who == last and lines:
            lines[-1] += " " + said
        else:
            lines.append(f"{who}: {said}" if who else said)
            last = who
    return "\n".join(lines)


def _files() -> tuple[list[dict], dict]:
    docs, stats = [], dict(files=0, file_errors=0)
    if not TRANSCRIPTS_DIR.is_dir():
        return docs, stats
    for p in sorted(TRANSCRIPTS_DIR.rglob("*")):
        if not p.is_file() or p.name.startswith(".") or p.suffix.lower() not in TYPES:
            continue
        try:
            if p.suffix.lower() == ".vtt":
                body, who = _vtt(p.read_text("utf-8", errors="replace")), None
            else:
                body, who = onedrive._extract(p)
        except Exception as e:  # one bad file must not stop the build
            stats["file_errors"] += 1
            print(f"  ! {p.name}: {type(e).__name__}: {e}")
            continue
        speakers = sorted({l.split(": ", 1)[0] for l in body.splitlines() if ": " in l[:60]})
        docs.append(dict(
            ref=str(p),
            ts=datetime.fromtimestamp(p.stat().st_mtime).isoformat(timespec="minutes"),
            who=", ".join(speakers[:8]) or who,
            title=f"[transcript] {p.stem}",
            body=body,
        ))
        stats["files"] += 1
    return docs, stats


def _markers() -> tuple[list[dict], dict]:
    """One marker per meeting recording whose transcript was saved to OneDrive, from the Teams cache."""
    topics, best = {}, {}
    for ldb, blob in teams.stores():
        w = teams.idb.WrappedIndexDB(ldb, blob)
        for conv in teams._records(w, "Teams:conversation-manager", "conversations"):
            tp = teams._v(conv.get("threadProperties")) or {}
            if teams._v(tp.get("topic")):
                topics[teams._v(conv.get("id"))] = teams._v(tp.get("topic"))
        for store in ("replychains", "replychains-2"):
            for chain in teams._records(w, "Teams:replychain-manager", store):
                for msg in (teams._v(chain.get("messageMap")) or {}).values():
                    content = str(teams._v(msg.get("content")) or "")
                    if "onedriveForBusinessTranscript" not in content:
                        continue
                    d = DURATION.search(content)
                    secs = int(d.group(1)) * 3600 + int(d.group(2)) * 60 + int(d.group(3)) if d else 0
                    if secs == 0:
                        continue
                    conv_id = teams._v(msg.get("conversationId")) or ""
                    ts = teams._ts(teams._v(msg.get("originalArrivalTime")))
                    key = (conv_id, (ts or "")[:10])  # one recording per meeting chat per day
                    if key not in best or (ts or "") > (best[key]["ts"] or ""):
                        org = ORGANIZER.search(unquote(content))
                        link = next((u for u in re.findall(r'https://[^"\s<>]+', content) if "/:v:/" in u), None)
                        best[key] = dict(conv_id=conv_id, ts=ts, secs=secs,
                                         organizer=org.group(1).replace("_", ".") if org else None, link=link)
    docs = []
    for (conv_id, day), m in best.items():
        topic = topics.get(conv_id) or conv_id
        h, rest = divmod(m["secs"], 3600)
        docs.append(dict(
            ref=f"transcript-marker:{conv_id}:{day}",
            ts=m["ts"],
            who=m["organizer"],
            title=f"[transcript available] {topic}",
            body=(f"Meeting: {topic}\nRecorded: {m['ts']}\nLength: {h}:{rest // 60:02d}\n"
                  f"Stored in the OneDrive of: {m['organizer'] or '-'}\nRecording: {m['link'] or '-'}\n"
                  f"Text not on this Mac yet: in Teams open the meeting, Recap > Transcript > Download, "
                  f"save to {TRANSCRIPTS_DIR}"),
        ))
    return docs, dict(markers=len(docs))


def collect() -> tuple[list[dict], str]:
    files, s1 = _files()
    markers, s2 = _markers()
    return files + markers, ", ".join(f"{k}={v}" for k, v in {**s1, **s2}.items())


def last_change() -> float:
    latest = teams.last_change()
    if TRANSCRIPTS_DIR.is_dir():
        for p in [TRANSCRIPTS_DIR, *TRANSCRIPTS_DIR.rglob("*")]:
            try:
                latest = max(latest, p.stat().st_mtime)
            except OSError:
                pass
    return latest
