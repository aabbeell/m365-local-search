"""Teams meeting transcripts, AI recaps and Loop meeting notes.

Four kinds of documents:
- Transcript files downloaded from Teams (Recap > Transcript > Download) into TRANSCRIPTS_DIR
  (default ~/Documents/auchan-transcripts): .vtt, .docx, .txt, .md. Full text, speaker per line.
- "[transcript available]" markers from the Teams cache: one per recorded meeting that has a
  transcript in the organizer's OneDrive, so it is visible which ones are still to be downloaded.
  The transcript text itself is never cached by Teams; only the speaker timeline is.
- "[meeting recap]" from the Teams meeting-recap cache (Recap tab, AI notes): the Copilot summary,
  the topic sections with their detailed notes, and the AI action items. Cached once the Recap was
  opened in the Teams app on this Mac.
- "[meeting notes]" from the Loop (Fluid) cache: the facilitator / meeting-notes Loop page
  (decisions, open questions, agenda, notes, follow-up tasks), rebuilt from its edit operations.
"""
import html
import os
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote

from . import fluid, onedrive, teams

TRANSCRIPTS_DIR = Path(os.environ.get("M365SEARCH_TRANSCRIPTS", Path.home() / "Documents/auchan-transcripts"))
TYPES = {".vtt", ".docx", ".txt", ".md"}

CUE_TIME = re.compile(r"^(\d{2}:)?\d{2}:\d{2}[.,]\d{3} --> ")
VOICE = re.compile(r"<v\s+([^>]+)>(.*?)(?:</v>|$)", re.S)
DURATION = re.compile(r'duration="(\d+):(\d{2}):(\d{2})')
CALL_ID = re.compile(r'<callId>([0-9a-f-]{36})</callId>|type="callId" value="([0-9a-f-]{36})"')
CITATION = re.compile(r"\s*cite\S*turn\d+\w+\d+\S*")
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


def _local(utc: str | None) -> str | None:
    if not utc:
        return None
    try:
        return datetime.fromisoformat(utc.replace("Z", "+00:00")).astimezone().replace(tzinfo=None) \
            .isoformat(timespec="minutes")
    except ValueError:
        return None


def _recap_text(d: dict, gpt: str | None) -> str:
    out = []
    if gpt:
        out += ["## Copilot summary", re.sub(r"</?Person>", "", CITATION.sub("", gpt)).strip(), ""]
    sections = d.get("meetingSummary") or []
    if sections:
        out.append("## Meeting notes")
    for sec in sections:
        out.append(f"### {sec.get('summaryTitle') or ''} ({sec.get('startUtteranceOffsetTimespan') or ''})")
        out.append(sec.get("summary") or "")
        for det in sec.get("detailedSummaries") or []:
            out.append(f"- {det.get('topic') or ''}: {det.get('text') or ''}")
        out.append("")
    items = d.get("actionItems") or []
    if items:
        out.append("## Action items")
    for it in items:
        owner = f" ({it.get('ownerName')})" if it.get("ownerName") else ""
        out.append(f"- {it.get('actionItemTitle') or ''}: {it.get('displayCleanContent') or ''}{owner}")
    return "\n".join(out).strip()


def _teams_meetings() -> tuple[list[dict], dict]:
    """Transcript markers, AI recaps and Loop meeting notes, all from the Teams caches."""
    topics, best, call_conv, agenda_conv, recaps, gpt = {}, {}, {}, {}, {}, {}
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
                    conv_id = teams._v(msg.get("conversationId")) or ""
                    for m in CALL_ID.finditer(content):
                        call_conv.setdefault(m.group(1) or m.group(2), conv_id)
                    if "Facilitator" in str(teams._v(msg.get("properties")) or "") and "<li>" in content:
                        for item in re.findall(r"<li>(.*?)</li>", content):
                            agenda_conv[re.sub(r"\s*\([^)]*\)\s*$", "", html.unescape(item)).strip().lower()] = conv_id
                    if "onedriveForBusinessTranscript" not in content:
                        continue
                    d = DURATION.search(content)
                    secs = int(d.group(1)) * 3600 + int(d.group(2)) * 60 + int(d.group(3)) if d else 0
                    if secs == 0:
                        continue
                    ts = teams._ts(teams._v(msg.get("originalArrivalTime")))
                    key = (conv_id, (ts or "")[:10])  # one recording per meeting chat per day
                    if key not in best or (ts or "") > (best[key]["ts"] or ""):
                        org = ORGANIZER.search(unquote(content))
                        link = next((u for u in re.findall(r'https://[^"\s<>]+', content) if "/:v:/" in u), None)
                        best[key] = dict(conv_id=conv_id, ts=ts, secs=secs,
                                         organizer=org.group(1).replace("_", ".") if org else None, link=link)
        for d in w.database_ids:
            if not d.name.startswith("Teams:meeting-recap-manager"):
                continue
            db = w[d.dbid_no]
            for r in db["meeting-recap-catchup"].iterate_records(errors_to_stdout=False):
                v = teams._v(r.value)
                if isinstance(v, dict) and isinstance(teams._v(v.get("data")), dict):
                    recaps[teams._v(v.get("recapId"))] = _plain(v["data"])
            if "meeting-recap-gpt-content" in db.object_store_names:
                for r in db["meeting-recap-gpt-content"].iterate_records(errors_to_stdout=False):
                    v = teams._v(r.value)
                    if isinstance(v, dict):
                        g = (_plain(v.get("data")) or {}).get("gptContent")
                        if g:  # keyed "<recapId>_<hash>"
                            gpt[str(teams._v(v.get("vroomId_template"))).rsplit("_", 1)[0]] = g
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
                  f"Teams does not cache transcript text on the Mac. The AI recap is a separate "
                  f"[meeting recap] hit once the Recap tab was opened in Teams. For the verbatim text: "
                  f"Recap > Transcript > Download, save to {TRANSCRIPTS_DIR}"),
        ))
    n_markers = len(docs)
    for rid, d in recaps.items():
        conv_id = call_conv.get(d.get("callId") or "")
        topic = topics.get(conv_id) or "Teams meeting"
        start, end = _local(d.get("meetingStartTime")), _local(d.get("meetingEndTime"))
        speakers = [s.get("displayName") for s in d.get("speakersSegmentations") or [] if s.get("displayName")]
        body = _recap_text(d, gpt.get(rid))
        if not body:
            continue
        docs.append(dict(
            ref=f"recap:{rid}",
            ts=start,
            who=", ".join(speakers[:10]) or None,
            title=f"[meeting recap] {topic}",
            body=f"Meeting: {topic}\nTime: {start} - {(end or '')[11:]}\nSpeakers: {', '.join(speakers)}\n\n{body}",
        ))
    n_recaps = len(docs) - n_markers
    for doc in fluid.documents():
        text = doc["text"]
        conv_id = next((c for item, c in agenda_conv.items() if item and item in text.lower()), None)
        first = re.search(r"Napirend\n(.+)|Agenda\n(.+)", text)
        topic = topics.get(conv_id) or (re.sub(r"\s*\([^)]*\)\s*$", "", (first.group(1) or first.group(2))) if first else "Loop page")
        docs.append(dict(
            ref=f"loop:{doc['id']}",
            ts=datetime.fromtimestamp(doc["first_ms"] / 1000).isoformat(timespec="minutes"),
            who=None,
            title=f"[meeting notes] {topic}",
            body=f"Meeting: {topic}\nLoop meeting notes, cached by Teams on this Mac.\n\n{text}",
        ))
    return docs, dict(markers=n_markers, recaps=n_recaps, loop_notes=len(docs) - n_markers - n_recaps)


def _plain(x):
    x = teams._v(x)
    if isinstance(x, dict):
        return {k: _plain(v) for k, v in x.items()}
    if isinstance(x, list):
        return [_plain(v) for v in x]
    return x


def collect() -> tuple[list[dict], str]:
    files, s1 = _files()
    markers, s2 = _teams_meetings()
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
