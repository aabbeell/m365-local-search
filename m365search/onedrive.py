"""OneDrive / SharePoint libraries synced under ~/Library/CloudStorage/OneDrive-*.

Every file is indexed by name. Content is extracted only from files already on disk,
unless hydrate=True, because reading a cloud-only file makes OneDrive download it.
"""
import email
import html
import os
import re
import zipfile
from datetime import datetime
from email import policy
from pathlib import Path

OFFICE = {".docx", ".xlsx", ".pptx"}
PLAIN = {".txt", ".md", ".csv", ".json", ".vtt", ".xml", ".html", ".htm"}
CONTENT_TYPES = OFFICE | PLAIN | {".pdf", ".eml"}

TAG = re.compile(r"<[^>]+>")
VTT_TIMING = re.compile(r"^\d{2}:\d{2}[:.\d]* --> .*$|^WEBVTT.*$|^\d+$", re.M)


def roots() -> list[Path]:
    return sorted(p for p in (Path.home() / "Library/CloudStorage").glob("OneDrive-*") if p.is_dir())


def _xml_text(raw: str) -> str:
    raw = re.sub(r"</(w:p|a:p|si|row)>", "\n", raw)
    raw = re.sub(r"<(w:tab|w:br)\b[^>]*/>", " ", raw)
    return html.unescape(TAG.sub("", raw))


def _office(path: Path) -> tuple[str, str | None]:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        ext = path.suffix.lower()
        if ext == ".docx":
            parts = [n for n in names if re.match(r"word/(document|header\d*|footer\d*|footnotes|comments)\.xml$", n)]
        elif ext == ".xlsx":
            parts = [n for n in names if n in ("xl/sharedStrings.xml", "xl/workbook.xml")]
        else:
            parts = sorted(n for n in names if re.match(r"ppt/(slides/slide|notesSlides/notesSlide)\d+\.xml$", n))
        text = "\n".join(_xml_text(z.read(n).decode("utf-8", "replace")) for n in parts)
        if ext == ".xlsx":  # inline strings live in the sheets themselves
            for n in names:
                if re.match(r"xl/worksheets/sheet\d+\.xml$", n):
                    raw = z.read(n).decode("utf-8", "replace")
                    text += "\n" + "\n".join(_xml_text(m) for m in re.findall(r"<is>.*?</is>", raw))
        who = None
        if "docProps/core.xml" in names:
            core = z.read("docProps/core.xml").decode("utf-8", "replace")
            m = re.search(r"<dc:creator>([^<]+)<", core) or re.search(r"<cp:lastModifiedBy>([^<]+)<", core)
            who = html.unescape(m.group(1)) if m else None
    return text, who


def _extract(path: Path) -> tuple[str, str | None]:
    ext = path.suffix.lower()
    if ext in OFFICE:
        return _office(path)
    if ext == ".pdf":
        from pypdf import PdfReader
        return "\n".join(page.extract_text() or "" for page in PdfReader(path).pages), None
    if ext == ".eml":
        msg = email.message_from_bytes(path.read_bytes(), policy=policy.default)
        body = msg.get_body(preferencelist=("plain", "html"))
        text = body.get_content() if body else ""
        return f"{msg.get('subject', '')}\n{TAG.sub(' ', text)}", msg.get("from")
    text = path.read_text("utf-8", errors="replace")
    if ext == ".vtt":
        text = VTT_TIMING.sub("", text)
    return text, None


def collect(hydrate: bool = False, max_mb: int = 50) -> tuple[list[dict], str]:
    docs = []
    stats = dict(content=0, name_only_cloud=0, name_only_type=0, too_big=0, errors=0)
    for root in roots():
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if not d.startswith(".")]
            for fn in filenames:
                if fn.startswith(".") or fn == "Icon\r":
                    continue
                p = Path(dirpath) / fn
                st = os.stat(p)
                rel = str(p.relative_to(root))
                doc = dict(ref=str(p), ts=datetime.fromtimestamp(st.st_mtime).isoformat(timespec="minutes"),
                           who=None, title=rel, body="")
                if p.suffix.lower() not in CONTENT_TYPES:
                    stats["name_only_type"] += 1
                elif st.st_size > max_mb * 1024 * 1024:
                    stats["too_big"] += 1
                elif st.st_blocks == 0 and not hydrate:
                    stats["name_only_cloud"] += 1
                else:
                    try:
                        doc["body"], doc["who"] = _extract(p)
                        stats["content"] += 1
                    except Exception as e:  # one bad file must not stop the build
                        stats["errors"] += 1
                        doc["body"] = ""
                        print(f"  ! {rel}: {type(e).__name__}: {e}")
                docs.append(doc)
    notes = ", ".join(f"{k}={v}" for k, v in stats.items())
    return docs, notes


def last_change() -> float:
    latest = 0.0
    for root in roots():
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if not d.startswith(".")]
            for name in filenames + [""]:
                try:
                    latest = max(latest, os.stat(os.path.join(dirpath, name)).st_mtime)
                except OSError:
                    pass
    return latest
