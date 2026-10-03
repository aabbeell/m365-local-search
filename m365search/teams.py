"""New Teams (v2) local cache: the Chromium IndexedDB inside the Teams container.

Read in place with ccl_chromium_reader; LevelDB .ldb files are immutable once written,
so no copy and no LOCK are needed. This is a rolling cache, not full history.
"""
import html
import re
from datetime import datetime
from pathlib import Path

from ccl_chromium_reader import ccl_chromium_indexeddb as idb

EBWEBVIEW = Path.home() / "Library/Containers/com.microsoft.teams2/Data/Library/Application Support/Microsoft/MSTeams/EBWebView"
TAG = re.compile(r"<[^>]+>")
KINDS = (("19:meeting_", "meeting"), ("@thread.tacv2", "channel"), ("@unq.gbl.spaces", "chat"), ("@thread.v2", "group"))


def stores() -> list[tuple[Path, Path | None]]:
    out = []
    for ldb in sorted(EBWEBVIEW.glob("WV2Profile_*/IndexedDB/https_teams.microsoft.com_0.indexeddb.leveldb")):
        blob = ldb.with_name("https_teams.microsoft.com_0.indexeddb.blob")
        out.append((ldb, blob if blob.exists() else None))
    return out


def _v(x):
    """ccl marks absent JS values with an _Undefined sentinel; text may arrive as bytes."""
    if type(x).__name__ == "_Undefined":
        return None
    if isinstance(x, bytes):  # V8 one-byte strings are Latin-1, not UTF-8
        try:
            return x.decode("utf-8")
        except UnicodeDecodeError:
            return x.decode("latin-1")
    return x


def _ts(ms) -> str | None:
    return datetime.fromtimestamp(ms / 1000).isoformat(timespec="minutes") if isinstance(ms, (int, float)) and ms > 0 else None


def _text(content) -> str:
    content = re.sub(r"<(br|/p|/div|/li)\b[^>]*>", "\n", content or "", flags=re.I)
    return html.unescape(TAG.sub("", content)).strip()


def _kind(conv_id: str) -> str:
    return next((k for pat, k in KINDS if pat in conv_id), "chat")


def _records(wrapper, db_prefix: str, store: str):
    for d in wrapper.database_ids:
        if d.name.startswith(db_prefix):
            db = wrapper[d.dbid_no]
            if store in db.object_store_names:
                for attempt in (1, 2):  # LevelDB compaction can delete a file mid-read
                    try:
                        yield from (r.value for r in db[store].iterate_records(errors_to_stdout=False)
                                    if isinstance(r.value, dict))  # skip tombstones
                        break
                    except FileNotFoundError:
                        if attempt == 2:
                            raise


def collect() -> tuple[list[dict], str]:
    messages: dict[tuple[str, str], dict] = {}
    topics: dict[str, str] = {}
    stats = dict(messages=0, skipped_system=0, skipped_deleted=0, conversations=0)
    for ldb, blob in stores():
        w = idb.WrappedIndexDB(ldb, blob)
        for conv in _records(w, "Teams:conversation-manager", "conversations"):
            tp = _v(conv.get("threadProperties")) or {}
            topic = _v(tp.get("topic"))
            if topic:
                topics[_v(conv.get("id")) or ""] = topic
            stats["conversations"] += 1
        for store in ("replychains", "replychains-2"):
            for chain in _records(w, "Teams:replychain-manager", store):
                for msg in (_v(chain.get("messageMap")) or {}).values():
                    mtype = _v(msg.get("messageType")) or ""
                    if not (mtype.startswith("RichText") or mtype == "Text"):
                        stats["skipped_system"] += 1
                        continue
                    if _v(msg.get("deletionInfo")):
                        stats["skipped_deleted"] += 1
                        continue
                    key = (_v(msg.get("conversationId")) or "", _v(msg.get("id")) or "")
                    prev = messages.get(key)
                    if prev is None or str(msg.get("version", "")) > str(prev.get("version", "")):
                        messages[key] = msg
    # conversations without a topic (1:1 and ad-hoc chats) are named after their senders
    senders: dict[str, list[str]] = {}
    for (conv_id, _), m in messages.items():
        name = _v(m.get("imDisplayName"))
        if name and name not in senders.setdefault(conv_id, []):
            senders[conv_id].append(name)
    docs = []
    for (conv_id, msg_id), m in messages.items():
        body = _text(_v(m.get("content")))
        if not body:
            continue
        title = topics.get(conv_id) or ", ".join(senders.get(conv_id, [])[:3]) or conv_id
        docs.append(dict(
            ref=f"teams:{conv_id}:{msg_id}",
            ts=_ts(_v(m.get("originalArrivalTime"))),
            who=_v(m.get("imDisplayName")),
            title=f"[{_kind(conv_id)}] {title}",
            body=body,
        ))
    stats["messages"] = len(docs)
    return docs, ", ".join(f"{k}={v}" for k, v in stats.items())


def last_change() -> float:
    return max((f.stat().st_mtime for ldb, _ in stores() for f in ldb.iterdir()), default=0.0)
