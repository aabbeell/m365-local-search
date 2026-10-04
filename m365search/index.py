"""One FTS5 index shared by every source. Sources are rebuilt wholesale, never merged."""
import os
import sqlite3
from pathlib import Path

INDEX_PATH = Path(os.environ.get(
    "M365SEARCH_INDEX",
    Path.home() / "Library/Application Support/m365search/index.db",
))

SCHEMA = """
CREATE VIRTUAL TABLE IF NOT EXISTS docs USING fts5(
    source UNINDEXED, ref UNINDEXED, ts UNINDEXED, who, title, body,
    tokenize = 'unicode61 remove_diacritics 2'
);
CREATE TABLE IF NOT EXISTS runs (
    source TEXT PRIMARY KEY, built_at TEXT, docs INTEGER, notes TEXT, built_epoch REAL
);
-- when each ref first appeared in the index, so items that reach the local caches late still count as new
CREATE TABLE IF NOT EXISTS seen (ref TEXT PRIMARY KEY, source TEXT, first_seen TEXT);
-- named bookmarks for "changes since my last call"
CREATE TABLE IF NOT EXISTS cursors (name TEXT PRIMARY KEY, at TEXT);
"""

NOW = "strftime('%Y-%m-%dT%H:%M:%S', 'now', 'localtime')"

FTS_OPERATORS = ('"', " OR ", " AND ", " NOT ", "NEAR(", "*", "^")


def connect(path: Path = INDEX_PATH) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(path)
    con.executescript(SCHEMA)
    return con


def replace_source(con, source: str, docs: list[dict], notes: str = "") -> int:
    with con:
        con.execute("DELETE FROM docs WHERE source = ?", (source,))
        con.executemany(
            "INSERT INTO docs (source, ref, ts, who, title, body) VALUES (?, ?, ?, ?, ?, ?)",
            [(source, d["ref"], d.get("ts"), d.get("who"), d.get("title"), d.get("body") or "") for d in docs],
        )
        # the first build of a source backfills first_seen from the item's own time, not "now"
        first = con.execute("SELECT 1 FROM seen WHERE source = ? LIMIT 1", (source,)).fetchone() is None
        regrouped = False
        if source == "outlook" and not first:  # the first build with unique refs renames old mail: backfill it too
            prev = con.execute("SELECT notes FROM runs WHERE source = 'outlook'").fetchone()
            first = prev is not None and "preview_only=" not in (prev[0] or "")
            # so does the first build that merges by local id: it re-dates mail without a Message-ID, adds ids;
            # and the first with meeting items, which were not indexed before and are not new, and which also
            # takes a summary's sender from its From address, moving the user's own mail to the user
            regrouped = not first and prev is not None and not all(
                k in (prev[0] or "") for k in ("body_by_id=", "meetings="))
        when = "coalesce(d.ts, '')" if first else NOW
        if source == "outlook" and not first:  # a ref that gained, lost or changed its "#suffix" keeps its first_seen
            base = "substr(d.ref, 1, instr(d.ref || '#', '#') - 1)"
            # base#anything as a range on the primary key ('$' follows '#'), so seen is searched, not scanned
            when = (f"coalesce((SELECT min(s.first_seen) FROM seen s WHERE s.ref = {base} "
                    f"OR (s.ref >= {base} || '#' AND s.ref < {base} || '$')), "
                    + ("d.ts, '')" if regrouped else f"{NOW})"))
        con.execute(f"INSERT OR IGNORE INTO seen SELECT d.ref, d.source, {when} FROM docs d WHERE d.source = ?",
                    (source,))
        con.execute(
            "INSERT OR REPLACE INTO runs VALUES (?, datetime('now', 'localtime'), ?, ?, unixepoch('subsec'))",
            (source, len(docs), notes),
        )
    return len(docs)


def to_fts_query(q: str) -> str:
    """Plain words become an AND of quoted terms; anything with FTS5 syntax passes through."""
    if any(op in q for op in FTS_OPERATORS):
        return q
    return " ".join('"' + t.replace('"', '""') + '"' for t in q.split())


def search(con, q: str, sources: list[str] | None = None, limit: int = 10):
    where, args = "docs MATCH ?", [to_fts_query(q)]
    if sources:
        where += f" AND source IN ({','.join('?' * len(sources))})"
        args += sources
    counts = dict(con.execute(f"SELECT source, count(*) FROM docs WHERE {where} GROUP BY source", args))
    rows = con.execute(
        f"SELECT source, ts, who, title, ref, snippet(docs, -1, '[', ']', '...', 12) "
        f"FROM docs WHERE {where} ORDER BY rank LIMIT ?",
        args + [limit],
    ).fetchall()
    return counts, rows


def built_epoch(con, source: str) -> float:
    row = con.execute("SELECT built_epoch FROM runs WHERE source = ?", (source,)).fetchone()
    return row[0] or 0.0 if row else 0.0


def get(con, ref: str) -> list[tuple]:
    """Every document with this ref. An Outlook ref from before refs were made unique (no "#suffix")
    returns all the copies that now carry it with a suffix; a suffixed ref whose other copies left the
    cache finds the one left without it."""
    cols = "SELECT source, ts, who, title, ref, body FROM docs"
    rows = (con.execute(f"{cols} WHERE ref = ? ORDER BY ts DESC", (ref,)).fetchall()
            or con.execute(f"{cols} WHERE substr(ref, 1, length(?) + 1) = ? || '#' ORDER BY ts DESC",
                           (ref, ref)).fetchall())
    if not rows and ref.startswith("outlook:") and "#" in ref:
        rows = con.execute(f"{cols} WHERE ref = ? ORDER BY ts DESC", (ref.rpartition("#")[0],)).fetchall()
    return rows


def _bound(s: str, end: bool = False) -> str:
    """'2026-10-02' or '2026-10-02T14:00' -> comparable ts string; a bare end date covers the whole day."""
    s = s.strip().replace(" ", "T")
    return s + "T99" if end and len(s) == 10 else s


def changes(con, since: str, until: str | None = None, sources: list[str] | None = None, limit: int = 200):
    """Every document whose timestamp falls in [since, until], newest first, with per-source counts."""
    where, args = "ts IS NOT NULL AND ts >= ? AND ts <= ?", [_bound(since), _bound(until or since, end=True)]
    if sources:
        where += f" AND source IN ({','.join('?' * len(sources))})"
        args += sources
    counts = dict(con.execute(f"SELECT source, count(*) FROM docs WHERE {where} GROUP BY source", args))
    rows = con.execute(
        f"SELECT source, ts, who, title, ref, substr(body, 1, 400) FROM docs WHERE {where} ORDER BY ts DESC LIMIT ?",
        args + [limit],
    ).fetchall()
    return counts, rows


def changes_since_cursor(con, cursor: str = "default", sources: list[str] | None = None, limit: int = 300,
                         first_since: str | None = None):
    """Items that are new (first indexed) or newly timestamped since this cursor's previous call; then
    moves the cursor to now. A cursor's first call covers the last 24 hours unless first_since is given."""
    row = con.execute("SELECT at FROM cursors WHERE name = ?", (cursor,)).fetchone()
    since = row[0] if row else (first_since or con.execute(
        "SELECT strftime('%Y-%m-%dT%H:%M:%S', 'now', 'localtime', '-1 day')").fetchone()[0])
    now = con.execute(f"SELECT {NOW}").fetchone()[0]
    where = "(d.ts > :since OR s.first_seen > :since)"
    args = dict(since=since)
    if sources:
        where += f" AND d.source IN ({','.join(':s%d' % i for i in range(len(sources)))})"
        args.update({f"s{i}": v for i, v in enumerate(sources)})
    base = f"FROM docs d LEFT JOIN seen s ON s.ref = d.ref WHERE {where}"
    counts = dict(con.execute(f"SELECT d.source, count(*) {base} GROUP BY d.source", args))
    rows = con.execute(
        f"SELECT d.source, d.ts, d.who, d.title, d.ref, substr(d.body, 1, 400) {base} "
        f"ORDER BY coalesce(d.ts, s.first_seen) DESC LIMIT :limit", {**args, "limit": limit}).fetchall()
    with con:
        con.execute("INSERT OR REPLACE INTO cursors VALUES (?, ?)", (cursor, now))
    return since, now, counts, rows
