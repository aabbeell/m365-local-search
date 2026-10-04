"""MCP server (stdio) over the local m365search index: Auchan Outlook, Teams, OneDrive and meeting transcripts.

Read-only. Everything stays on this Mac; the index is refreshed from the local caches on each call.
"""
import io
import json
from contextlib import redirect_stderr, redirect_stdout

from mcp.server.mcpserver import MCPServer

from . import cli, index

mcp = MCPServer("m365search", instructions=(
    "Read-only search over the Auchan Microsoft 365 data synced to this Mac: Outlook mail and attachments, "
    "Teams chats and meeting chats, Auchan OneDrive files, and Teams meeting transcripts. Use `search` for "
    "topics, `changes` for what is new in a date or time range, `show` for the full text of one hit. "
    "Outlook has the full body of most cached mail and a ~255-character preview of the rest; Teams is a "
    "rolling cache, not full history. `previews` lists the mail that has only a preview, so the user can open "
    "those messages in Outlook on this Mac. "
    "`search` matches whole words; use `grep` for regex or partial strings (ticket numbers, codes, parts of "
    "names). Also search the files directly with your own grep/glob tools: OneDrive root "
    "~/Library/CloudStorage/OneDrive-AuchanMagyarországKft and transcripts ~/Documents/auchan-transcripts."
))

SOURCES = list(cli.SOURCES)


def _con(sources, refresh=True):
    con = index.connect()
    if refresh:
        with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):  # stdout is the MCP channel
            cli.refresh(con, sources or SOURCES)
    return con


def _check(sources):
    bad = [s for s in sources or [] if s not in SOURCES]
    if bad:
        raise ValueError(f"unknown source {bad}; use any of {SOURCES}")


@mcp.tool()
def search(query: str, sources: list[str] | None = None, limit: int = 20) -> str:
    """Full-text search. Plain words are ANDed and accent-insensitive; FTS5 syntax ("exact phrase", OR,
    NEAR, prefix*) passes through. sources: any of onedrive, teams, outlook, transcripts (default all).
    Returns hits with a `ref` to pass to `show`."""
    _check(sources)
    counts, rows = index.search(_con(sources), query, sources, limit)
    return json.dumps(dict(query=query, counts=counts, hits=[
        dict(source=s, ts=ts, who=who, title=title, ref=ref, snippet=" ".join(snip.split()))
        for s, ts, who, title, ref, snip in rows]), ensure_ascii=False)


@mcp.tool()
def changes(since: str | None = None, until: str | None = None, sources: list[str] | None = None,
            limit: int = 300, cursor: str = "default") -> str:
    """Everything new or changed: mail, Teams messages, modified OneDrive files, transcripts. Newest first.
    - With `since` (and optional `until`): items timestamped in that range. 'YYYY-MM-DD' or
      'YYYY-MM-DDTHH:MM' local time; a bare date as until covers the whole day; until defaults to since.
    - Without `since`: everything new since the previous since-less call with the same `cursor` (also items
      that reached the local caches late), then the cursor moves to now. The first call covers 24 hours.
      Use a distinct cursor name per consumer (e.g. "daily-brief") so they don't consume each other's news.
    Returns per-source counts and items with a short preview and a `ref` for `show`."""
    _check(sources)
    con = _con(sources)
    if since is None:
        frm, to, counts, rows = index.changes_since_cursor(con, cursor, sources, limit)
        head = dict(since_last_call=frm, now=to, cursor=cursor)
    else:
        counts, rows = index.changes(con, since, until, sources, limit)
        head = dict(since=since, until=until or since)
    return json.dumps(dict(**head, counts=counts, truncated=sum(counts.values()) > limit,
                           items=[dict(source=s, ts=ts, who=who, title=title, ref=ref, preview=" ".join(body.split()))
                                  for s, ts, who, title, ref, body in rows]), ensure_ascii=False)


@mcp.tool()
def grep(pattern: str, sources: list[str] | None = None, ignore_case: bool = True, limit: int = 50) -> str:
    """Regular-expression search over the full indexed text (title, sender, body) of every source, including
    Teams and Outlook, which are not files on disk. Use for partial strings, codes, numbers, or patterns
    `search` cannot express. Returns hits with the matching line and a `ref` for `show`."""
    import re
    _check(sources)
    rx = re.compile(pattern, re.I if ignore_case else 0)
    con = _con(sources)
    where, args = "", []
    if sources:
        where, args = f"WHERE source IN ({','.join('?' * len(sources))})", sources
    hits, total = [], 0
    for src, ts, who, title, ref, body in con.execute(
            f"SELECT source, ts, who, title, ref, body FROM docs {where} ORDER BY ts DESC", args):
        text = f"{title or ''}\n{who or ''}\n{body or ''}"
        m = rx.search(text)
        if not m:
            continue
        total += 1
        if len(hits) < limit:
            a, b = text.rfind("\n", 0, m.start()) + 1, text.find("\n", m.end())
            hits.append(dict(source=src, ts=ts, who=who, title=title, ref=ref,
                             line=" ".join(text[a:b if b != -1 else None].split())[:300]))
    return json.dumps(dict(pattern=pattern, total=total, hits=hits), ensure_ascii=False)


@mcp.tool()
def show(ref: str) -> str:
    """Full text of one document, by the `ref` from a search or changes result."""
    rows = index.get(index.connect(), ref)
    if not rows:
        return f"no document with ref {ref!r}"
    return cli.show_text(ref, rows, pad=False)


@mcp.tool()
def previews(since: str | None = None, limit: int = 20, cursor: str = "default") -> str:
    """Outlook mail that has only its ~255-character preview in the local cache, no full body, newest first.
    Each item has an Outlook search string (`search`) the user can paste to find and open it in Outlook on
    this Mac. `opened_since_last_check` lists messages from this cursor's previous check that have a full
    body now; each call moves the cursor, so use a distinct cursor name per automated consumer. Also
    refreshes the Outlook part of the index. since: 'YYYY-MM-DD' or 'YYYY-MM-DDTHH:MM' local time."""
    with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
        report = cli.preview_report(since, limit, cursor)
    return json.dumps(report, ensure_ascii=False)


@mcp.tool()
def calendar(days: int = 7) -> str:
    """Agenda from the Teams calendar cache for the next `days` days, local time."""
    buf = io.StringIO()
    with redirect_stdout(buf):
        cli.cmd_calendar(type("A", (), {"days": days})())
    return buf.getvalue() or "no events"


@mcp.tool()
def status() -> str:
    """Which sources are indexed, when they were built, and whether they are stale."""
    buf = io.StringIO()
    with redirect_stdout(buf), redirect_stderr(io.StringIO()):
        cli.cmd_status(None)
    return buf.getvalue()


def main():
    mcp.run()


if __name__ == "__main__":
    main()
