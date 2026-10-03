"""Local, read-only full-text search over OneDrive, Teams and Outlook data synced to this Mac."""
import argparse
import json
import shutil
import sys
from pathlib import Path

from . import index, onedrive, outlook, teams, transcripts

SOURCES = {"onedrive": onedrive, "teams": teams, "outlook": outlook, "transcripts": transcripts}


def build(con, names, hydrate=False, quiet=False):
    for name in names:
        kwargs = {"hydrate": hydrate} if name == "onedrive" else {}
        docs, notes = SOURCES[name].collect(**kwargs)
        index.replace_source(con, name, docs, notes)
        if not quiet:
            print(f"{name}: {len(docs)} docs ({notes})")


def refresh(con, names):
    """Rebuild only the sources whose files changed since their last build."""
    stale = [n for n in names if SOURCES[n].last_change() > index.built_epoch(con, n)]
    if stale:
        print(f"(refreshing {', '.join(stale)})", file=sys.stderr)
        build(con, stale, quiet=True)


def cmd_build(a):
    build(index.connect(), a.source or list(SOURCES), a.hydrate)


def cmd_search(a):
    con = index.connect()
    if not a.no_refresh:
        refresh(con, a.source or list(SOURCES))
    counts, rows = index.search(con, a.query, a.source, a.n)
    if a.json:
        print(json.dumps(dict(query=a.query, counts=counts, hits=[
            dict(source=s, ts=ts, who=who, title=title, ref=ref, snippet=" ".join(snip.split()))
            for s, ts, who, title, ref, snip in rows]), ensure_ascii=False, indent=1))
        return
    print(f'"{a.query}": {sum(counts.values())} hits (' + ", ".join(f"{k} {v}" for k, v in sorted(counts.items())) + ")")
    for source, ts, who, title, ref, snip in rows:
        print(f"{source:<8} | {(ts or '')[:16]:<16} | {(who or '-')[:28]:<28} | {(title or '')[:70]}")
        if not a.no_snippet:
            print("         " + " ".join(snip.split())[:160])


def cmd_changes(a):
    con = index.connect()
    if not a.no_refresh:
        refresh(con, a.source or list(SOURCES))
    if a.since == "last":
        a.since, a.until, counts, rows = index.changes_since_cursor(con, a.cursor, a.source, a.n)
    else:
        counts, rows = index.changes(con, a.since, a.until, a.source, a.n)
    if a.json:
        print(json.dumps(dict(since=a.since, until=a.until or a.since, counts=counts, items=[
            dict(source=s, ts=ts, who=who, title=title, ref=ref, preview=" ".join(body.split()))
            for s, ts, who, title, ref, body in rows]), ensure_ascii=False, indent=1))
        return
    print(f"{a.since} .. {a.until or a.since}: {sum(counts.values())} items (" + ", ".join(f"{k} {v}" for k, v in sorted(counts.items())) + ")")
    for source, ts, who, title, ref, body in rows:
        print(f"{source:<11} | {(ts or '')[:16]:<16} | {(who or '-')[:28]:<28} | {(title or '')[:70]}")


def cmd_show(a):
    row = index.get(index.connect(), a.ref)
    if not row:
        sys.exit(f"no document with ref {a.ref!r}")
    source, ts, who, title, ref, body = row
    print(f"source: {source}\ndate:   {ts or '-'}\nfrom:   {who or '-'}\ntitle:  {title}\nref:    {ref}\n\n{body}")


def cmd_save(a):
    """Copy a cached Outlook attachment out under its original name."""
    row = index.get(index.connect(), a.ref)
    if not row or not row[4].startswith("attachment:"):
        sys.exit(f"not an attachment ref: {a.ref!r}")
    body = row[5]
    original = body.split("\n", 1)[0].removeprefix("Attachment: ").strip()
    src = Path(next(l for l in body.splitlines() if l.startswith("File: ")).removeprefix("File: "))
    dest = Path(a.dir).expanduser() / original
    if dest.exists() and not a.force:
        sys.exit(f"exists: {dest} (use --force to overwrite)")
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)
    print(dest)


def cmd_calendar(a):
    """Agenda from the Teams calendar cache, local time."""
    from datetime import datetime, timedelta
    from .calsync import teams_events
    today = datetime.now().astimezone().replace(hour=0, minute=0, second=0, microsecond=0)
    for e in teams_events():
        s, t = e["start"].astimezone(), e["end"].astimezone()
        if today <= s < today + timedelta(days=a.days):
            print(f"{s:%a %m-%d %H:%M}-{t:%H:%M} | {e['subject']} | {e['organizer']} | {e['response']}")


def cmd_calendar_sync(a):
    from . import calsync
    calsync.probe() if a.probe else calsync.sync(a.calendar, a.dry_run)


def cmd_status(a):
    con = index.connect()
    print(f"index: {index.INDEX_PATH}")
    for root in onedrive.roots():
        print(f"onedrive root: {root}")
    for ldb, _ in teams.stores():
        print(f"teams store:   {ldb.parent.parent.name}/{ldb.name}")
    for store in outlook.stores():
        print(f"outlook store: {store.parent.name}/{store.name} ({store.stat().st_size / 1e6:.1f} MB)")
    print(f"transcripts:   {transcripts.TRANSCRIPTS_DIR}{'' if transcripts.TRANSCRIPTS_DIR.is_dir() else ' (missing)'}")
    for name in SOURCES:
        row = con.execute("SELECT built_at, docs, notes FROM runs WHERE source = ?", (name,)).fetchone()
        stale = SOURCES[name].last_change() > index.built_epoch(con, name)
        print(f"  {name:<11} " + (f"built {row[0]} | {row[1]} docs{' | STALE' if stale else ''} | {row[2]}" if row else "not built"))


def main(argv=None):
    ap = argparse.ArgumentParser(prog="m365search", description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build", help="rebuild the index for one or more sources")
    b.add_argument("--source", action="append", choices=list(SOURCES))
    b.add_argument("--hydrate", action="store_true", help="OneDrive: download cloud-only files to read their content")
    b.set_defaults(fn=cmd_build)
    s = sub.add_parser("search", help="full-text search across sources (refreshes changed sources first)")
    s.add_argument("query", help='words are ANDed; FTS5 syntax ("exact phrase", OR, NEAR, prefix*) passes through')
    s.add_argument("--source", action="append", choices=list(SOURCES))
    s.add_argument("-n", type=int, default=10)
    s.add_argument("--json", action="store_true")
    s.add_argument("--no-snippet", action="store_true")
    s.add_argument("--no-refresh", action="store_true")
    s.set_defaults(fn=cmd_search)
    c = sub.add_parser("changes", help="everything new or changed in a date or time range, newest first")
    c.add_argument("since", help='YYYY-MM-DD or YYYY-MM-DDTHH:MM (local time), or "last": new since the previous "last" call')
    c.add_argument("until", nargs="?", help="same format; a bare date covers the whole day (default: same as since)")
    c.add_argument("--source", action="append", choices=list(SOURCES))
    c.add_argument("--cursor", default="default", help='bookmark name for "last" (one per consumer)')
    c.add_argument("-n", type=int, default=200)
    c.add_argument("--json", action="store_true")
    c.add_argument("--no-refresh", action="store_true")
    c.set_defaults(fn=cmd_changes)
    sh = sub.add_parser("show", help="print one document in full, by the ref from a search")
    sh.add_argument("ref")
    sh.set_defaults(fn=cmd_show)
    sv = sub.add_parser("save", help="copy an Outlook attachment (ref from a search hit) to a folder, original name")
    sv.add_argument("ref")
    sv.add_argument("dir")
    sv.add_argument("--force", action="store_true")
    sv.set_defaults(fn=cmd_save)
    ca = sub.add_parser("calendar", help="agenda from the Teams calendar cache")
    ca.add_argument("--days", type=int, default=7)
    ca.set_defaults(fn=cmd_calendar)
    cs = sub.add_parser("calendar-sync", help="mirror the Teams calendar into an Apple Calendar")
    cs.add_argument("--calendar", default="Auchan (Teams)")
    cs.add_argument("--dry-run", action="store_true")
    cs.add_argument("--probe", action="store_true", help="request calendar access and list sources")
    cs.set_defaults(fn=cmd_calendar_sync)
    st = sub.add_parser("status", help="show discovered sources and index state")
    st.set_defaults(fn=cmd_status)
    a = ap.parse_args(argv)
    a.fn(a)


if __name__ == "__main__":
    sys.exit(main())
