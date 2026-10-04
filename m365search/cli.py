"""Local, read-only full-text search over OneDrive, Teams and Outlook data synced to this Mac."""
import argparse
import hashlib
import json
import re
import shutil
import sys
from pathlib import Path

from . import index, onedrive, outlook, teams, transcripts

SOURCES = {"onedrive": onedrive, "teams": teams, "outlook": outlook, "transcripts": transcripts}
PREVIEWS_STATE = index.INDEX_PATH.parent / "previews_state.json"


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


def show_text(ref: str, rows: list, pad: bool = True) -> str:
    """The documents index.get found for a ref. Different refs (an old ref whose copies now carry a "#suffix")
    are listed to pick from; one ref held by several documents (an index not rebuilt yet) prints each in full."""
    if len({r[4] for r in rows}) > 1:
        lines = [f"{len(rows)} documents match {ref!r}; show one by its full ref:"]
        lines += [f"{r[4]} | {r[2] or '-'} | {(r[1] or '-')[:16]} | {r[3]}" for r in rows]
        return "\n".join(lines)
    label = (lambda k: f"{k + ':':<8}") if pad else (lambda k: k + ": ")
    texts = [f"{label('source')}{source}\n{label('date')}{ts or '-'}\n{label('from')}{who or '-'}\n"
             f"{label('title')}{title}\n{label('ref')}{r}\n\n{body}" for source, ts, who, title, r, body in rows]
    head = f"{len(rows)} documents share the ref {ref!r}:\n\n" if len(rows) > 1 else ""
    return head + f"\n\n{'-' * 40}\n\n".join(texts)


def cmd_show(a):
    rows = index.get(index.connect(), a.ref)
    if not rows:
        sys.exit(f"no document with ref {a.ref!r}")
    print(show_text(a.ref, rows))


def cmd_save(a):
    """Copy a cached Outlook attachment out under its original name."""
    rows = index.get(index.connect(), a.ref)
    row = rows[0] if len(rows) == 1 else None
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


def _who(d: dict) -> str:
    return f'{d["name"]} <{d["sender"]}>' if d["name"] and d["sender"] else d["name"] or d["sender"] or "-"


def _outlook_query(d: dict) -> str:
    """A search string to paste into Outlook: from:address "subject words" (or the preview's first words)."""
    subject = outlook._strip_prefix(d["subject"] or "")
    words = [w for w in (subject or d["preview"] or "").replace('"', " ").split() if any(c.isalnum() for c in w)]
    words = words[:8 if subject else 6]
    parts = [f"from:{d['sender']}"] if d["sender"] else []
    if words:
        parts.append('"' + " ".join(words) + '"')
    return " ".join(parts)


def _state_key(ref: str) -> str:
    return hashlib.sha1(ref.encode()).hexdigest()[:16]  # a ref can hold preview text; the state file must not


def preview_report(since: str | None = None, limit: int = 20, cursor: str = "default") -> dict:
    """Outlook messages that have only a preview, newest first, and which of this cursor's previous check
    have a full body now. The same parse rebuilds the Outlook part of the index, so show and search see new
    bodies at once. The state file next to the index keeps, per cursor, only hashes of the listed messages."""
    from datetime import datetime
    docs, notes = outlook.collect()
    index.replace_source(index.connect(), "outlook", docs, notes)
    docs = [d for d in docs if "has_body" in d and not d["meeting"]]  # messages, not attachments or meeting items
    # to_open: no body here, in another copy of the same Message-ID or in a cached body object, and a sender
    todo, listed = [], set()
    for d in sorted(filter(outlook.to_open, docs), key=lambda d: d["ts"] or "", reverse=True):
        if outlook.same_mail(d) not in listed:  # copies with the same sender, subject and preview: the newest
            listed.add(outlook.same_mail(d))
            todo.append(d)
    try:
        cursors = json.loads(PREVIEWS_STATE.read_text()).get("cursors", {})
    except (OSError, ValueError, AttributeError):
        cursors = {}
    prev = cursors.get(cursor)
    opened, gone = [], 0
    if prev:
        by_key, by_uid = {_state_key(d["ref"]): d for d in docs}, {d["uid"]: d for d in docs}
        for item in prev.get("items", []):  # a ref gains a #suffix when a second copy appears; the uid does not
            d = by_key.get(item.get("key")) or by_uid.get(item.get("uid"))
            if d is None:
                gone += 1
            elif d["has_body"] or d["body_ref"]:
                opened.append(d)
    cursors[cursor] = dict(checked=datetime.now().isoformat(timespec="seconds"),
                           items=[dict(key=_state_key(d["ref"]), uid=d["uid"]) for d in todo])
    try:
        PREVIEWS_STATE.parent.mkdir(parents=True, exist_ok=True)
        tmp = PREVIEWS_STATE.with_suffix(".tmp")
        tmp.write_text(json.dumps(dict(cursors=cursors)))
        tmp.replace(PREVIEWS_STATE)
    except OSError as err:
        print(f"(could not save {PREVIEWS_STATE}: {err})", file=sys.stderr)
    shown = [d for d in todo if not since or (d["ts"] or "") >= index._bound(since)][:max(limit, 0)]

    def row(d):
        return dict(ts=d["ts"], name=d["name"], sender=d["sender"], subject=d["subject"], preview=d["preview"],
                    ref=d["ref"], search=_outlook_query(d))
    copies = sum(1 for d in docs if d["body_ref"])
    return dict(cursor=cursor, preview_only=len(todo), with_full_body=sum(d["has_body"] for d in docs),
                copies_with_body_elsewhere=copies, last_check=prev.get("checked") if prev else None,
                last_check_preview_only=len(prev.get("items", [])) if prev else None,
                opened_since_last_check=[row(d) for d in sorted(opened, key=lambda d: d["ts"] or "", reverse=True)],
                gone_since_last_check=gone, items=[row(d) for d in shown])


def cmd_previews(a):
    rep = preview_report(a.since, a.n, a.cursor)
    if a.json:
        print(json.dumps(rep, ensure_ascii=False, indent=1))
        return
    date = lambda ts: (ts or "-")[:16].replace("T", " ")

    def flat(text, alt, n):
        s = " ".join((text or alt).split())
        return s[:n] + ("..." if len(s) > n else "")
    if rep["last_check"] is None:
        print(f"First check: {rep['preview_only']} Outlook messages have only a short preview here, not the full "
              "text.\nRun this command again after opening some; it then says which ones arrived.")
    else:
        print(f"Since the last check ({date(rep['last_check'])}): {len(rep['opened_since_last_check'])} of the "
              f"{rep['last_check_preview_only']} listed then have their full text now.")
        for d in rep["opened_since_last_check"]:
            print(f"  {date(d['ts'])} | {_who(d)} | {flat(d['subject'], '(no subject)', 70)}")
        if rep["gone_since_last_check"]:
            print(f"  ({rep['gone_since_last_check']} of them are no longer in the Outlook cache on this Mac)")
    print(f"\n{rep['preview_only']} messages have only a preview"
          + (f", newest {len(rep['items'])} below" if rep["items"] else "") + ":")
    if rep["copies_with_body_elsewhere"]:
        print(f"(not listed: {rep['copies_with_body_elsewhere']} copies of mail whose full text is already here)")
    for d in rep["items"]:
        print(f"\n{date(d['ts'])} | {_who(d)}\n  subject: {flat(d['subject'], '(no subject)', 80)}\n"
              f"  starts:  {flat(d['preview'], '-', 80)}\n  search in Outlook:  {d['search']}")
    print("\nOpen these in Outlook on this Mac (paste a search line into Outlook's search box), then run\n"
          "this command again. One that stays listed after you opened it may already be here inside\n"
          "another copy of the mail.")


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
    row = con.execute("SELECT notes FROM runs WHERE source = 'outlook'").fetchone()
    m = re.search(r"\bpreview_only=(\d+)", row[0] or "") if row else None
    if m:
        print(f"  outlook preview-only: {m[1]} messages without a full body; 'm365search previews' lists which to open in Outlook")


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
    pv = sub.add_parser("previews", help="Outlook mail with only a preview: which to open in Outlook to get the full body")
    pv.add_argument("--since", help="YYYY-MM-DD or YYYY-MM-DDTHH:MM (local time)")
    pv.add_argument("-n", type=int, default=20)
    pv.add_argument("--cursor", default="default", help="name of the previous check to compare with (one per consumer)")
    pv.add_argument("--json", action="store_true")
    pv.set_defaults(fn=cmd_previews)
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
