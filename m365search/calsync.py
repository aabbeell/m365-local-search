"""Mirror the Auchan Teams calendar (the Teams desktop cache) into a dedicated Apple Calendar.

One-way and stateless: every mirrored event carries a `teams-sync:<key>` line in its notes,
each run re-reads both sides and creates, updates or removes events only inside the date
window Teams currently has cached, so events outside it are never touched. Declined and
cancelled meetings are skipped. The target calendar lives in "On My Mac" or iCloud, never
in a Google or Exchange account, and gets no meeting body and no attendee list.
"""
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import EventKit
from Foundation import NSDate, NSRunLoop, NSURL

from ccl_chromium_reader import ccl_chromium_indexeddb as idb

from . import teams

CALENDAR_TITLE = "Auchan (Teams)"
MARKER = "teams-sync:"
CANCELLED = ("canceled:", "cancelled:", "lemondva:", "törölve:")
LOG = Path.home() / "Library/Logs/m365search/calendar-sync.log"
SOURCE_TYPES = {0: "On My Mac", 1: "Exchange", 2: "CalDAV", 3: "MobileMe", 4: "Subscribed", 5: "Birthdays"}


def teams_events() -> list[dict]:
    """Calendar entries cached by the Teams desktop app. Times are stored as naive UTC."""
    events = {}
    for ldb, blob in teams.stores():
        w = idb.WrappedIndexDB(ldb, blob)
        for d in w.database_ids:
            if not d.name.startswith("Teams:calendar:"):
                continue
            db = w[d.dbid_no]
            if "calendar" not in db.object_store_names:
                continue
            for r in db["calendar"].iterate_records(errors_to_stdout=False):
                v = r.value
                if not isinstance(v, dict):
                    continue
                start, end = teams._v(v.get("startTime")), teams._v(v.get("endTime"))
                if not isinstance(start, datetime) or not isinstance(end, datetime):
                    continue
                key = teams._v(v.get("objectId")) or f'{teams._v(v.get("iCalUID"))}|{start.isoformat()}'
                events[str(key)] = dict(
                    key=str(key),
                    subject=(teams._v(v.get("subject")) or "(no subject)").strip(),
                    start=start.replace(tzinfo=timezone.utc),
                    end=end.replace(tzinfo=timezone.utc),
                    location=(teams._v(v.get("location")) or "").strip(),
                    organizer=teams._v(v.get("organizerName")) or teams._v(v.get("organizerAddress")) or "",
                    response=teams._v(v.get("myResponseType")) or "",
                    show_as=teams._v(v.get("showAs")) or "",
                    join=teams._v(v.get("skypeTeamsMeetingUrl")) or teams._v(v.get("shortOnlineMeetingJoinUrl")) or "",
                )
    return sorted(events.values(), key=lambda e: e["start"])


def _mirrored(e: dict) -> bool:
    return e["response"] != "Declined" and not e["subject"].lower().startswith(CANCELLED)


def _ns(dt: datetime) -> NSDate:
    return NSDate.dateWithTimeIntervalSince1970_(dt.timestamp())


def _notes(e: dict) -> str:
    lines = [f"Organizer: {e['organizer']}", f"My response: {e['response']}", f"Show as: {e['show_as']}"]
    if e["join"]:
        lines.append(f"Join: {e['join']}")
    return "\n".join(lines + ["", "Mirrored from the Teams calendar by m365search; edits here are overwritten.", MARKER + e["key"]])


def _request_access(store, timeout: float = 120) -> bool:
    if EventKit.EKEventStore.authorizationStatusForEntityType_(EventKit.EKEntityTypeEvent) == 3:
        return True
    done = {}
    store.requestFullAccessToEventsWithCompletion_(lambda granted, err: done.update(granted=granted))
    t0 = time.time()
    while "granted" not in done and time.time() - t0 < timeout:
        NSRunLoop.currentRunLoop().runUntilDate_(NSDate.dateWithTimeIntervalSinceNow_(0.5))
    return bool(done.get("granted"))


def _pick_source(store):
    sources = list(store.sources())
    local = [s for s in sources if s.sourceType() == 0]
    icloud = [s for s in sources if s.sourceType() == 2 and "icloud" in (s.title() or "").lower()]
    return (local or icloud or [None])[0]


def _target_calendar(store, title: str):
    for cal in store.calendarsForEntityType_(EventKit.EKEntityTypeEvent):
        if cal.title() == title:
            return cal
    source = _pick_source(store)
    if source is None:
        raise SystemExit("no 'On My Mac' or iCloud calendar source; refusing to put Auchan meetings into a Google or Exchange account")
    cal = EventKit.EKCalendar.calendarForEntityType_eventStore_(EventKit.EKEntityTypeEvent, store)
    cal.setTitle_(title)
    cal.setSource_(source)
    ok, err = store.saveCalendar_commit_error_(cal, True, None)
    if not ok:
        raise SystemExit(f"could not create calendar {title!r}: {err}")
    return cal


def _apply(ev, e: dict) -> bool:
    """Write the Teams fields onto an event; True when something changed."""
    want = dict(title=e["subject"], start=int(e["start"].timestamp()), end=int(e["end"].timestamp()),
                location=e["location"], url=e["join"], notes=_notes(e))
    have = dict(title=ev.title() or "", start=int(ev.startDate().timeIntervalSince1970()) if ev.startDate() else None,
                end=int(ev.endDate().timeIntervalSince1970()) if ev.endDate() else None,
                location=ev.location() or "", url=ev.URL().absoluteString() if ev.URL() else "", notes=ev.notes() or "")
    if want == have:
        return False
    ev.setTitle_(want["title"])
    ev.setStartDate_(_ns(e["start"]))
    ev.setEndDate_(_ns(e["end"]))
    ev.setLocation_(want["location"] or None)
    ev.setURL_(NSURL.URLWithString_(want["url"]) if want["url"] else None)
    ev.setNotes_(want["notes"])
    return True


def log(line: str) -> None:
    LOG.parent.mkdir(parents=True, exist_ok=True)
    stamped = f"{datetime.now():%Y-%m-%d %H:%M:%S} {line}"
    print(stamped)
    with LOG.open("a") as fh:
        fh.write(stamped + "\n")


def probe() -> None:
    import os
    try:
        os.listdir(teams.EBWEBVIEW)
        log(f"probe: Teams cache readable, {len(teams.stores())} store(s)")
    except OSError as err:
        log(f"probe: Teams cache NOT readable: {type(err).__name__}: {err}")
    store = EventKit.EKEventStore.alloc().init()
    granted = _request_access(store)
    log(f"probe: calendar access {'granted' if granted else 'NOT granted'}")
    if granted:
        store = EventKit.EKEventStore.alloc().init()
        for s in store.sources():
            log(f"probe: source {s.title()!r} ({SOURCE_TYPES.get(s.sourceType(), s.sourceType())})")
        chosen = _pick_source(store)
        log(f"probe: target source would be {chosen.title()!r}" if chosen else "probe: no acceptable source")


SNAPSHOT = Path.home() / "Library/Application Support/m365search/teams_calendar.json"


def _write_snapshot(events: list[dict]) -> None:
    """Plain JSON copy of the Teams calendar for local readers that lack access to the Teams folder."""
    import json
    rows = [{"start": e["start"].isoformat(), "end": e["end"].isoformat(), "title": e["subject"],
             "organizer": e["organizer"], "response": e["response"], "join": e.get("join") or None,
             "location": e.get("location") or ""} for e in events]
    try:
        SNAPSHOT.parent.mkdir(parents=True, exist_ok=True)
        tmp = SNAPSHOT.with_suffix(".tmp")
        tmp.write_text(json.dumps(rows, ensure_ascii=False))
        tmp.replace(SNAPSHOT)
    except OSError as err:
        log(f"snapshot write failed: {err}")


def sync(title: str = CALENDAR_TITLE, dry_run: bool = False) -> None:
    import os
    try:
        os.listdir(teams.EBWEBVIEW)
    except PermissionError:
        log("no access to the Teams data folder; give 'M365 Calendar Sync' Full Disk Access in System Settings > Privacy & Security")
        return
    except FileNotFoundError:
        pass
    events = teams_events()
    _write_snapshot(events)
    if not events:
        log("Teams calendar cache empty or unreadable (is Teams running and signed in?); nothing changed")
        return
    wanted = {e["key"]: e for e in events if _mirrored(e)}
    lo = events[0]["start"] - timedelta(days=1)
    hi = max(e["end"] for e in events) + timedelta(days=1)
    store = EventKit.EKEventStore.alloc().init()
    if not _request_access(store, timeout=0 if dry_run else 120):
        if dry_run:
            for e in wanted.values():
                print(f"  would mirror {e['start']:%Y-%m-%d %H:%M}Z  {e['subject']}")
        log("no calendar access; grant it to 'M365 Calendar Sync' in System Settings > Privacy & Security > Calendars")
        return
    store = EventKit.EKEventStore.alloc().init()
    cal = _target_calendar(store, title)
    existing = {}
    for ev in store.eventsMatchingPredicate_(store.predicateForEventsWithStartDate_endDate_calendars_(_ns(lo), _ns(hi), [cal])) or []:
        for line in (ev.notes() or "").splitlines():
            if line.startswith(MARKER):
                existing[line[len(MARKER):].strip()] = ev
    counts = dict(created=0, updated=0, removed=0, unchanged=0)
    for key, e in wanted.items():
        ev = existing.get(key)
        if ev is None:
            ev = EventKit.EKEvent.eventWithEventStore_(store)
            ev.setCalendar_(cal)
            _apply(ev, e)
            counts["created"] += 1
        elif _apply(ev, e):
            counts["updated"] += 1
        else:
            counts["unchanged"] += 1
            continue
        if not dry_run:
            ok, err = store.saveEvent_span_commit_error_(ev, EventKit.EKSpanThisEvent, False, None)
            if not ok:
                log(f"save failed for {e['subject']!r}: {err}")
    for key, ev in existing.items():
        if key not in wanted:  # gone from Teams, declined or cancelled; inside the window by construction
            counts["removed"] += 1
            if not dry_run:
                store.removeEvent_span_commit_error_(ev, EventKit.EKSpanThisEvent, False, None)
    if not dry_run:
        store.commit_(None)
    src = cal.source().title() if cal.source() else "?"
    log(f"{'DRY RUN ' if dry_run else ''}{title!r} in {src!r}: +{counts['created']} ~{counts['updated']} "
        f"-{counts['removed']} ={counts['unchanged']} (Teams window {lo:%Y-%m-%d}..{hi:%Y-%m-%d}, "
        f"{len(events)} cached, {len(events) - len(wanted)} declined or cancelled)")
