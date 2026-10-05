# m365search

Local, read-only full-text search over the Microsoft 365 data already synced to this Mac:

| Source   | Where it reads                                                                 | What you get |
|----------|---------------------------------------------------------------------------------|--------------|
| OneDrive | `~/Library/CloudStorage/OneDrive-*` (incl. Teams/SharePoint libraries added as shortcuts) | Every file by name; text of docx, xlsx, pptx, pdf, eml, vtt, txt/md/csv for files on disk |
| Teams    | New Teams IndexedDB cache in `~/Library/Containers/com.microsoft.teams2/...`    | Chat, group, meeting and channel messages still in the local cache (rolling, not full history) |
| Outlook  | New Outlook `HxStore.hxd` in `~/Library/Group Containers/UBF8T346G9.Office/...`, plus the large bodies it keeps in `.../Main Identity/Files/S*/*/EFMData/` | Sender, subject, delivery time, recipients and the full body of almost every cached message; a ~255 character preview for the rest. Meeting requests, cancellations and replies as items of their own, with start, end and location |
| Transcripts and recaps | Files downloaded from Teams (Recap > Transcript > Download) into `~/Documents/auchan-transcripts` (override with `M365SEARCH_TRANSCRIPTS`), plus the Teams meeting-recap and Loop (Fluid) caches | Full text of downloaded .vtt/.docx/.txt/.md transcripts, speaker per line; `[meeting recap]` per meeting whose Recap was opened in Teams (Copilot summary, topic notes, AI action items); `[meeting notes]` for the Loop meeting-notes page (decisions, open questions, notes, follow-ups); `[transcript available]` per recorded meeting. Recaps and Loop pages are kept in `meetings/` next to the index after Teams evicts them. Teams never caches the transcript text itself |
| Outlook attachments | Attachments New Outlook downloaded to `.../Main Identity/Files/S*/*/Attachments/` | One `[attachment]` hit per distinct file, linked to its mail by original name, with the text of pdf, docx, xlsx and pptx files; copy out with `save` |

Nothing is written next to the sources and nothing leaves the machine. The index lives at
`~/Library/Application Support/m365search/index.db` (override with `M365SEARCH_INDEX`); the attachment
text cache and the `previews` state file sit in the same folder.

## Use

```bash
uv run --project ~/m365-local-search m365search search "Data Governance"
uv run --project ~/m365-local-search m365search search "BigQuery" --source teams --json
uv run --project ~/m365-local-search m365search show "<ref from a search hit>"
uv run --project ~/m365-local-search m365search save "<attachment ref>" <folder>   # original file name
uv run --project ~/m365-local-search m365search status
uv run --project ~/m365-local-search m365search previews -n 10                  # mail to open in Outlook for its full body
uv run --project ~/m365-local-search m365search changes 2026-10-02             # everything new that day
uv run --project ~/m365-local-search m365search changes last --cursor me          # new since my previous 'last' call
uv run --project ~/m365-local-search m365search changes 2026-09-28 2026-10-02T12:00 --source teams --json
uv run --project ~/m365-local-search m365search build --hydrate   # also download cloud-only OneDrive files
```

`search` rebuilds any source whose files changed since its last build (a few seconds), so the
index does not go stale. Plain words are ANDed and accent-insensitive; FTS5 syntax
(`"exact phrase"`, `OR`, `NEAR`, `prefix*`) passes through.

## MCP server

`m365search-mcp` serves the same index over stdio with the tools `search`, `grep` (regex over all indexed text), `changes` (everything in a
date or time range, newest first), `show`, `previews`, `calendar` and `status`. Registered for Claude Code with:

```bash
claude mcp add --scope user m365search -- uv run --project ~/m365-local-search m365search-mcp
```

`changes` uses each item's own time: mail delivery time, Teams message time, OneDrive modification time,
transcript recording or download time. Outlook items without a recovered time are not included.
Called without `since`, it returns what is new since the previous such call with the same `cursor`
(default "default"; first call covers 24 hours): items timestamped after it, plus items first indexed
after it (mail or messages that reached the local caches late). Use one cursor per consumer.

## Outlook previews and full bodies

New Outlook stores a message as a header object and a body object that share one local id, and almost
every message whose header is cached here has its body too. A body too large for its object sits whole
in a gzip file under `Files/S*/*/EFMData/`, named by the body object, which keeps only its opening part;
that file is read. What keeps only a ~255 character preview is mostly a conversation summary whose
message is not cached; opening the mail in Outlook on this Mac should bring it in. Whether Outlook keeps
the body of mail never opened there is not established. `previews` lists the messages that have only a preview,
newest first: date, sender, subject, the start of the preview and a search string to paste
into Outlook (`from:address "subject words"`). Options: `--since DATE`, `-n N`, `--cursor NAME`, `--json`.
A copy of a mail whose preview sits in the body another copy of the same Message-ID holds is not listed,
copies with the same sender, subject and preview are listed once, and meeting items are not listed (a
reply to a meeting request mostly has no body to fetch). Nor is a message whose body object is cached
but empty or too short to keep (such as a mail with only an attachment): there is nothing more to
fetch. A record without a sender stays searchable but is not listed, since nothing places it in Outlook.

Each run starts by reporting how many messages listed by the previous run with the same cursor
(default "default") have a full body now. Open a few in Outlook, run `previews` again, and that line
shows whether opening fetched them. The run also rebuilds the Outlook part of the index, so `show`
prints the new bodies at once. The state file (`previews_state.json` next to the index) holds only
hashes of the listed messages per cursor and is rewritten on every run. `status` prints the
preview-only count of the last build.

Refs are unique: when several messages carry the same Message-ID (two local copies of one mail, or
replies whose own id was not found), each gets `#` plus 8 hex characters. An older ref without the
suffix still works with `show`, which lists the matching copies when there is more than one, and an
older suffixed ref finds the copy left without one. The first build with unique refs renames some old
mail; it counts as old for `changes last`, not as new, and so does a ref that later gains, loses or
changes its suffix. The same holds for the first build that merges by local id, which re-dates mail
without a Message-ID to its delivery minute and finds a few Message-IDs the old walk missed, and for the
first build with meeting items, which were cached before but not indexed. That build also takes a
summary's sender from its From address, which moves summaries of the user's own mail from a recipient
to the user, and takes about twenty rows off the `previews` list (junk, copies of cached mail, mail
whose cached body is empty); the first `previews` run after it reports most of them as no longer in
the Outlook cache.

## Notes

- OneDrive: content is read only from files already on disk unless `--hydrate`, because reading a
  cloud-only file makes OneDrive download it. Pin folders with "Always Keep on This Device".
- Teams: read in place with `ccl_chromium_reader`; `.ldb` files are immutable, so no copy or lock.
- Outlook: `outlook.py` is a Python port of the field walk in
  [securized/hxstore-reverse-engineering](https://github.com/securized/hxstore-reverse-engineering)
  (MIT), with these changes. Blocks are split into typed objects with a local object id, after the
  record model of [mitchell-johnson/hxstore-decode](https://github.com/mitchell-johnson/hxstore-decode)
  (MIT); the object layout and type codes used here were found on this store. A message's
  header object and its HTML body object carry the same id: that id hands the body to its header and
  merges every revision and copy of the message, so two mails from one sender in one minute stay
  apart and copies of one mail no longer split by minute. Each field walk stays inside its own object
  instead of reading a neighbour's sender, subject or time. A body no header claims by id falls back
  to the header whose preview opens it, never to a header with a body object of its own. A body
  stored as several HTML parts (a reply's own text and the quoted mail) is read whole, each part up to
  the NUL behind it, so the suggested replies and ids stored after it stay out. The time is
  the delivery time stored in the header object; elsewhere it is the earliest timestamp followed by
  a delivery stamp, not the plain minimum, which on this store sometimes picked stray older values.
  A conversation summary that repeats a cached message is folded into it. A reply's Message-ID is
  the one stored right after its In-Reply-To, not the parent's. Previews skip opaque values (`eyJ...`
  tokens, ids), and every distinct preview of a message stays searchable. A body kept in an
  `EFMData` file is read from it (gzip of the whole HTML) when its text holds the start of the part
  in the store.
- Outlook drafts: every header, body and summary object carries a change number that rises with
  each save. Sender, name and subject are the values most copies of the last save agree on, the
  preview and time those of the last save that has them; the delivery time is no order, a later save
  of a draft can carry an earlier one. Next to a draft's text Outlook keeps a second body object
  with only the quoted mail or nothing, which can carry the higher change number, so a doc shows the
  body with the highest change number that holds the preview of the header's last save (else the
  highest change number), not the longest.
  Of a message's own saves only the last keeps its preview searchable.
- Outlook junk: the ItemClass also sits in action backups and search caches (`~/ActionsBackup/...`:
  search cards, people suggestions, image names); only header, body and summary objects count, one
  record per object. A summary's sender is its From address, which comes first; the second address is
  who Outlook's list shows, the recipient of the user's own mail, and the addresses behind a summary
  belong to the next summaries, not to its recipients. The summary Outlook writes while the user
  composes a message carries that message's own id and leftovers of earlier saves, and is skipped.
  Display names (stored twice, or packed) are never taken as a preview, the line `Sent from Outlook
  for Mac` never as a subject or name, and the subject of the user's own reply never as a name; in a
  header object a subject stored once is the topic right behind its own copy, not the first run that looks like one. A
  subject still being typed when a draft was first saved can stay as the topic; a longer subject
  that extends it and is stored twice whole is the one the mail went out with.
- Outlook meeting items: a meeting request, its cancellation and each accept, decline or tentative
  reply is its own hit (ref, sender and delivery time as for mail), never merged into a mail. It is
  stored like a message, with its own ItemClass, plus a time object and, for a request, a place object
  under the same id; its text starts with plain lines such as `Meeting request`, `Start: 2026-01-15 10:00`,
  `End: ...` (local time) and `Location: ...`, so they are searchable. For a recurring series the start
  is the first occurrence. Conversation summaries of meeting items are skipped (each one repeats a cached
  item), and so are calendar copies of an event without a header object. A report that a meeting request
  was not delivered counts as mail.
