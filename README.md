# m365search

Local, read-only full-text search over the Microsoft 365 data already synced to this Mac:

| Source   | Where it reads                                                                 | What you get |
|----------|---------------------------------------------------------------------------------|--------------|
| OneDrive | `~/Library/CloudStorage/OneDrive-*` (incl. Teams/SharePoint libraries added as shortcuts) | Every file by name; text of docx, xlsx, pptx, pdf, eml, vtt, txt/md/csv for files on disk |
| Teams    | New Teams IndexedDB cache in `~/Library/Containers/com.microsoft.teams2/...`    | Chat, group, meeting and channel messages still in the local cache (rolling, not full history) |
| Outlook  | New Outlook `HxStore.hxd` in `~/Library/Group Containers/UBF8T346G9.Office/...` | Sender, subject, send time, recipients and a ~255 character preview; full body only for mail opened in Outlook |
| Transcripts | Files downloaded from Teams (Recap > Transcript > Download) into `~/Documents/auchan-transcripts` (override with `M365SEARCH_TRANSCRIPTS`), plus the Teams cache | Full text of .vtt/.docx/.txt/.md transcripts, speaker per line; one `[transcript available]` marker per recorded meeting whose transcript sits in the organizer's OneDrive |
| Outlook attachments | Attachments New Outlook downloaded to `.../Main Identity/Files/S*/*/Attachments/` | One `[attachment]` hit per distinct file, linked to its mail by original name, with the text of pdf, docx, xlsx and pptx files; copy out with `save` |

Nothing is written next to the sources and nothing leaves the machine. The index lives at
`~/Library/Application Support/m365search/index.db` (override with `M365SEARCH_INDEX`).

## Use

```bash
uv run --project ~/m365-local-search m365search search "Data Governance"
uv run --project ~/m365-local-search m365search search "BigQuery" --source teams --json
uv run --project ~/m365-local-search m365search show "<ref from a search hit>"
uv run --project ~/m365-local-search m365search save "<attachment ref>" <folder>   # original file name
uv run --project ~/m365-local-search m365search status
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
date or time range, newest first), `show`, `calendar` and `status`. Registered for Claude Code with:

```bash
claude mcp add --scope user m365search -- uv run --project ~/m365-local-search m365search-mcp
```

`changes` uses each item's own time: mail send time, Teams message time, OneDrive modification time,
transcript recording or download time. Outlook items without a recovered time are not included.
Called without `since`, it returns what is new since the previous such call with the same `cursor`
(default "default"; first call covers 24 hours): items timestamped after it, plus items first indexed
after it (mail or messages that reached the local caches late). Use one cursor per consumer.

## Notes

- OneDrive: content is read only from files already on disk unless `--hydrate`, because reading a
  cloud-only file makes OneDrive download it. Pin folders with "Always Keep on This Device".
- Teams: read in place with `ccl_chromium_reader`; `.ldb` files are immutable, so no copy or lock.
- Outlook: `outlook.py` is a Python port of the field walk in
  [securized/hxstore-reverse-engineering](https://github.com/securized/hxstore-reverse-engineering)
  (MIT), with one change: the send time is the earliest timestamp followed by a delivery stamp,
  not the plain minimum, which on this store sometimes picked stray older values.
