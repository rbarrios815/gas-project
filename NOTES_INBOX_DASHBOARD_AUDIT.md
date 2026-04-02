# NOTES INBOX ↔ Dashboard Interaction Audit

Generated on 2026-04-02 from repository history and current code.

## Current version (HEAD)

### Server-side interaction in `Code.gs`
- `updateNotes(clientName, type, newText, originalText)` updates `DASHBOARD 8.0` and, for `pastWork`, derives line-level note changes and calls `syncDashboardChangesToInbox_`.
- `syncDashboardChangesToInbox_` updates existing `NOTES INBOX` rows by exact note-text match scoped to assigned client, and returns `false` if any changed note is not found.
- Notes Inbox sheet model:
  - Sheet name constant: `NOTES INBOX`
  - Column A = note, B = assigned client, C = timestamp
  - `ensureNotesInbox_` lazily creates sheet + header row
- Inbox APIs used by dashboard UI:
  - `inboxAddNote(rawNote, assignedClient)` appends note/timestamp, optionally assigns canonical client, sends non-blocking email
  - `inboxGetRecent(limit)` returns timestamp-sorted recent notes plus raw rows and chip metadata (`getChipStateForClients`)
  - `inboxUpdateNote(row, newNote)` enforces row validity + non-empty note + required timestamp, then syncs back to dashboard through `syncInboxNoteToDashboard_`
  - `inboxAssignToClient(row, clientTypedName)` canonicalizes client, writes column B assignment, then calls `updatePastWork`
- Inbox→Dashboard sync path:
  - `syncInboxNoteToDashboard_` looks up assigned client row in dashboard, finds matching note line in Column C (`pastWork`) using `findNoteLineIndex_`, preserves existing date prefix, and rewrites that line.
  - If no matching line is found, current HEAD returns `false` (does not append).

### Client-side interaction in `Index.html`
- Add note flow: `addNote()` calls `.inboxAddNote(txt, preferredClient)` where `preferredClient` is current active client when available.
- Edit flow: `saveEditedNote(row, text, noteEl)` calls `.inboxUpdateNote(row, text)`; on failure restores original text in UI.
- Load/render flow:
  - `loadRecent()` calls `.inboxGetRecent(100)` and `renderRecent(payload)`.
  - `renderRecent` builds editable note rows, assignment input, and optional chip metadata.
- Assign flow:
  - UI suggests likely clients via local fuzzy scoring (`scoreClientAgainstNote`, Levenshtein-based) and commits assignment through `.inboxAssignToClient(it.row, name)`.
- Dashboard submit mirroring:
  - `submitPastWork()` calls `logPastWorkSubmissionToInbox(pastWorkContent, clientName)`.
  - `logPastWorkSubmissionToInbox` mirrors submitted past work into inbox through `.inboxAddNote(trimmed, clientName || '')`.

## Last 25 commits from HEAD: did this integration change?

Scope checked: `HEAD` through `HEAD~24`.

Result: **no direct changes to NOTES INBOX API call names or their invocation points in these 25 commits**; most changes target category/label and birthday logic. Some large restore commits shift line numbers and simplify helper internals.

Notable behavior regressions within this 25-commit window (detected by comparing `HEAD~24` vs `HEAD`):
- `deriveNoteChanges_` no longer carries `dateText` per changed line.
- `syncDashboardChangesToInbox_` no longer auto-appends unmatched changed notes into `NOTES INBOX`; now it only updates existing matches and fails otherwise.
- `syncInboxNoteToDashboard_` no longer appends `newNote` to dashboard past-work when no matching line is found; now it returns `false`.

These are robust-sync behavior reductions and can look like a reversion in cross-view consistency.

## Older commits where this integration changed materially

Chronology (older than the last 25 commits):
- `d7732a3` initial wiring in UI to call `inboxAddNote`, `inboxAssignToClient`, and `inboxGetRecent`.
- `cccac70` added core `NOTES INBOX` backend block (`NOTES_SHEET`, `ensureNotesInbox_`, `inboxAddNote`, `inboxGetRecent`, `inboxAssignToClient`).
- `e5baa9d` changed UI call to `inboxAddNote(txt, preferredClient)` and increased fetch to `inboxGetRecent(50)`.
- `b176457` updated backend signature to `inboxAddNote(rawNote, assignedClient)`.
- `518913e` added `logPastWorkSubmissionToInbox` to mirror `submitPastWork` into inbox.
- `00a8249` added `inboxUpdateNote` editing flow and increased list fetch to 100.
- `c6ee2e4` restored assignment parameter support when copying notes to inbox.
- `c403e56` added bidirectional strict sync helpers (`deriveNoteChanges_`, `syncDashboardChangesToInbox_`, `syncInboxNoteToDashboard_`, `findNoteLineIndex_`).

## Reimplementation candidates (most robust features)

If rebuilding to avoid reversion risk, the highest-value behaviors to preserve are:
1. **Bidirectional sync with deterministic matching and explicit failure signals** (`EDIT FAILED ON OTHER VIEW`).
2. **Auto-append fallback on unmatched edits** (present in older `HEAD~24`, absent in current HEAD).
3. **Date-aware line matching for past-work note edits** so same text on different dates resolves correctly.
4. **Canonical client normalization** across inbox assignment and dashboard lookup.
5. **UI edit rollback on server failure** and non-blocking UX to keep dashboard usable.
6. **Submission mirroring into inbox** from main dashboard submit action.
