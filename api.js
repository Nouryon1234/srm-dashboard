# Data Mapping & Behavior Notes

Reference for how the frontend's in-memory shapes map to SharePoint list fields, and the behavioral decisions made when the old hardcoded-data model didn't translate 1:1 into a live, multi-user data store.

## Field mapping tables

See `api/src/lib/mappers.js` for the authoritative, executable version of every mapping below — this doc explains the *why*, the code has the exact *what*.

| Frontend shape | SharePoint list | Notes |
|---|---|---|
| `S[]` (suppliers) | Suppliers | `comm`/`focus` arrays and the `extra.*` sub-object are stored as JSON-in-a-text-column (`CommFlags`, `FocusAreas`, `News`) since SharePoint list columns are flat |
| `KPIS[sid][metricKey]` | KPIRecords | One row per supplier per metric; API groups them back into the nested shape via `?grouped=1` |
| `NS[sid][]` (meetings) | Meetings | `files` (attachments) and `kpis` (KPI snapshot) are JSON-in-text-column, same reasoning as above |
| `RISKS[sid][]` (actions) | Actions | Also referred to as "risks" in the frontend code/UI — same underlying list, see below |
| History points for trend charts | PerformanceHistory | Populated either directly (`POST /api/history`) or via a meeting's `KpiSnapshot` |

## Actions vs. Risks: one list, two names

The original static dashboard used `RISKS[sid]` as the variable name for what the UI actually labels "Pending Actions" / "Actions" throughout — severity + status + owner + due date, materialized from meeting action-item text. There was never a functionally separate "risk register" in the original code; it was one data structure wearing two names depending on which part of the file you were reading.

This build keeps that as one SharePoint list (`Actions`) rather than inventing a second list that the original app never actually used. `docs/sharepoint-schema.md` documents a separate `Risks` list schema for completeness/future use, but the API's `RISKS_LIST_NAME` setting defaults to pointing at `Actions` — change that setting (and populate a real `Risks` list) only if you decide governance requires actually splitting them.

## Deletion behavior

**Deleting a supplier does not cascade-delete its meetings, actions, or history rows.** This is a deliberate change from the static version, where `removeSupplier()` used JavaScript's `delete NS[id]` etc. to instantly and silently wipe every related record along with the supplier.

Why this changed: in a shared, persistent data store, "delete supplier" is a much heavier action than it was in a single-session in-memory prototype — other people may be looking at that supplier's meeting history right now, and SharePoint's audit trail (`Created By`, `Modified`, version history) makes those child records a real historical record, not scratch data. Silently cascading the delete would destroy that history with no undo.

Current behavior: deleting a supplier removes only the Suppliers list item. Its Meetings/Actions/PerformanceHistory rows remain, now "orphaned" (pointing at a `SupplierId` that no longer resolves to a name in the UI). If you want cascade deletion restored, add it explicitly as a follow-up — it's a straightforward addition to the `suppliers.js` Function's `DELETE` handler (loop through the other four lists filtering by `SupplierId` and delete matches) — but it wasn't included by default given the stakes of an irreversible bulk delete against a shared list.

**Deleting a meeting or an action is a permanent SharePoint delete**, gated by a `confirm()` dialog in the UI. There's no "trash" — SharePoint's own list recycle bin (Site Settings → Recycle Bin) is the safety net if something gets deleted by mistake, exactly as it would be for any other SharePoint list.

## "Solved Actions" no longer a separate store

The static version had a `SOLVED_ACTIONS{}` object that both resolved actions *and* deleted actions moved into, as a manually-maintained "historical record" living only in that browser tab's memory (gone on refresh).

That concept doesn't map cleanly onto a shared data store: SharePoint already keeps a real, durable history (every item has system `Created`/`Modified` timestamps and, if versioning is enabled on the list, a full version history you can open from the SharePoint UI itself). Duplicating that into a second client-side array would mean maintaining two sources of truth that can drift.

What changed:
- **Resolving** an action now just sets its `Status` field to `Resolved` in SharePoint — reversible via "Reopen" in the UI (`unresolveAction`), which sets `Status` back to `Open`.
- **Deleting** an action is a real, permanent delete (see above) — it no longer gets silently archived into a shadow list first.

If you want a formal "resolved actions" audit view, the more sustainable version is a SharePoint view or Power BI report filtered to `Status = Resolved`, rather than a second parallel list this app maintains — turn on versioning on the Actions list (List Settings → Versioning Settings → Yes) if you want to browse the exact edit history of each item's Status changes over time.

## Undo behavior

The static version's undo stack could instantly reverse array splices because everything lived in memory. Against a live API:

- **Edits and status changes** remain undoable — the previous field values are snapshotted client-side before the API call, and "Undo" sends a follow-up `PUT` restoring them.
- **Deletions are not undoable** through the Undo button. Recreating a deleted SharePoint item via the API would give it a new item ID and blank `Created`/`Modified` history — it wouldn't actually be "the same" record. This is why `deleteAction`/`deleteNote`/`removeSupplier` all require an explicit `confirm()` dialog instead of relying on the Undo stack as a safety net.

## KPI trend history: no more synthetic data

The static build's `synthesizeKpiHistory()` function used a seeded pseudo-random generator to fabricate a plausible-looking trend line for every supplier/metric combination, anchored loosely to the one real "current" KPI scorecard value it had. This existed purely because there was no real historical data source to draw from.

That function is gone. Trend charts now read only real `PerformanceHistory` rows (via `hydrateKpiHistoryOntoMeetings()` in `app.js`, which merges history points onto matching-date meetings so the existing chart-rendering code needs no changes). A supplier with meetings logged but no history rows yet will show "Not enough history yet" on its trend charts — this is expected and correct: it means real data hasn't been entered for that metric yet, not that something is broken.

## Auto-refresh and multi-user editing

`js/sync.js` polls `/api/dashboard` every 30 seconds (configurable via `window.SRM_CONFIG.refreshIntervalMs`) and re-renders only if the fetched snapshot actually differs from what's currently shown (cheap hash comparison, not a full diff). This means:

- Two people editing different suppliers will each see the other's changes appear within ~30 seconds, without a manual refresh.
- There's no field-level conflict resolution — if two people edit the *same* meeting note at the same time, the last `PUT` to reach the API wins, same as any other last-write-wins system without optimistic locking. SharePoint's own version history lets you see and manually recover a clobbered edit after the fact, but the app doesn't warn you about the conflict in the moment.
- The manual "↻ Refresh" button in the header calls `window.SRM_SYNC.refreshNow()` for anyone who doesn't want to wait out the poll interval.

## Known gaps / explicitly out of scope for this pass

- **No authentication.** Per this project's current scope, anyone who can reach the deployed URL can read and write all data. Adding Entra ID sign-in is described as a follow-up in `docs/azure-setup.md` §6.
- **No server-side validation beyond required-field checks.** The Functions check that required fields are present but don't validate formats (e.g. a malformed date string will be accepted and likely render oddly in the calendar).
- **No pagination on large lists in the UI.** The API's `getAllListItems` already pages through SharePoint's 200-item response limit internally, so it returns complete data regardless of list size — but the frontend's tables/charts render everything at once, which is fine at dozens-to-low-hundreds of suppliers/meetings but would need real pagination or virtualization at much larger scale.
