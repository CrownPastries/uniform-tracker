# UniTrack v3 — Codebase Context

## Architecture
This is a **Vanilla JS + HTML + CSS** web application. It uses:
- **IndexedDB**: For primary local storage (`UniTrackDB`). Stores `store`, `transactions`, and `syncQueue`.
- **Supabase**: For cloud synchronization (PostgreSQL).
- **Offline-First Data Flow**: 
  - All writes immediately update the local in-memory arrays (`DB_MEMORY`) and IndexedDB.
  - The app attempts to push to Supabase immediately (`CloudSync.pushTransaction` or `pushEmployee`).
  - If the push fails (e.g. offline), the action is saved into the `syncQueue` in IndexedDB.
  - When the app comes back online (`window.addEventListener('online')`), the `syncQueue` is processed automatically, re-attempting the failed API calls.
  - There is also a periodic auto-sync (`CloudSync.pullAll`) every 2 minutes that pulls changes from the server.

## File Structure
- `index.html`: The main UI structure (single-page app layout, modals, sidebars).
- `styles.css`: All styling, using native CSS variables for themes.
- `app.js`: Contains all the application logic, state management, IndexedDB setup, Supabase integration, and DOM manipulation.
- `fix_rls_policies.sql`: Contains the PostgreSQL commands to fix Row Level Security (RLS) on Supabase so that the anonymous client can insert/update records.

## Key Data Structures
- **Employees**: `id`, `firstName`, `lastName`, `employeeId`, `department`, `productionCentre`.
- **Transactions**: `id`, `barcode`, `action`, `date`, `employeeId`, `inferred`.
- **SyncQueue**: Stores offline actions to be processed later. E.g. `{ queueId: "...", type: "pushTransaction", payload: { ... } }`

## Common Gotchas
- The app uses an anonymous Supabase key. By default, Supabase creates tables with strict RLS (Row Level Security). You **must** run `fix_rls_policies.sql` on the database to enable inserts and updates!
- DOM elements are referenced directly via `document.getElementById()`. If you rename an ID in `index.html`, be sure to update `app.js` and `styles.css`.
