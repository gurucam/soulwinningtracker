# Soulwinning Tracker

Local-first soulwinning logbook with tagged people, session tracking, flexible stats, and optional Supabase account sync.

## Features
- Track sessions by date, people saved, doors knocked, and named people.
- Add people with multiple tags per person.
- Filter stats by date range, week/month/year, and tags.
- Export JSON backups for transfer and CSV for spreadsheets.
- Optional cloud sync (Supabase Auth + per-user snapshot) for cross-device restore.

## Quick start (web)
1. Run `npm install`.
2. Run `npm run dev`.
3. Build with `npm run build` and deploy `dist` to Cloudflare Pages.
4. Configure Supabase env vars (see `.env.example` and `docs/SUPABASE_SYNC.md`).

## Quick start (desktop)
1. Run `npm install`.
2. Run `npm run dev:desktop`.

Full setup details are in `docs/SETUP.md`.

Cloud sync setup details are in `docs/SUPABASE_SYNC.md`.

## Scripts
- `npm run dev` start the Vite dev server (browser)
- `npm run dev:desktop` start the Electron desktop app
- `npm run build` build production assets
- `npm run build:desktop` build the Windows installer
- `npm run preview` preview production build
- `npm run lint` run ESLint
