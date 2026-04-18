# Soulwinning Tracker Setup (Windows)

## 1) Install dependencies
1. Open a terminal in `c:\.PROJECTS\soulwinning-tracker`.
2. Run `npm install`.

## 2) Run the desktop app
1. Run `npm run dev:desktop`.
2. The app opens in an Electron window.

## 3) Build a Windows installer (optional)
1. Run `npm run build:desktop`.
2. Find the installer in the `release` folder.

## Data storage & transfer
- Data is stored locally inside the app (no accounts or servers).
- Use **Stats -> Backup & transfer** to export/import a JSON backup.
- Use **Stats -> CSV export** for spreadsheets.

## Optional cloud sync (Supabase)
- If deployed to Cloudflare Pages, you can enable account-based cross-device sync.
- Setup instructions: `docs/SUPABASE_SYNC.md`.
