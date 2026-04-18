# Supabase Account Sync Setup

This app now syncs through **Supabase Auth + per-user snapshots** in **Settings -> Cloud sync**.

## 1) Create a Supabase project
1. Go to Supabase Dashboard.
2. Create a new project.
3. Open **Project Settings -> API** and copy:
   - `Project URL`
   - `anon public key`

## 2) Create the sync table and policies
In Supabase, open **SQL Editor** and run:

```sql
create table if not exists public.user_snapshots (
  user_id uuid primary key references auth.users(id) on delete cascade,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.user_snapshots enable row level security;

create policy "read own snapshot"
on public.user_snapshots
for select
to authenticated
using (auth.uid() = user_id);

create policy "insert own snapshot"
on public.user_snapshots
for insert
to authenticated
with check (auth.uid() = user_id);

create policy "update own snapshot"
on public.user_snapshots
for update
to authenticated
using (auth.uid() = user_id)
with check (auth.uid() = user_id);
```

## 3) Configure email/password auth
1. In Supabase, go to **Authentication -> Providers -> Email**.
2. Ensure email/password auth is enabled.
3. Decide whether email confirmation is required.

## 4) Add environment variables (required)

### Local development
Create `.env.local` in the project root:

```bash
VITE_SUPABASE_URL=YOUR_PROJECT_URL
VITE_SUPABASE_ANON_KEY=YOUR_ANON_PUBLIC_KEY
```

### Cloudflare Pages deployment
1. Go to Cloudflare Dashboard -> **Workers & Pages**.
2. Open your `soulwinningtracker` Pages project.
3. Go to **Settings -> Variables and Secrets**.
4. Under **Environment Variables**, add:
   - `VITE_SUPABASE_URL` = your Supabase Project URL
   - `VITE_SUPABASE_ANON_KEY` = your Supabase anon public key
5. Add them for both **Production** and **Preview** environments.
6. Redeploy the site.

## 5) Use cloud sync in the app
1. Open **Settings -> Cloud sync**.
2. Create account or sign in.
3. While signed in, data auto-saves to cloud after changes.
4. On another device, sign in to auto-pull the latest cloud backup.
5. Use **Download from cloud** anytime to force a manual refresh from cloud.
6. Conflict guard prevents automatic cloud pulls from overwriting unsynced local changes.
7. If a cloud download replaces local data, use **Restore rollback** to recover the previous local snapshot.
