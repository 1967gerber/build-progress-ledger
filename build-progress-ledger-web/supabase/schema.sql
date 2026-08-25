-- ============================================================
-- BUILD Progress Ledger — Supabase schema
-- Run this once in your Supabase project's SQL Editor.
-- ============================================================

-- One row per teacher. `data` holds everything the app tracks
-- (students, checkpoint entries, custom checkpoints, mastery log)
-- as a single JSON blob — this mirrors the app's existing data
-- shape exactly, so the app code barely has to change.
create table if not exists teacher_data (
  user_id uuid primary key references auth.users(id) on delete cascade,
  teacher_name text not null,
  data jsonb not null default '{"students":[],"entries":[],"customCheckpoints":[],"masteryLog":[]}'::jsonb,
  updated_at timestamptz not null default now()
);

alter table teacher_data enable row level security;

-- Any signed-in teacher can VIEW every other teacher's data.
-- This matches the "All Teachers" visibility the app already has —
-- there's no separate admin role, every account can see everyone.
-- (If you want per-teacher privacy with an admin override later,
-- this is the policy to change — ask and it can be added.)
create policy "Authenticated users can read all teacher data"
  on teacher_data for select
  to authenticated
  using (true);

-- A teacher can only create their OWN row.
create policy "Users can insert their own row"
  on teacher_data for insert
  to authenticated
  with check (auth.uid() = user_id);

-- A teacher can only edit their OWN row — so "viewing" another
-- teacher's data in the app is naturally read-only, enforced here,
-- not just in the UI.
create policy "Users can update their own row"
  on teacher_data for update
  to authenticated
  using (auth.uid() = user_id);

-- Keep updated_at current on every save.
create or replace function set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

create trigger teacher_data_set_updated_at
  before update on teacher_data
  for each row
  execute function set_updated_at();
