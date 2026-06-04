-- PJLC: Run this entire script in Supabase Dashboard → SQL Editor → Run
-- Fixes: "Could not find the table 'public.soa_rows' in the schema cache"

-- 1) Students
create table if not exists public.students (
  student_id uuid primary key default gen_random_uuid(),
  student_name text not null,
  grade_level_id text not null,
  created_at timestamptz not null default now()
);

-- 2) Ledger lines (app reads/writes this table as "soa_rows")
create table if not exists public.soa_rows (
  id uuid primary key default gen_random_uuid(),
  student_id uuid not null references public.students (student_id) on delete cascade,
  date date not null,
  description text not null,
  amount numeric(12, 2) not null check (amount > 0),
  entry_type text not null check (entry_type in ('DEBIT', 'CREDIT')),
  or_number text not null default '',
  purpose_key text not null default '',
  meta jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists soa_rows_student_idx
  on public.soa_rows (student_id, date, created_at);

-- 3) Row Level Security (required for anon key access)
alter table public.students enable row level security;
alter table public.soa_rows enable row level security;

drop policy if exists "students_all" on public.students;
drop policy if exists "soa_rows_all" on public.soa_rows;

create policy "students_all"
  on public.students for all
  using (true) with check (true);

create policy "soa_rows_all"
  on public.soa_rows for all
  using (true) with check (true);

-- 4) Tell PostgREST to reload table list (fixes schema cache errors)
notify pgrst, 'reload schema';
