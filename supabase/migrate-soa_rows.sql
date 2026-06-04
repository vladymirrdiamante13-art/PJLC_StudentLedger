-- PJLC: Run this entire script in Supabase Dashboard → SQL Editor → Run
-- Fixes: "Could not find the table 'public.soa_rows' in the schema cache"

-- 1) School years
create table if not exists public.school_years (
  id uuid primary key default gen_random_uuid(),
  label text not null,
  is_current boolean not null default false,
  is_archived boolean not null default false,
  archived_at timestamptz,
  created_at timestamptz not null default now()
);

create unique index if not exists school_years_one_current_idx
  on public.school_years (is_current)
  where is_current = true;

insert into public.school_years (label, is_current, is_archived)
select 'Current School Year', true, false
where not exists (select 1 from public.school_years);

-- 2) Students
create table if not exists public.students (
  student_id uuid primary key default gen_random_uuid(),
  school_year_id uuid references public.school_years (id),
  student_name text not null,
  grade_level_id text not null,
  created_at timestamptz not null default now()
);

alter table public.students
  add column if not exists school_year_id uuid references public.school_years (id);

update public.students
set school_year_id = (select id from public.school_years where is_current = true limit 1)
where school_year_id is null;

alter table public.students
  alter column school_year_id set not null;

create index if not exists students_school_year_idx
  on public.students (school_year_id, student_name);

-- 3) Ledger lines (app reads/writes this table as "soa_rows")
create table if not exists public.soa_rows (
  id uuid primary key default gen_random_uuid(),
  school_year_id uuid references public.school_years (id),
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

alter table public.soa_rows
  add column if not exists school_year_id uuid references public.school_years (id);

update public.soa_rows sr
set school_year_id = s.school_year_id
from public.students s
where sr.student_id = s.student_id
  and sr.school_year_id is null;

update public.soa_rows
set school_year_id = (select id from public.school_years where is_current = true limit 1)
where school_year_id is null;

alter table public.soa_rows
  alter column school_year_id set not null;

create index if not exists soa_rows_student_idx
  on public.soa_rows (student_id, date, created_at);

create index if not exists soa_rows_school_year_idx
  on public.soa_rows (school_year_id, student_id, date, created_at);

-- 4) Row Level Security (required for anon key access)
alter table public.school_years enable row level security;
alter table public.students enable row level security;
alter table public.soa_rows enable row level security;

drop policy if exists "school_years_all" on public.school_years;
drop policy if exists "students_all" on public.students;
drop policy if exists "soa_rows_all" on public.soa_rows;

create policy "school_years_all"
  on public.school_years for all
  using (true) with check (true);

create policy "students_all"
  on public.students for all
  using (true) with check (true);

create policy "soa_rows_all"
  on public.soa_rows for all
  using (true) with check (true);

-- 5) API grants for Supabase anon/authenticated keys.
grant usage on schema public to anon, authenticated;
grant all on table public.school_years to anon, authenticated;
grant all on table public.students to anon, authenticated;
grant all on table public.soa_rows to anon, authenticated;

-- 6) Tell PostgREST to reload table list (fixes schema cache errors)
notify pgrst, 'reload schema';
