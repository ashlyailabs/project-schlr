-- ─────────────────────────────────────────────────────────────────────────────
-- Project Scheduler — Supabase Schema
-- Run this in your Supabase SQL Editor (Database → SQL Editor → New query)
-- ─────────────────────────────────────────────────────────────────────────────

-- Projects
create table projects (
  id         uuid default gen_random_uuid() primary key,
  name       text not null,
  location   text,
  start_date date,
  end_date   date,
  created_at timestamp with time zone default now()
);

-- Tasks
create table tasks (
  id          uuid default gen_random_uuid() primary key,
  project_id  uuid references projects(id) on delete cascade not null,
  sl          integer not null default 1,
  name        text not null,
  description text,
  start_date  date,
  end_date    date,
  status      text not null default 'not-started'
              check (status in ('not-started','in-progress','done','delayed')),
  progress    integer not null default 0
              check (progress >= 0 and progress <= 100),
  created_at  timestamp with time zone default now()
);

-- Highlighted calendar days (per task)
create table highlights (
  id         uuid default gen_random_uuid() primary key,
  project_id uuid references projects(id) on delete cascade not null,
  task_id    uuid references tasks(id) on delete cascade not null,
  date       date not null,
  created_at timestamp with time zone default now(),
  unique(project_id, task_id, date)
);

-- ── Row Level Security (open access — no login required) ───────────────────
alter table projects  enable row level security;
alter table tasks     enable row level security;
alter table highlights enable row level security;

create policy "public_access" on projects   for all using (true) with check (true);
create policy "public_access" on tasks      for all using (true) with check (true);
create policy "public_access" on highlights for all using (true) with check (true);

-- ── Migration (run if upgrading from project-level highlights) ─────────────
-- Drop old unique constraint and add task_id
alter table highlights add column if not exists task_id uuid references tasks(id) on delete cascade;
alter table highlights drop constraint if exists highlights_project_id_date_key;
alter table highlights add constraint if not exists highlights_unique unique (project_id, task_id, date);

-- ── Migration: project location and date range ─────────────────────────────
alter table projects add column if not exists location text;
alter table projects add column if not exists start_date date;
alter table projects add column if not exists end_date date;

-- ── Migration: task dates optional (Gantt uses project range) ───────────────
alter table tasks alter column start_date drop not null;
alter table tasks alter column end_date drop not null;
