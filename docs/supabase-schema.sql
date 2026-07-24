-- ARIA Supabase schema — run this in the SQL editor of a fresh project.
-- The Worker talks to these tables with the service-role key (RLS stays enabled
-- with no public policies, so anon clients can't read anything).

-- One row per signed-in Google account. session_id is a random, unguessable id
-- (never derived from the email) that ties this account to its conversations
-- and connected apps across devices.
create table if not exists public.users (
  id         bigint generated always as identity primary key,
  email      text        not null unique,
  name       text,
  picture    text,
  session_id text        not null,
  created_at timestamptz default now()
);

create index if not exists users_session on public.users (session_id);

create table if not exists public.user_apps (
  id            bigint generated always as identity primary key,
  session_id    text        not null,
  app_name      text        not null,
  access_token  text,
  refresh_token text,
  email         text,
  connected_at  timestamptz default now(),
  unique (session_id, app_name)
);

create table if not exists public.conversations (
  id         bigint generated always as identity primary key,
  session_id text        not null,
  role       text        not null,
  content    text,
  persona    text,
  created_at timestamptz default now()
);

create index if not exists conversations_session_created
  on public.conversations (session_id, created_at desc);

alter table public.users         enable row level security;
alter table public.user_apps     enable row level security;
alter table public.conversations enable row level security;
