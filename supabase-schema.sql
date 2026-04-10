create table if not exists public.tickets (
  id text primary key,
  event_name text not null,
  pass_label text not null,
  pass_number integer not null,
  name text not null,
  email text not null,
  contact text not null,
  pass_type text not null,
  payment_ss_url text not null,
  payment_ss_public_id text,
  verify_url text not null,
  assigned boolean not null default false,
  assigned_time timestamptz,
  mail_sent boolean not null default false,
  mail_status text default 'Pending',
  created_at timestamptz not null default now()
);

create index if not exists tickets_event_name_idx on public.tickets (event_name);
create index if not exists tickets_created_at_idx on public.tickets (created_at);
