-- PIM PM & CM Online - jalankan sekali di Supabase SQL Editor
create extension if not exists pgcrypto;

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  full_name text not null default '',
  role text not null default 'viewer' check (role in ('admin','pm_inspector','technician','supervisor','viewer')),
  is_active boolean not null default true,
  created_at timestamptz not null default now()
);

create table if not exists public.functional_locations (
  id uuid primary key default gen_random_uuid(),
  funloc_code text unique not null,
  description text,
  area text,
  is_active boolean not null default true,
  created_at timestamptz not null default now()
);

create table if not exists public.work_reports (
  id uuid primary key default gen_random_uuid(),
  report_no text unique not null,
  report_type text not null check (report_type in ('PM','CM')),
  pm_reference uuid references public.work_reports(id) on delete set null,
  funloc_id uuid not null references public.functional_locations(id),
  inspection_result text,
  priority text not null default 'Medium',
  description text not null,
  recommendation text,
  corrective_action text,
  spare_part text,
  status text not null,
  assigned_to uuid references public.profiles(id),
  created_by uuid not null references public.profiles(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.report_photos (
  id uuid primary key default gen_random_uuid(),
  report_id uuid not null references public.work_reports(id) on delete cascade,
  storage_path text not null,
  photo_type text not null default 'documentation',
  uploaded_by uuid not null references public.profiles(id),
  created_at timestamptz not null default now()
);

create or replace function public.handle_new_user() returns trigger language plpgsql security definer set search_path=public as $$
begin
 insert into public.profiles(id,full_name,role)
 values(new.id,coalesce(new.raw_user_meta_data->>'full_name',split_part(new.email,'@',1)),'viewer')
 on conflict(id) do nothing;
 return new;
end $$;
drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created after insert on auth.users for each row execute procedure public.handle_new_user();

alter table public.profiles enable row level security;
alter table public.functional_locations enable row level security;
alter table public.work_reports enable row level security;
alter table public.report_photos enable row level security;

create or replace function public.my_role() returns text language sql stable security definer set search_path=public as $$
 select role from public.profiles where id=auth.uid() and is_active=true
$$;

-- Profiles
create policy "profile read authenticated" on public.profiles for select to authenticated using (true);
create policy "profile update admin" on public.profiles for update to authenticated using (public.my_role()='admin') with check (public.my_role()='admin');

-- Functional Location
create policy "funloc read authenticated" on public.functional_locations for select to authenticated using (is_active=true or public.my_role()='admin');
create policy "funloc admin insert" on public.functional_locations for insert to authenticated with check (public.my_role()='admin');
create policy "funloc admin update" on public.functional_locations for update to authenticated using (public.my_role()='admin') with check (public.my_role()='admin');
create policy "funloc admin delete" on public.functional_locations for delete to authenticated using (public.my_role()='admin');

-- Reports
create policy "reports read authenticated" on public.work_reports for select to authenticated using (true);
create policy "reports create operational" on public.work_reports for insert to authenticated with check (created_by=auth.uid() and public.my_role() in ('admin','pm_inspector','technician','supervisor'));
create policy "reports update operational" on public.work_reports for update to authenticated using (public.my_role() in ('admin','technician','supervisor') or created_by=auth.uid()) with check (public.my_role() in ('admin','technician','supervisor') or created_by=auth.uid());
create policy "reports delete admin" on public.work_reports for delete to authenticated using (public.my_role()='admin');

-- Photo metadata
create policy "photos read authenticated" on public.report_photos for select to authenticated using (true);
create policy "photos create operational" on public.report_photos for insert to authenticated with check (uploaded_by=auth.uid() and public.my_role() in ('admin','pm_inspector','technician','supervisor'));
create policy "photos delete admin owner" on public.report_photos for delete to authenticated using (public.my_role()='admin' or uploaded_by=auth.uid());

insert into storage.buckets(id,name,public,file_size_limit,allowed_mime_types)
values('maintenance-photos','maintenance-photos',false,5242880,array['image/jpeg','image/png','image/webp'])
on conflict(id) do update set public=false;

create policy "storage photo read auth" on storage.objects for select to authenticated using (bucket_id='maintenance-photos');
create policy "storage photo upload auth" on storage.objects for insert to authenticated with check (bucket_id='maintenance-photos' and (storage.foldername(name))[1]=auth.uid()::text);
create policy "storage photo delete owner" on storage.objects for delete to authenticated using (bucket_id='maintenance-photos' and (storage.foldername(name))[1]=auth.uid()::text);

insert into public.functional_locations(funloc_code,description,area) values
('IDWY.WY02.PRD.RDY01.EAI01','Bucket Elevator Area','RDS'),
('IDWY.WY02.PRD.RDY01.CNV01','Conveyor Area','RDS'),
('IDWY.WY02.PRD.HML01.MTR01','Hamermill Area','Milling'),
('IDWY.WY02.PRD.HSK01.SEP01','Husking Separator','Husking'),
('IDWY.WY02.PRD.PCK01.PKM01','Packing Machine Area','Packing'),
('IDWY.WY02.GOH.BLD01.CMP01','Air Compressor Area','Utility')
on conflict(funloc_code) do nothing;
