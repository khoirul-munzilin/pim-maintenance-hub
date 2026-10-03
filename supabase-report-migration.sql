-- Tambahan field agar format report menyerupai Maintenance Request
alter table public.work_reports add column if not exists operational_impact text;
alter table public.work_reports add column if not exists failure_cause text;
alter table public.work_reports add column if not exists technician_note text;
alter table public.work_reports add column if not exists test_result text;
alter table public.work_reports add column if not exists verification_note text;
alter table public.work_reports add column if not exists started_at timestamptz;
alter table public.work_reports add column if not exists verification_requested_at timestamptz;
alter table public.work_reports add column if not exists closed_at timestamptz;
-- photo_type yang digunakan: initial, result, verification
