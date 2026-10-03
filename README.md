# PIM PM & CM Online

## Instalasi singkat
1. Buat project Supabase.
2. Buka SQL Editor, jalankan seluruh isi `supabase-setup.sql`.
3. Buka Authentication > Users, buat user pertama.
4. SQL Editor: `update public.profiles set role='admin', full_name='Nama Admin' where id=(select id from auth.users where email='EMAIL_ADMIN');`
5. Salin Project URL dan Publishable Key ke `config.js`.
6. Upload seluruh isi folder ini ke repository GitHub, aktifkan Pages dari branch `main` root.
7. Di Supabase Authentication > URL Configuration, isi Site URL dengan URL GitHub Pages dan tambahkan URL tersebut ke Redirect URLs.

## Catatan keamanan
- Jangan masukkan Secret Key/Service Role ke `config.js`.
- Frontend hanya menggunakan Publishable Key.
- Akses data dibatasi oleh Supabase Auth dan RLS.
