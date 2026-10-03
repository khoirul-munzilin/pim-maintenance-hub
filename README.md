# Tambahan Report Dokumentasi
1. Jalankan `supabase-report-migration.sql` di SQL Editor.
2. Salin folder `scripts`, `.github`, dan `package.json` ke root repository.
3. Di GitHub Settings > Secrets and variables > Actions, buat: `SUPABASE_URL`, `SUPABASE_SECRET_KEY`, `RESEND_API_KEY`, `REPORT_EMAIL_FROM`, `REPORT_EMAIL_TO`.
4. Foto harus disimpan dengan `photo_type`: `initial`, `result`, atau `verification`.
5. Workflow berjalan setiap hari 07.30 WIB dan bisa dijalankan manual melalui Actions.

Catatan: `SUPABASE_SECRET_KEY` hanya disimpan di GitHub Secrets, tidak boleh ada di `config.js`.
