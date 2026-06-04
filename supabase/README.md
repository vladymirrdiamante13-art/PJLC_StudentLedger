# Supabase setup (fix `soa_rows` not found)

## Steps

1. Open [Supabase Dashboard](https://supabase.com/dashboard) → your project **ohbghepqiijkjzqcdaca**.
2. Go to **SQL Editor** → **New query**.
3. Copy and paste the full contents of **`migrate-soa_rows.sql`** from this folder.
4. Click **Run** (you should see “Success”).
5. In the left sidebar, open **Table Editor** and confirm you see:
   - `school_years`
   - `students`
   - `soa_rows`
6. If the app still shows a schema cache error, wait ~30 seconds or go to **Project Settings → API** and use **Reload schema** (if available), then restart `npm run dev`.

## Verify tables exist

Run this in SQL Editor:

```sql
select table_name
from information_schema.tables
where table_schema = 'public'
  and table_name in ('school_years', 'students', 'soa_rows');
```

You should get three rows.
