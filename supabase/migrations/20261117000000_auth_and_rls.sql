-- 2026-11-17 登入改用 Supabase Auth＋所有資料表開 RLS
--
-- 正式環境執行順序（詳見 docs/SECURITY_MIGRATION.md）：
--   1. 先部署 Edge Function admin-users
--   2. 執行本檔（執行後舊版前端會讀不到任何資料 → 要緊接著做 3、4）
--   3. 跑 scripts/migrate-users.mjs 把既有帳號搬進 Supabase Auth
--   4. 部署新版前端（GitHub Pages）
--   5. 確認大家都能登入後，執行 supabase/manual/20261117000001_drop_password.sql
--
-- 本檔可重複執行（IF NOT EXISTS / DROP POLICY IF EXISTS）。

-- ── 1. users 以 auth user id 關聯 ─────────────────────────────────────────
alter table public.users add column if not exists auth_id uuid;
do $$ begin
  if not exists (select 1 from pg_constraint where conname = 'users_auth_id_key') then
    alter table public.users add constraint users_auth_id_key unique (auth_id);
  end if;
  if not exists (select 1 from pg_constraint where conname = 'users_auth_id_fkey') then
    alter table public.users add constraint users_auth_id_fkey
      foreign key (auth_id) references auth.users(id) on delete set null;
  end if;
end $$;
-- email 統一小寫比對（Supabase Auth 內部是小寫）
create unique index if not exists users_email_lower_key on public.users (lower(email)) where email is not null and email <> '';

-- ── 2. 取得目前登入者身分的函式（security definer，避免 RLS 自我遞迴）──
-- 停用（active = false）的帳號一律回傳 NULL → 什麼都讀不到。
create or replace function public.app_role() returns text
language sql stable security definer set search_path = public as $$
  select role from public.users where auth_id = auth.uid() and active is true limit 1
$$;
create or replace function public.app_user_name() returns text
language sql stable security definer set search_path = public as $$
  select name from public.users where auth_id = auth.uid() and active is true limit 1
$$;
create or replace function public.app_user_company() returns text
language sql stable security definer set search_path = public as $$
  select company from public.users where auth_id = auth.uid() and active is true limit 1
$$;
revoke all on function public.app_role(), public.app_user_name(), public.app_user_company() from public, anon;
grant execute on function public.app_role(), public.app_user_name(), public.app_user_company() to authenticated;

-- ── 3. 未登入（anon）完全拿不到這幾張表：直接打 REST 會回 401 ───────────────
revoke all on table public.users, public.vendors, public.projects,
                    public.reconcile_batches, public.reconcile_rows from anon;
grant select, insert, update, delete on table public.vendors, public.projects,
                    public.reconcile_batches, public.reconcile_rows to authenticated;
-- users：一般登入者只能讀；新增／修改／停用一律走 Edge Function（service role）
-- 只開放指定欄位（password 欄位在 drop 之前任何登入者都讀不到）
revoke all on table public.users from authenticated;
grant select (id, auth_id, name, email, role, company, active) on table public.users to authenticated;

alter table public.users             enable row level security;
alter table public.vendors           enable row level security;
alter table public.projects          enable row level security;
alter table public.reconcile_batches enable row level security;
alter table public.reconcile_rows    enable row level security;

-- 正式環境原本的「anon 全開」政策（allow_all_*，USING true）一律移除
drop policy if exists allow_all_users             on public.users;
drop policy if exists allow_all_vendors           on public.vendors;
drop policy if exists allow_all_projects          on public.projects;
drop policy if exists allow_all_reconcile_batches on public.reconcile_batches;
drop policy if exists allow_all_reconcile_rows    on public.reconcile_rows;

-- ── 4. users：自己看得到自己；admin／finance 看得到全部（進度頁編輯申請人要名單）
drop policy if exists users_select on public.users;
create policy users_select on public.users for select to authenticated
  using (auth_id = auth.uid() or (select public.app_role()) in ('admin','finance'));

-- ── 5. projects：對齊前端原本的可見範圍（Dashboard / ProjectList 的 visible 過濾）
--   admin、finance：全部
--   manager：同公司
--   clerk：狀態為 申請中／已開立
--   所有人：自己申請的
drop policy if exists projects_select on public.projects;
create policy projects_select on public.projects for select to authenticated
  using (
    (select public.app_role()) in ('admin','finance')
    or ((select public.app_role()) = 'manager' and company = (select public.app_user_company()))
    or ((select public.app_role()) = 'clerk'   and status in ('申請中','已開立'))
    or ((select public.app_role()) is not null and applicant = (select public.app_user_name()))
  );

-- 新增：申請人必須是自己（admin／finance 可代填）
drop policy if exists projects_insert on public.projects;
create policy projects_insert on public.projects for insert to authenticated
  with check (
    (select public.app_role()) in ('admin','finance')
    or ((select public.app_role()) is not null and applicant = (select public.app_user_name()))
  );

-- 修改：看得到的才改得動（前端「標記已開立／已付款」按鈕原本就開放給看得到的人）
drop policy if exists projects_update on public.projects;
create policy projects_update on public.projects for update to authenticated
  using (
    (select public.app_role()) in ('admin','finance')
    or ((select public.app_role()) = 'manager' and company = (select public.app_user_company()))
    or ((select public.app_role()) = 'clerk'   and status in ('申請中','已開立'))
    or ((select public.app_role()) is not null and applicant = (select public.app_user_name()))
  )
  with check ((select public.app_role()) is not null);

-- 刪除：對齊前端 canDelete（finance、admin、manager；manager 只限同公司）
drop policy if exists projects_delete on public.projects;
create policy projects_delete on public.projects for delete to authenticated
  using (
    (select public.app_role()) in ('admin','finance')
    or ((select public.app_role()) = 'manager' and company = (select public.app_user_company()))
  );

-- ── 6. vendors：申請發票時所有角色都要選／新增／更新廠商；刪除限 admin、finance
drop policy if exists vendors_select on public.vendors;
create policy vendors_select on public.vendors for select to authenticated
  using ((select public.app_role()) is not null);
drop policy if exists vendors_insert on public.vendors;
create policy vendors_insert on public.vendors for insert to authenticated
  with check ((select public.app_role()) is not null);
drop policy if exists vendors_update on public.vendors;
create policy vendors_update on public.vendors for update to authenticated
  using ((select public.app_role()) is not null)
  with check ((select public.app_role()) is not null);
drop policy if exists vendors_delete on public.vendors;
create policy vendors_delete on public.vendors for delete to authenticated
  using ((select public.app_role()) in ('admin','finance'));

-- ── 7. 銀行對帳：只有 finance、admin（對齊 ROLE_PERMS 的 csv 頁）──────────
drop policy if exists reconcile_batches_all on public.reconcile_batches;
create policy reconcile_batches_all on public.reconcile_batches for all to authenticated
  using ((select public.app_role()) in ('admin','finance'))
  with check ((select public.app_role()) in ('admin','finance'));
drop policy if exists reconcile_rows_all on public.reconcile_rows;
create policy reconcile_rows_all on public.reconcile_rows for all to authenticated
  using ((select public.app_role()) in ('admin','finance'))
  with check ((select public.app_role()) in ('admin','finance'));
