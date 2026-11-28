-- 2026-11-27 發票系統權限收緊
-- 一句話：資料庫自己檢查「誰能改哪些欄位」，不再只靠前端藏按鈕；歸屬改用帳號 id；新增異動紀錄。
--
-- 規則（對照 docs/合約用印系統設計.md 第 3.2 節 1–5 項）：
--   1. projects 欄位分級（trigger projects_guard）
--      - admin、finance：照舊，什麼都能改。
--      - 申請人、主管、行政都「不能」改：申請人、公司、模組（applicant / applicant_id / company / theme）。
--      - 收款欄位（paid / paid_amount / paid_date / bank_fee / commission / commission_no / expense_no / voucher_no）：只有 finance、admin。
--      - 狀態：申請中 → 已開立（同時填發票號碼）只有 finance、admin、clerk；改成已付款或其他狀態只有 finance、admin。
--      - 申請內容（買方、統編、金額、品項、日期、委刊單、付款方式、廠商）：只在「申請中」，
--        而且只有申請人本人、或同公司主管能改；已開立以後只有 finance、admin。
--      - 新增：非財務新增的申請一律是「申請中」、不能帶發票號碼或收款資料。
--   2. 行政（clerk）：看得到申請中／已開立的全部案子，但只能做「標記已開立」，不能改別人的內容。
--   3. vendors 銀行資料（bank / branch / account_name / account）：已有值時只有 finance、admin 能改；
--      空白的可以由申請人補填（申請發票時順手補）。
--   4. projects 新增 applicant_id（對到 users.id），RLS 改用 id 判斷「自己的」；
--      舊資料依姓名回填，同名或找不到的留空，留空的那幾筆暫時照舊用姓名判斷。
--   5. audit_log：projects、vendors、users 每次新增／修改／刪除都記錄誰、何時、改了什麼。
--
-- 在 SQL Editor 用 postgres 身分執行時 auth.uid() 是空的，trigger 不擋（方便你手動修資料），但一樣會記 audit_log。
-- 可以重複執行。

-- ── 0. 目前登入者的 users.id ───────────────────────────────────────────
create or replace function public.app_user_id() returns bigint
language sql stable security definer set search_path = public as $$
  select id from public.users where auth_id = auth.uid() and active is true limit 1
$$;
revoke all on function public.app_user_id() from public, anon;
grant execute on function public.app_user_id() to authenticated;

-- ── 1. projects.applicant_id ──────────────────────────────────────────
alter table public.projects add column if not exists applicant_id bigint;
do $$ begin
  if not exists (select 1 from pg_constraint where conname = 'projects_applicant_id_fkey') then
    alter table public.projects add constraint projects_applicant_id_fkey
      foreign key (applicant_id) references public.users(id) on delete set null;
  end if;
end $$;
create index if not exists projects_applicant_id_idx on public.projects(applicant_id);

-- 回填：只填「姓名在 users 裡剛好一筆」的
update public.projects p
   set applicant_id = u.id
  from (select name, min(id) as id from public.users group by name having count(*) = 1) u
 where p.applicant_id is null and p.applicant = u.name;

-- ── 2. projects 欄位分級 trigger ──────────────────────────────────────
create or replace function public.projects_guard() returns trigger
language plpgsql security definer set search_path = public as $$
declare
  r   text   := public.app_role();
  uid bigint := public.app_user_id();
  is_mine boolean;
begin
  -- SQL Editor／service role（沒有登入者）不擋
  if auth.uid() is null then return new; end if;
  if r is null then raise exception '帳號已停用或不存在'; end if;

  -- ── 新增 ──
  if tg_op = 'INSERT' then
    if r in ('admin','finance') then
      if new.applicant_id is null and (select count(*) from public.users where name = new.applicant) = 1 then
        select id into new.applicant_id from public.users where name = new.applicant;
      end if;
    else
      new.applicant_id := uid;
      if new.status is distinct from '申請中'
         or coalesce(new.invoice_no, '') <> ''
         or coalesce(new.paid, 0) <> 0
         or new.paid_amount is not null or new.paid_date is not null
         or new.bank_fee is not null or new.commission is not null
         or coalesce(new.commission_no, '') <> '' or coalesce(new.expense_no, '') <> '' or coalesce(new.voucher_no, '') <> '' then
        raise exception '新申請只能是「申請中」，發票號碼與收款資料由財務填寫';
      end if;
    end if;
    return new;
  end if;

  -- ── 修改：財務、管理員不限制（改申請人時重新對應 applicant_id）──
  if r in ('admin','finance') then
    if new.applicant is distinct from old.applicant and new.applicant_id is not distinct from old.applicant_id then
      new.applicant_id := null;
      if (select count(*) from public.users where name = new.applicant) = 1 then
        select id into new.applicant_id from public.users where name = new.applicant;
      end if;
    end if;
    return new;
  end if;

  -- 以下是 manager、clerk、applicant
  if new.applicant is distinct from old.applicant or new.applicant_id is distinct from old.applicant_id
     or new.company is distinct from old.company or new.theme is distinct from old.theme then
    raise exception '申請人、公司只有財務可以更改';
  end if;

  if (new.paid, new.paid_amount, new.paid_date, new.bank_fee, new.commission,
      new.commission_no, new.expense_no, new.voucher_no)
     is distinct from
     (old.paid, old.paid_amount, old.paid_date, old.bank_fee, old.commission,
      old.commission_no, old.expense_no, old.voucher_no) then
    raise exception '收款資料只有財務可以填寫';
  end if;

  if new.status is distinct from old.status then
    if not coalesce(r = 'clerk' and old.status = '申請中' and new.status = '已開立', false) then
      raise exception '標記已開立限財務、行政；標記已付款限財務';
    end if;
  end if;

  if new.invoice_no is distinct from old.invoice_no then
    if not coalesce(r = 'clerk' and new.status = '已開立', false) then
      raise exception '發票號碼限財務、行政填寫';
    end if;
  end if;

  if (new.title, new.client_title, new.tax_id, new.amount, new.tax_amount, new.invoice_date,
      new.expected_pay_date, new.order_no, new.pay_method, new.vendor_id, new.items_json)
     is distinct from
     (old.title, old.client_title, old.tax_id, old.amount, old.tax_amount, old.invoice_date,
      old.expected_pay_date, old.order_no, old.pay_method, old.vendor_id, old.items_json) then
    if old.status is distinct from '申請中' then
      raise exception '已開立的申請只有財務可以修改';
    end if;
    -- coalesce：applicant_id 空的時候比較結果是 NULL，不包起來會被當成「不擋」
    is_mine := coalesce(old.applicant_id = uid, false)
            or coalesce(old.applicant_id is null and old.applicant = public.app_user_name(), false);
    if not (is_mine or coalesce(r = 'manager' and old.company = public.app_user_company(), false)) then
      raise exception '只能修改自己的申請';
    end if;
  end if;

  return new;
end $$;
revoke all on function public.projects_guard() from public, anon, authenticated;

drop trigger if exists projects_guard on public.projects;
create trigger projects_guard before insert or update on public.projects
  for each row execute function public.projects_guard();

-- ── 3. RLS：「自己的」改用 applicant_id（沒對到 id 的舊資料照舊用姓名）──
drop policy if exists projects_select on public.projects;
create policy projects_select on public.projects for select to authenticated
  using (
    (select public.app_role()) in ('admin','finance')
    or ((select public.app_role()) = 'manager' and company = (select public.app_user_company()))
    or ((select public.app_role()) = 'clerk'   and status in ('申請中','已開立'))
    or applicant_id = (select public.app_user_id())
    or (applicant_id is null and (select public.app_role()) is not null and applicant = (select public.app_user_name()))
  );

drop policy if exists projects_update on public.projects;
create policy projects_update on public.projects for update to authenticated
  using (
    (select public.app_role()) in ('admin','finance')
    or ((select public.app_role()) = 'manager' and company = (select public.app_user_company()))
    or ((select public.app_role()) = 'clerk'   and status in ('申請中','已開立'))
    or applicant_id = (select public.app_user_id())
    or (applicant_id is null and (select public.app_role()) is not null and applicant = (select public.app_user_name()))
  )
  with check ((select public.app_role()) is not null);
-- projects_insert、projects_delete 不變（insert 仍要求 applicant = 自己姓名，trigger 會補上 applicant_id）

-- ── 4. vendors 銀行資料 ──────────────────────────────────────────────
create or replace function public.vendors_guard() returns trigger
language plpgsql security definer set search_path = public as $$
declare r text := public.app_role();
begin
  if auth.uid() is null or r in ('admin','finance') then return new; end if;
  if r is null then raise exception '帳號已停用或不存在'; end if;
  if (coalesce(old.bank,'')         <> '' and new.bank         is distinct from old.bank)
  or (coalesce(old.branch,'')       <> '' and new.branch       is distinct from old.branch)
  or (coalesce(old.account_name,'') <> '' and new.account_name is distinct from old.account_name)
  or (coalesce(old.account,'')      <> '' and new.account      is distinct from old.account) then
    raise exception '廠商的銀行帳號已經有資料，要修改請找財務';
  end if;
  if new.theme is distinct from old.theme then
    raise exception '廠商所屬模組只有財務可以更改';
  end if;
  return new;
end $$;
revoke all on function public.vendors_guard() from public, anon, authenticated;

drop trigger if exists vendors_guard on public.vendors;
create trigger vendors_guard before update on public.vendors
  for each row execute function public.vendors_guard();

-- ── 5. audit_log ─────────────────────────────────────────────────────
create table if not exists public.audit_log (
  id          bigint generated always as identity primary key,
  at          timestamptz not null default now(),
  table_name  text not null,
  row_id      text,
  action      text not null,          -- INSERT / UPDATE / DELETE
  actor_auth  uuid,                   -- 空的 = SQL Editor 或 Edge Function（service role）
  actor_name  text,
  changes     jsonb                   -- UPDATE 只存有變的欄位 {欄位: [改前, 改後]}；INSERT/DELETE 存整列
);
create index if not exists audit_log_row_idx on public.audit_log(table_name, row_id);
create index if not exists audit_log_at_idx  on public.audit_log(at desc);

alter table public.audit_log enable row level security;
revoke all on table public.audit_log from public, anon, authenticated;
grant select on table public.audit_log to authenticated;
drop policy if exists audit_log_select on public.audit_log;
create policy audit_log_select on public.audit_log for select to authenticated
  using ((select public.app_role()) in ('admin','finance'));

create or replace function public.audit_trigger() returns trigger
language plpgsql security definer set search_path = public as $$
declare
  o jsonb := case when tg_op <> 'INSERT' then to_jsonb(old) end;
  n jsonb := case when tg_op <> 'DELETE' then to_jsonb(new) end;
  diff jsonb;
begin
  -- 保險：就算哪天 users 又有 password 欄位，也不寫進紀錄
  o := o - 'password';
  n := n - 'password';
  if tg_op = 'UPDATE' then
    select jsonb_object_agg(k, jsonb_build_array(o -> k, n -> k)) into diff
      from jsonb_object_keys(n) k
     where (o -> k) is distinct from (n -> k);
    if diff is null then return new; end if;   -- 沒有實際變動就不記
  else
    diff := coalesce(n, o);
  end if;
  insert into public.audit_log (table_name, row_id, action, actor_auth, actor_name, changes)
  values (tg_table_name, coalesce(n, o) ->> 'id', tg_op, auth.uid(),
          (select name from public.users where auth_id = auth.uid() limit 1), diff);
  return coalesce(new, old);
end $$;
revoke all on function public.audit_trigger() from public, anon, authenticated;

drop trigger if exists audit_projects on public.projects;
create trigger audit_projects after insert or update or delete on public.projects
  for each row execute function public.audit_trigger();
drop trigger if exists audit_vendors on public.vendors;
create trigger audit_vendors after insert or update or delete on public.vendors
  for each row execute function public.audit_trigger();
drop trigger if exists audit_users on public.users;
create trigger audit_users after insert or update or delete on public.users
  for each row execute function public.audit_trigger();
