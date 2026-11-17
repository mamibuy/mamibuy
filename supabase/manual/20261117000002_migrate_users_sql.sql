-- 2026-11-18 帳號搬進 Supabase Auth（SQL 版，取代需要 service role key 的 scripts/migrate-users.mjs）
-- 在 SQL Editor 執行；必須先跑過 migrations/20261117000000_auth_and_rls.sql（要有 users.auth_id）。
-- 沿用舊密碼：在資料庫內直接以 bcrypt 雜湊寫進 auth.users，不會輸出任何密碼。
-- 可重複執行：已連結（auth_id 不為空）的跳過；同 email 已有 Auth 帳號就直接連結。
-- 沒 email、沒密碼或密碼少於 6 字元的帳號不搬，最後列出名單（只有 id、姓名）。

do $$
declare
  r   record;
  uid uuid;
  em  text;
begin
  for r in
    select id, email, password, active from public.users
    where auth_id is null and coalesce(trim(email), '') <> '' and length(coalesce(password, '')) >= 6
  loop
    em := lower(trim(r.email));
    select id into uid from auth.users where lower(email) = em limit 1;
    if uid is null then
      uid := gen_random_uuid();
      insert into auth.users (
        instance_id, id, aud, role, email, encrypted_password, email_confirmed_at,
        raw_app_meta_data, raw_user_meta_data, created_at, updated_at,
        confirmation_token, recovery_token, email_change_token_new, email_change,
        banned_until
      ) values (
        '00000000-0000-0000-0000-000000000000', uid, 'authenticated', 'authenticated', em,
        extensions.crypt(r.password, extensions.gen_salt('bf', 10)), now(),
        '{"provider":"email","providers":["email"]}'::jsonb, '{}'::jsonb, now(), now(),
        '', '', '', '',
        case when r.active is true then null else now() + interval '100 years' end
      );
      insert into auth.identities (id, user_id, provider_id, identity_data, provider, last_sign_in_at, created_at, updated_at)
      values (gen_random_uuid(), uid, uid::text,
              jsonb_build_object('sub', uid::text, 'email', em, 'email_verified', true),
              'email', now(), now(), now());
    end if;
    update public.users set auth_id = uid, password = null where id = r.id;
  end loop;
end $$;

-- 結果：還沒搬的帳號（要到帳號管理補 Email → 🔑 寄邀請信）
select id, name, '尚未建立登入' as status from public.users where auth_id is null order by id;
