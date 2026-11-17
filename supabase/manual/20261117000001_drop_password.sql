-- 2026-11-17 移除 users.password 明碼欄位
-- ⚠️ 正式環境：等 scripts/migrate-users.mjs 跑完、確認大家都能用新方式登入後再執行。
-- 執行後舊密碼就永久消失（搬進 Supabase Auth 的是雜湊後的版本，不受影響）。
alter table public.users drop column if exists password;
