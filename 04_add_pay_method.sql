-- 新增付款方式欄位到 projects 資料表
-- 在 Supabase SQL Editor 執行此腳本（只需執行一次）
ALTER TABLE projects ADD COLUMN IF NOT EXISTS pay_method TEXT;
