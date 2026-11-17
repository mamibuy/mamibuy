-- 本機測試資料（假資料、假密碼；只在 supabase db reset 時載入本機）
insert into public.users (name, email, role, company, active, password) values
  ('測試管理員', 'admin@test.local',     'admin',     '紳太', true,  'LocalTest-Admin1'),
  ('測試財務',   'finance@test.local',   'finance',   '紳太', true,  'LocalTest-Fin123'),
  ('測試主管',   'manager@test.local',   'manager',   '匯太', true,  'LocalTest-Mgr123'),
  ('測試行政',   'clerk@test.local',     'clerk',     '紳太', true,  'LocalTest-Clk123'),
  ('測試申請者', 'applicant@test.local', 'applicant', '紳太', true,  'LocalTest-App123'),
  ('短密碼者',   'shortpw@test.local',   'applicant', '紳太', true,  '123'),
  ('無密碼者',   'nopw@test.local',      'applicant', '和和研', true, null),
  ('無信箱者',   null,                   'applicant', '紳太', true,  'LocalTest-NoMail'),
  ('已停用者',   'disabled@test.local',  'applicant', '紳太', false, 'LocalTest-Dis123');

insert into public.vendors (name, tax_id, theme) values ('測試廠商A', '12345678', 'invoice'), ('新竹廠商B', '87654321', 'hsinchu');

insert into public.projects (applicant, company, title, amount, tax_amount, status, invoice_date, theme) values
  ('測試申請者', '紳太', '申請者-申請中', 1000, 1050, '申請中', '2026-11-01', 'invoice'),
  ('測試申請者', '紳太', '申請者-已付款', 2000, 2100, '已付款', '2026-10-01', 'invoice'),
  ('測試主管',   '匯太', '主管-已開立',   3000, 3150, '已開立', '2026-11-05', 'invoice'),
  ('無密碼者',   '匯太', '匯太他人-已付款', 4000, 4200, '已付款', '2026-09-01', 'invoice'),
  ('測試財務',   '和和研', '和和研-申請中', 5000, 5250, '申請中', '2026-11-10', 'hsinchu');

insert into public.reconcile_batches (id, file_name, bank_type, saved_at) values ('b1', 'test.csv', 'esun', '2026-11-17');
