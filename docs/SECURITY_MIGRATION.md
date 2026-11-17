# 2026-11-17 登入安全修補：改用 Supabase Auth＋RLS

一句話：密碼不再放在 `users` 表、也不再整表送到瀏覽器；登入改由 Supabase Auth 驗證，所有資料表依登入者角色由資料庫（RLS）把關。

## 原本的問題

- 前端用公開的 anon key `select * from users`，整張表（含明碼 `password`）都進了瀏覽器，帳密比對在前端做。
- 任何人拿 anon key 直接打 REST 就能讀、改、刪所有資料表（帳號、發票、廠商銀行帳號、對帳）。
- 沒設密碼的帳號可以直接登入。

## 改了什麼

| 項目 | 改動 |
|---|---|
| 登入 | Email＋密碼，用 Supabase Auth（`signInWithPassword`）。登入頁有「忘記密碼？寄重設信」。 |
| users 表 | 新增 `auth_id uuid`（關聯 `auth.users`），`password` 欄位任何登入者都讀不到，搬完帳號後整欄刪掉。 |
| 未登入 | 對 users / projects / vendors / reconcile_* 直接打 REST → **401**。 |
| RLS | 見下表，對齊前端原本各角色看得到的範圍。停用帳號一律什麼都讀不到。 |
| 帳號管理 | 新增／改 Email／改角色／停用／設定臨時密碼／寄重設信 都走 Edge Function `admin-users`，只有 admin 能呼叫；service role key 只在 Supabase 端。停用會同時封鎖 Auth 帳號。 |
| 變更自己的密碼 | 右上角「🔑 密碼」改用 `supabase.auth.updateUser`。 |

### 各角色資料權限（RLS）

| 表 | admin | finance | manager | clerk | applicant |
|---|---|---|---|---|---|
| users 讀 | 全部 | 全部 | 自己 | 自己 | 自己 |
| users 寫 | 只能透過 Edge Function | ✗ | ✗ | ✗ | ✗ |
| projects 讀／改 | 全部 | 全部 | 同公司＋自己的 | 申請中／已開立＋自己的 | 自己的 |
| projects 新增 | 可代填 | 可代填 | 只能用自己名字 | 只能用自己名字 | 只能用自己名字 |
| projects 刪 | 全部 | 全部 | 同公司 | ✗ | ✗ |
| vendors | 讀寫刪 | 讀寫刪 | 讀寫 | 讀寫 | 讀寫 |
| 銀行對帳 | 全部 | 全部 | ✗ | ✗ | ✗ |

> 「改」沿用原本前端的行為：看得到的申請都能按「標記已開立／已付款」。如果要收緊（例如申請者不能自己標已付款），再跟我說要怎麼分。

## 正式環境上線步驟（每一步都要你確認後再做）

**事前**
1. Supabase Dashboard → Authentication → URL Configuration
   - Site URL：`https://mamibuy.github.io/mamibuy/`
   - Redirect URLs 加上：`https://mamibuy.github.io/mamibuy/`
2. 決定帳號搬移方式（見下節），若選「寄邀請信」且人數超過 3、4 人，先到 Authentication → SMTP Settings 設定自訂寄信（內建寄信每小時只有個位數額度）。
3. 在 SQL Editor 先跑檢查（只看 email 是否重複／缺漏，不會讀到密碼）：
   ```sql
   select lower(email), count(*) from users where coalesce(email,'') <> '' group by 1 having count(*) > 1;
   select id, name from users where coalesce(email,'') = '';
   select tablename, rowsecurity from pg_tables where schemaname = 'public';
   ```
   重複的 email 要先處理掉；最後一個查詢若有本文件沒提到的表，也要一起決定權限。

**上線（約 10 分鐘，期間系統無法使用，建議下班後做）**
1. 部署 Edge Function：
   ```bash
   supabase login
   supabase link --project-ref kyzyozfdjqlhpcxtzzfp
   supabase functions deploy admin-users
   ```
2. SQL Editor 貼上執行 `supabase/migrations/20261117000000_auth_and_rls.sql`
   （**不要**跑 `20260101000000_baseline_local_only.sql`，那是本機測試用的推測結構）
3. 在自己電腦跑帳號搬移（service role key 到 Project Settings → API 複製，只放在指令的環境變數，不要存檔、不要貼給任何人）：
   ```bash
   SUPABASE_URL=https://kyzyozfdjqlhpcxtzzfp.supabase.co SUPABASE_SERVICE_ROLE_KEY=<貼上> node scripts/migrate-users.mjs --dry-run
   ```
   看過清單沒問題，再改成 `--keep-password` 或 `--invite` 正式執行。
4. 合併到 main → GitHub Actions 自動部署到 GitHub Pages。
5. 用 admin 登入，到帳號管理處理「尚未建立」的帳號（補 Email → 🔑 密碼 → 寄邀請信或設定臨時密碼）。
6. 確認大家都能登入後，SQL Editor 執行 `supabase/manual/20261117000001_drop_password.sql` 把明碼欄位刪掉。
7. 驗證：瀏覽器無痕視窗打 `https://kyzyozfdjqlhpcxtzzfp.supabase.co/rest/v1/users?select=*&apikey=<anon key>` 應回 401。

## 帳號搬移方式（擇一）

> ✅ 2026-11-17 決定：採用 `--keep-password` 沿用舊密碼。上線後請大家登入、右上角「🔑 密碼」自己改一次。

| 方式 | 使用者感受 | 注意 |
|---|---|---|
| `--keep-password` 沿用舊密碼（建議） | 改成用 Email 登入，密碼不變，不用收信 | 舊密碼由腳本在伺服器端直接交給 Supabase Auth 雜湊保存，不會印出；沒密碼或密碼少於 6 字元的人會列在「需要寄邀請信」清單 |
| `--invite` 全部寄邀請信 | 每人收到信、點連結自己設新密碼 | 需要自訂 SMTP；舊密碼直接作廢（建議：舊密碼曾外洩給所有能開 DevTools 的人，這樣最乾淨） |

> 考量：舊密碼過去對所有登入者（甚至未登入者）都是可讀的，嚴格來說應視為已外洩。沿用舊密碼比較順，但安全上建議至少請大家登入後自己改一次密碼，或直接用 `--invite`。

## 本機測試

```bash
colima start                     # 或 Docker Desktop
supabase start -x studio,imgproxy,logflare,vector,supavisor,realtime,storage-api,postgres-meta
supabase db reset                # 套用 migrations＋seed.sql（假帳號、假資料）
SUPABASE_URL=http://127.0.0.1:54321 SUPABASE_SERVICE_ROLE_KEY=<supabase status 的 SERVICE_ROLE_KEY> node scripts/migrate-users.mjs --keep-password
LOCAL_ANON_KEY=<supabase status 的 ANON_KEY> node scripts/test-rls-local.mjs   # 45 項權限測試
npm run dev                      # 需要 .env.local 指到本機（VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY）
```
本機寄出的信在 Mailpit：http://127.0.0.1:54324

## 給勞報系統（mamibuy-sign）串接用

前端登入後拿到的 `session.access_token` 是 Supabase 簽發的 JWT。後端要信任它：
1. 驗簽：用 Supabase 專案的 JWT secret（HS256）或 JWKS `https://kyzyozfdjqlhpcxtzzfp.supabase.co/auth/v1/.well-known/jwks.json`（若專案已啟用非對稱金鑰），或最簡單直接呼叫 `GET /auth/v1/user`（帶該 token＋anon key）由 Supabase 驗證。
2. 取 `sub`（= `auth.users.id`），查 `public.users where auth_id = sub and active` 拿到 name / role / company。
3. 不要信任前端自己傳的角色。
