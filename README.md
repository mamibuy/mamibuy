# 媽咪拜｜發票開立申請與收款追蹤系統

## 模組
- 🧾 發票申請&對帳（儀表板、申請發票、進度、廠商管理、銀行對帳）
- 🏙️ 新竹申請&對帳（使用申請、全部帳單、費用報告）

## 登入與權限
- 登入用 Supabase Auth（Email＋密碼），資料表全部開 RLS，帳號管理走 Edge Function `admin-users`。
- 上線步驟、權限表、本機測試：見 [docs/SECURITY_MIGRATION.md](docs/SECURITY_MIGRATION.md)

## 部署
```bash
git init && git add .
git commit -m "init"
git remote add origin https://github.com/YOUR/REPO.git
git push -u origin main
```
GitHub → Settings → Pages → Source → GitHub Actions

## 本地
```bash
npm install && npm run dev
```
