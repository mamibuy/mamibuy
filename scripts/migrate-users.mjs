#!/usr/bin/env node
// 把 public.users 既有帳號搬進 Supabase Auth，並把 users.auth_id 連起來。
// ⚠️ 需要 service role key，只在自己電腦跑，不要 commit key、不要貼到對話裡。
//
// 用法：
//   SUPABASE_URL=https://xxxx.supabase.co SUPABASE_SERVICE_ROLE_KEY=... node scripts/migrate-users.mjs [模式]
//
// 模式（擇一，預設 --dry-run）：
//   --dry-run        只列出會怎麼處理（只印姓名／email／有無密碼，不印密碼本身）
//   --keep-password  沿用現在的密碼建立登入帳號（大家照舊密碼登入，不用寄信）；
//                    沒有密碼、或密碼太短不符合 Supabase 規則的人，改成「需要寄邀請信」清單
//   --invite         一律寄邀請信，每個人點信裡的連結自己設定新密碼
//                    （注意：Supabase 內建寄信每小時只有個位數額度，人多要先設定自訂 SMTP）
//
// 共通：成功建立登入帳號後，會立刻把該列 users.password 清成 NULL。
//      已經有 auth_id 的列會略過，可以重複執行。
import { createClient } from "@supabase/supabase-js";

const url = process.env.SUPABASE_URL;
const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
const redirectTo = process.env.REDIRECT_TO || "https://mamibuy.github.io/mamibuy/";
const mode = process.argv.includes("--keep-password") ? "keep-password"
  : process.argv.includes("--invite") ? "invite" : "dry-run";

if (!url || !key) {
  console.error("請設定 SUPABASE_URL 與 SUPABASE_SERVICE_ROLE_KEY 環境變數");
  process.exit(1);
}
const admin = createClient(url, key, { auth: { persistSession: false, autoRefreshToken: false } });

const emailOk = (e) => typeof e === "string" && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e.trim());

// 先把 Auth 既有使用者 email → id 建表（重跑時可接回已存在的 Auth 帳號）
async function listAuthUsers() {
  const map = new Map();
  for (let page = 1; ; page++) {
    const { data, error } = await admin.auth.admin.listUsers({ page, perPage: 1000 });
    if (error) throw error;
    for (const u of data.users) if (u.email) map.set(u.email.toLowerCase(), u.id);
    if (data.users.length < 1000) break;
  }
  return map;
}

const { data: rows, error } = await admin.from("users").select("id,name,email,role,active,auth_id,password").order("id");
if (error) { console.error("讀取 users 失敗：", error.message); process.exit(1); }
const authByEmail = await listAuthUsers();

const result = { linked: [], created: [], invited: [], needEmail: [], needInvite: [], failed: [], skipped: [] };
console.log(`模式：${mode}，共 ${rows.length} 個帳號\n`);

for (const r of rows) {
  const label = `#${r.id} ${r.name}${r.active ? "" : "（停用中）"}`;
  if (r.auth_id) { result.skipped.push(label); continue; }
  if (!emailOk(r.email)) { result.needEmail.push(label); continue; }
  const email = r.email.trim().toLowerCase();
  const hasPw = !!(r.password && r.password.length);

  if (mode === "dry-run") {
    const plan = authByEmail.has(email) ? "接回既有登入帳號"
      : hasPw ? "可沿用舊密碼（--keep-password）或寄邀請（--invite）"
      : "沒有密碼 → 需要寄邀請信或由管理員設定密碼";
    console.log(`  ${label} <${email}>：${plan}`);
    continue;
  }

  try {
    let authId = authByEmail.get(email);
    let how = "linked";
    if (!authId) {
      if (mode === "keep-password" && hasPw) {
        const { data, error } = await admin.auth.admin.createUser({ email, password: r.password, email_confirm: true });
        if (error) {
          // 多半是密碼太短不符合規則 → 不建帳號，列入待寄信清單（不印出原因裡的任何密碼）
          result.needInvite.push(`${label} <${email}>（舊密碼不符合規則：${error.message}）`);
          continue;
        }
        authId = data.user.id; how = "created";
      } else if (mode === "keep-password") {
        result.needInvite.push(`${label} <${email}>（原本沒有密碼）`);
        continue;
      } else {
        const { data, error } = await admin.auth.admin.inviteUserByEmail(email, { redirectTo });
        if (error) throw error;
        authId = data.user.id; how = "invited";
      }
    }
    const { error: upErr } = await admin.from("users").update({ auth_id: authId, email, password: null }).eq("id", r.id);
    if (upErr) throw upErr;
    if (!r.active) await admin.auth.admin.updateUserById(authId, { ban_duration: "876000h" });
    result[how].push(`${label} <${email}>`);
  } catch (e) {
    result.failed.push(`${label} <${email}>：${e.message}`);
  }
}

if (mode !== "dry-run") {
  const show = (title, list) => { if (list.length) console.log(`\n${title}（${list.length}）\n  ` + list.join("\n  ")); };
  show("✅ 已用舊密碼建立登入帳號", result.created);
  show("✅ 已寄邀請信", result.invited);
  show("🔗 接回既有登入帳號", result.linked);
  show("⏭ 已連結過、略過", result.skipped);
  show("📧 需要寄邀請信或由管理員設定密碼（到帳號管理頁按「🔑 密碼」）", result.needInvite);
  show("✏️ 沒有有效 Email，請先在帳號管理補上 Email 再處理", result.needEmail);
  show("❌ 失敗", result.failed);
} else if (result.needEmail.length) {
  console.log(`\n✏️ 沒有有效 Email（${result.needEmail.length}）：\n  ` + result.needEmail.join("\n  "));
}
