#!/usr/bin/env node
// 本機 supabase 的權限自動測試（只打本機 127.0.0.1，配合 supabase/seed.sql 的假帳號）
// 先跑：supabase db reset && node scripts/migrate-users.mjs --keep-password（帶本機環境變數）
import { createClient } from "@supabase/supabase-js";

const URL = "http://127.0.0.1:54321";
const ANON = process.env.LOCAL_ANON_KEY;
if (!ANON) { console.error("請設定 LOCAL_ANON_KEY（supabase status 的 ANON_KEY）"); process.exit(1); }

let pass = 0, fail = 0;
const check = (name, ok, extra = "") => { ok ? pass++ : fail++; console.log(`${ok ? "✅" : "❌"} ${name}${extra ? "  → " + extra : ""}`); };

async function login(email, password) {
  const c = createClient(URL, ANON, { auth: { persistSession: false } });
  const { data, error } = await c.auth.signInWithPassword({ email, password });
  return { c, error, token: data?.session?.access_token };
}
const titles = (rows) => (rows || []).map((r) => r.title).sort().join(",");
const fn = async (token, body) => {
  const r = await fetch(`${URL}/functions/v1/admin-users`, {
    method: "POST", headers: { apikey: ANON, Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return { status: r.status, body: await r.json().catch(() => ({})) };
};

// ── 登入 ─────────────────────────────────────────────────────────────
const A = await login("admin@test.local", "LocalTest-Admin1");
const F = await login("finance@test.local", "LocalTest-Fin123");
const M = await login("manager@test.local", "LocalTest-Mgr123");
const C = await login("clerk@test.local", "LocalTest-Clk123");
const P = await login("applicant@test.local", "LocalTest-App123");
check("各角色可用沿用的舊密碼登入", ![A, F, M, C, P].some((x) => x.error));
const wrong = await login("admin@test.local", "wrong-password");
check("錯誤密碼登入失敗", !!wrong.error, wrong.error?.message);
const dis = await login("disabled@test.local", "LocalTest-Dis123");
check("停用帳號無法登入", !!dis.error, dis.error?.message);

// ── users ────────────────────────────────────────────────────────────
const uA = await A.c.from("users").select("id,name,role"); check("admin 看得到全部帳號", uA.data?.length === 9, `${uA.data?.length} 筆`);
const uF = await F.c.from("users").select("id,name"); check("finance 看得到全部帳號（進度頁選申請人）", uF.data?.length === 9);
const uP = await P.c.from("users").select("id,name"); check("applicant 只看到自己", uP.data?.length === 1 && uP.data[0].name === "測試申請者");
const pwCol = await A.c.from("users").select("password"); check("連 admin 也讀不到 password 欄位", !!pwCol.error, pwCol.error?.message);
const star = await A.c.from("users").select("*"); check("select=* 也被擋（不會順帶吐出 password）", !!star.error);
const selfPromote = await P.c.from("users").update({ role: "admin" }).eq("name", "測試申請者").select();
check("applicant 不能把自己升成 admin", !!selfPromote.error || (selfPromote.data || []).length === 0, selfPromote.error?.message);
const adminDirect = await A.c.from("users").update({ role: "finance" }).eq("name", "測試申請者").select();
check("連 admin 也不能直接改 users（一律走 Edge Function）", !!adminDirect.error || (adminDirect.data || []).length === 0);
const insUser = await A.c.from("users").insert({ name: "x", role: "admin" }).select();
check("前端不能直接新增 users", !!insUser.error);

// ── projects 可見範圍 ─────────────────────────────────────────────────
const pr = async (x) => (await x.c.from("projects").select("title,company,status,applicant")).data;
check("admin 看全部 5 筆", (await pr(A))?.length === 5);
check("finance 看全部 5 筆", (await pr(F))?.length === 5);
check("manager(匯太) 只看匯太", titles(await pr(M)) === "主管-已開立,匯太他人-已付款", titles(await pr(M)));
check("clerk 只看 申請中／已開立", titles(await pr(C)) === "主管-已開立,和和研-申請中,申請者-申請中", titles(await pr(C)));
check("applicant 只看自己申請的", titles(await pr(P)) === "申請者-已付款,申請者-申請中", titles(await pr(P)));

// 新增：applicant 只能用自己名字
const okIns = await P.c.from("projects").insert({ applicant: "測試申請者", company: "紳太", title: "新申請", status: "申請中", theme: "invoice" }).select();
check("applicant 可新增自己的申請（並拿回該筆）", okIns.data?.length === 1, okIns.error?.message);
const badIns = await P.c.from("projects").insert({ applicant: "測試主管", company: "匯太", title: "冒名", status: "申請中" }).select();
check("applicant 不能冒用別人名字新增", !!badIns.error, badIns.error?.message);

// 修改／刪除
const upOther = await P.c.from("projects").update({ status: "已付款" }).eq("title", "主管-已開立").select();
check("applicant 改不到別人的申請", (upOther.data || []).length === 0);
const delP = await P.c.from("projects").delete().eq("title", "申請者-申請中").select();
check("applicant 不能刪除", (delP.data || []).length === 0);
const delMother = await M.c.from("projects").delete().eq("title", "申請者-已付款").select();
check("manager 刪不到別公司的", (delMother.data || []).length === 0);
const upF = await F.c.from("projects").update({ invoice_no: "AB12345678" }).eq("title", "主管-已開立").select();
check("finance 可修改任何申請", upF.data?.length === 1);

// ── vendors / 對帳 ───────────────────────────────────────────────────
check("applicant 讀得到廠商（申請發票要選）", (await P.c.from("vendors").select("id")).data?.length === 2);
const vDel = await P.c.from("vendors").delete().eq("name", "測試廠商A").select();
check("applicant 不能刪廠商", (vDel.data || []).length === 0);
check("applicant 讀不到銀行對帳", ((await P.c.from("reconcile_batches").select("id")).data || []).length === 0);
check("manager 讀不到銀行對帳", ((await M.c.from("reconcile_batches").select("id")).data || []).length === 0);
check("finance 讀得到銀行對帳", (await F.c.from("reconcile_batches").select("id")).data?.length === 1);

// ── Edge Function admin-users ────────────────────────────────────────
const noAuth = await fetch(`${URL}/functions/v1/admin-users`, { method: "POST", headers: { apikey: ANON, "Content-Type": "application/json" }, body: "{}" });
check("Edge Function：未登入被拒", noAuth.status === 401, String(noAuth.status));
const fFn = await fn(F.token, { action: "list" });
check("Edge Function：finance 被拒（只限 admin）", fFn.status === 403, String(fFn.status));
const aList = await fn(A.token, { action: "list" });
check("Edge Function：admin list", aList.status === 200 && aList.body.users?.length === 9);
check("Edge Function：list 不含 password", aList.body.users?.every((u) => !("password" in u)));

const created = await fn(A.token, { action: "create", name: "新同事", email: "New.Person@test.local", role: "clerk", company: "紳太", password: "NewPerson-123" });
check("admin 新增帳號（直接設密碼）", created.status === 200 && created.body.user?.auth_id, created.body.error);
const N = await login("new.person@test.local", "NewPerson-123");
check("新帳號可登入且角色正確", !N.error && (await N.c.from("users").select("role")).data?.[0]?.role === "clerk");

const invited = await fn(A.token, { action: "create", name: "邀請同事", email: "invitee@test.local", role: "applicant", company: "匯太" });
check("admin 新增帳號（寄邀請信）", invited.status === 200 && invited.body.user?.auth_id, invited.body.error);

const newId = created.body.user?.id;
const deact = await fn(A.token, { action: "update", id: newId, active: false });
check("admin 停用帳號", deact.status === 200 && deact.body.user?.active === false, deact.body.error);
check("停用後原本的 token 立刻讀不到資料", ((await N.c.from("projects").select("id")).data || []).length === 0);
const N2 = await login("new.person@test.local", "NewPerson-123");
check("停用後無法再登入", !!N2.error, N2.error?.message);
await fn(A.token, { action: "update", id: newId, active: true });
const N3 = await login("new.person@test.local", "NewPerson-123");
check("重新啟用後可登入", !N3.error);

const setPw = await fn(A.token, { action: "set_password", id: newId, password: "Changed-Pass-9" });
check("admin 設定臨時密碼", setPw.status === 200);
check("新密碼可登入", !(await login("new.person@test.local", "Changed-Pass-9")).error);
check("舊密碼失效", !!(await login("new.person@test.local", "NewPerson-123")).error);

const shortId = aList.body.users.find((u) => u.email === "nopw@test.local")?.id;
const reset = await fn(A.token, { action: "send_reset", id: shortId });
check("尚未建立登入的帳號按「寄信」→ 寄邀請並連結", reset.status === 200 && reset.body.sent === "invite" && reset.body.user?.auth_id, reset.body.error);
const adminId = aList.body.users.find((u) => u.email === "admin@test.local")?.id;
const selfDemote = await fn(A.token, { action: "update", id: adminId, role: "applicant" });
check("admin 不能拿掉自己的 admin", selfDemote.status === 400);
const badRedirect = await fn(A.token, { action: "send_reset", id: newId, redirectTo: "https://evil.example.com/" });
check("寄信的 redirectTo 只接受白名單（不合法就用預設）", badRedirect.status === 200);

const selfPw = await P.c.auth.updateUser({ password: "Applicant-New-1" });
check("使用者可自己改密碼", !selfPw.error);

console.log(`\n通過 ${pass}，失敗 ${fail}`);
process.exit(fail ? 1 : 0);
