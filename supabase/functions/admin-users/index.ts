// admin-users：帳號管理需要 service role 的動作都在這裡做，service key 只存在 Supabase 端。
// 呼叫者必須帶自己的登入 JWT，且在 public.users 是 active 的 admin。
//
// POST body: { action, ... }
//   list                                  → 全部帳號（含是否已連結登入）
//   create   { name, email, role, company, active?, password?, redirectTo? }
//            有 password → 直接建立可登入帳號；沒有 → 寄邀請信讓本人設定密碼
//   update   { id, name?, email?, role?, company?, active? }
//            改 email 會同步改 Auth；停用會同時封鎖 Auth 帳號（現有 token 刷新即失效）
//   set_password { id, password }         → 管理員直接設定臨時密碼
//   send_reset   { id, redirectTo? }      → 寄重設密碼信（尚未連結登入的帳號會改寄邀請信）
import { createClient, type SupabaseClient } from "npm:@supabase/supabase-js@2.45.4";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
const ALLOWED_ORIGINS = (Deno.env.get("ALLOWED_ORIGINS") ??
  "https://mamibuy.github.io,http://localhost:5173,http://127.0.0.1:5173").split(",").map((s) => s.trim());
const DEFAULT_REDIRECT = Deno.env.get("DEFAULT_REDIRECT") ?? "https://mamibuy.github.io/mamibuy/";

const ROLES = ["applicant", "manager", "finance", "admin", "clerk"];
const USER_COLS = "id,auth_id,name,email,role,company,active";

function cors(origin: string | null) {
  const allow = origin && ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0];
  return {
    "Access-Control-Allow-Origin": allow,
    "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Vary": "Origin",
  };
}

class HttpError extends Error {
  constructor(public status: number, message: string) { super(message); }
}

function safeRedirect(r: unknown): string {
  if (typeof r !== "string" || !r) return DEFAULT_REDIRECT;
  try {
    const u = new URL(r);
    return ALLOWED_ORIGINS.includes(u.origin) ? u.origin + u.pathname : DEFAULT_REDIRECT;
  } catch { return DEFAULT_REDIRECT; }
}

function normEmail(e: unknown): string {
  const s = typeof e === "string" ? e.trim().toLowerCase() : "";
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s)) throw new HttpError(400, "Email 格式不正確");
  return s;
}

function checkPassword(p: unknown): string {
  if (typeof p !== "string" || p.length < 8) throw new HttpError(400, "密碼至少 8 個字元");
  return p;
}

async function getRow(admin: SupabaseClient, id: unknown) {
  const { data, error } = await admin.from("users").select(USER_COLS).eq("id", id).single();
  if (error || !data) throw new HttpError(404, "找不到此帳號");
  return data;
}

async function invite(admin: SupabaseClient, email: string, redirectTo: string) {
  const { data, error } = await admin.auth.admin.inviteUserByEmail(email, { redirectTo });
  if (error) throw new HttpError(400, "寄送邀請信失敗：" + error.message);
  return data.user!.id;
}

Deno.serve(async (req) => {
  const headers = { ...cors(req.headers.get("origin")), "Content-Type": "application/json" };
  if (req.method === "OPTIONS") return new Response("ok", { headers });
  try {
    if (req.method !== "POST") throw new HttpError(405, "Method not allowed");
    const admin = createClient(SUPABASE_URL, SERVICE_KEY, { auth: { persistSession: false, autoRefreshToken: false } });

    // ── 驗證呼叫者：有效 JWT ＋ active admin ─────────────────────────────
    const jwt = (req.headers.get("authorization") ?? "").replace(/^Bearer\s+/i, "");
    if (!jwt) throw new HttpError(401, "未登入");
    const { data: who, error: whoErr } = await admin.auth.getUser(jwt);
    if (whoErr || !who?.user) throw new HttpError(401, "登入已失效，請重新登入");
    const { data: me } = await admin.from("users").select("id,role,active").eq("auth_id", who.user.id).maybeSingle();
    if (!me || !me.active || me.role !== "admin") throw new HttpError(403, "只有管理員可以管理帳號");

    const body = await req.json().catch(() => ({}));
    const action = body.action;

    if (action === "list") {
      const { data, error } = await admin.from("users").select(USER_COLS).order("id");
      if (error) throw new HttpError(500, error.message);
      return new Response(JSON.stringify({ users: data }), { headers });
    }

    if (action === "create") {
      const name = String(body.name ?? "").trim();
      if (!name) throw new HttpError(400, "請填姓名");
      const email = normEmail(body.email);
      const role = ROLES.includes(body.role) ? body.role : "applicant";
      const company = String(body.company ?? "").trim() || null;
      const active = body.active !== false;
      let authId: string;
      if (body.password) {
        const { data, error } = await admin.auth.admin.createUser({
          email, password: checkPassword(body.password), email_confirm: true,
        });
        if (error) throw new HttpError(400, "建立登入帳號失敗：" + error.message);
        authId = data.user!.id;
      } else {
        authId = await invite(admin, email, safeRedirect(body.redirectTo));
      }
      const { data: row, error } = await admin.from("users")
        .insert({ name, email, role, company, active, auth_id: authId }).select(USER_COLS).single();
      if (error) {
        await admin.auth.admin.deleteUser(authId); // 回滾，避免留下孤兒登入帳號
        throw new HttpError(400, "寫入帳號資料失敗：" + error.message);
      }
      if (!active) await admin.auth.admin.updateUserById(authId, { ban_duration: "876000h" });
      return new Response(JSON.stringify({ user: row }), { headers });
    }

    if (action === "update") {
      const row = await getRow(admin, body.id);
      const patch: Record<string, unknown> = {};
      if (body.name !== undefined) {
        const n = String(body.name).trim(); if (!n) throw new HttpError(400, "請填姓名"); patch.name = n;
      }
      if (body.role !== undefined) {
        if (!ROLES.includes(body.role)) throw new HttpError(400, "角色不正確");
        if (row.id === me.id && body.role !== "admin") throw new HttpError(400, "不能把自己的管理員權限拿掉");
        patch.role = body.role;
      }
      if (body.company !== undefined) patch.company = String(body.company).trim() || null;
      if (body.active !== undefined) {
        if (row.id === me.id && body.active === false) throw new HttpError(400, "不能停用自己");
        patch.active = !!body.active;
      }
      if (body.email !== undefined) {
        const email = normEmail(body.email);
        if (email !== (row.email ?? "").toLowerCase()) {
          if (row.auth_id) {
            const { error } = await admin.auth.admin.updateUserById(row.auth_id, { email, email_confirm: true });
            if (error) throw new HttpError(400, "更新登入 Email 失敗：" + error.message);
          }
          patch.email = email;
        }
      }
      if (row.auth_id && patch.active !== undefined && patch.active !== row.active) {
        const { error } = await admin.auth.admin.updateUserById(row.auth_id, {
          ban_duration: patch.active ? "none" : "876000h",
        });
        if (error) throw new HttpError(400, "更新登入狀態失敗：" + error.message);
      }
      const { data, error } = await admin.from("users").update(patch).eq("id", row.id).select(USER_COLS).single();
      if (error) throw new HttpError(400, error.message);
      return new Response(JSON.stringify({ user: data }), { headers });
    }

    if (action === "set_password") {
      const row = await getRow(admin, body.id);
      const password = checkPassword(body.password);
      if (row.auth_id) {
        const { error } = await admin.auth.admin.updateUserById(row.auth_id, { password });
        if (error) throw new HttpError(400, "設定密碼失敗：" + error.message);
        return new Response(JSON.stringify({ user: row }), { headers });
      }
      // 尚未連結登入 → 建立一個並連結
      const email = normEmail(row.email);
      const { data, error } = await admin.auth.admin.createUser({ email, password, email_confirm: true });
      if (error) throw new HttpError(400, "建立登入帳號失敗：" + error.message);
      const { data: upd, error: e2 } = await admin.from("users").update({ auth_id: data.user!.id })
        .eq("id", row.id).select(USER_COLS).single();
      if (e2) throw new HttpError(400, e2.message);
      return new Response(JSON.stringify({ user: upd }), { headers });
    }

    if (action === "send_reset") {
      const row = await getRow(admin, body.id);
      const redirectTo = safeRedirect(body.redirectTo);
      if (!row.auth_id) {
        const authId = await invite(admin, normEmail(row.email), redirectTo);
        const { data: upd, error } = await admin.from("users").update({ auth_id: authId })
          .eq("id", row.id).select(USER_COLS).single();
        if (error) throw new HttpError(400, error.message);
        return new Response(JSON.stringify({ user: upd, sent: "invite" }), { headers });
      }
      const { error } = await admin.auth.resetPasswordForEmail(normEmail(row.email), { redirectTo });
      if (error) throw new HttpError(400, "寄送重設信失敗：" + error.message);
      return new Response(JSON.stringify({ user: row, sent: "reset" }), { headers });
    }

    throw new HttpError(400, "未知的動作");
  } catch (e) {
    const status = e instanceof HttpError ? e.status : 500;
    const message = e instanceof Error ? e.message : String(e);
    return new Response(JSON.stringify({ error: message }), { status, headers });
  }
});
