import { useState, useMemo, useEffect } from "react";

// ─── 新竹模組共用常數（HsinchuApp 使用）─────────────────────────────────────
const _inp={border:"1.5px solid #E8E8E8",borderRadius:8,padding:"9px 12px",fontSize:14,color:"#1a1a2e",outline:"none",width:"100%",boxSizing:"border-box",background:"#fff"};
const _td={padding:"10px 12px",color:"#444",verticalAlign:"middle"};
const _lbl={fontSize:12,fontWeight:600,color:"#666",marginBottom:6,display:"block"};
const _field={display:"flex",flexDirection:"column"};
const _HCard=({children,style={}})=><div style={{background:"#fff",borderRadius:16,padding:20,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",border:"1px solid #F0F0F0",...style}}>{children}</div>;

// ── 偵測手機版 ───────────────────────────────────────────────────────────────
function useIsMobile() {
  const [isMobile, setIsMobile] = useState(() => window.innerWidth <= 768);
  useEffect(() => {
    const handler = () => setIsMobile(window.innerWidth <= 768);
    window.addEventListener("resize", handler);
    return () => window.removeEventListener("resize", handler);
  }, []);
  return isMobile;
}
function GlobalStyles() {
  return (
    <style>{`
      * { box-sizing: border-box; }
      body { margin: 0; -webkit-tap-highlight-color: transparent; }

      /* 手機底部導覽列 */
      .sidebar-desktop { display: flex !important; }
      .sidebar-mobile  { display: none  !important; }

      /* 內容區在手機要為底部 tab bar 留空間 */
      .content-area { padding-bottom: 16px !important; }

      @media (max-width: 768px) {
        .sidebar-desktop { display: none !important; }
        .sidebar-mobile  { display: flex !important; }
        .content-area    { padding: 14px 12px 110px !important; overflow-x: hidden; }
        .navbar-subtitle { display: none !important; }
        .navbar-pw-btn   { display: none !important; }
        .rwd-grid2       { grid-template-columns: 1fr !important; }
        .rwd-card        { padding: 16px !important; }
        .rwd-modal       { width: calc(100vw - 32px) !important; max-width: 400px; }
        .rwd-h2          { font-size: 18px !important; }
        .rwd-stat-grid   { grid-template-columns: 1fr 1fr !important; }

        /* 手機版取消所有欄位凍結（含標題列）- 強制覆蓋 inline style */
        .rwd-scroll-container th,
        .rwd-scroll-container td,
        .rwd-scroll-container thead th,
        .rwd-scroll-container tbody td {
          position: static !important;
          left: auto !important;
          top: auto !important;
          z-index: 0 !important;
          box-shadow: none !important;
        }

        /* 手機版表格：取消 fixed layout，讓內容撐開寬度以觸發橫向捲軸 */
        .rwd-scroll-container {
          overflow-x: auto !important;
          overflow-y: auto !important;
          -webkit-overflow-scrolling: touch !important;
        }
        .rwd-scroll-container table {
          table-layout: auto !important;
          width: auto !important;
          min-width: 900px !important;
        }
      }

      @media (max-width: 480px) {
        .rwd-stat-grid { grid-template-columns: 1fr !important; }
      }

      /* 平滑捲動 */
      html { scroll-behavior: smooth; }

      /* 所有可捲動容器在 iOS 都滑順 */
      * { -webkit-overflow-scrolling: touch; }

      /* 輸入框在 iOS 不自動放大 */
      input, select, textarea { font-size: 16px; }
      @media (min-width: 769px) {
        input, select, textarea { font-size: 14px; }
      }
    `}</style>
  );
}

// ─── 主題設定 ─────────────────────────────────────────────────────────────────
const THEMES = {
  invoice: { id: "invoice", label: "發票申請&對帳", icon: "🧾", color: "#F2A0A0" },
  hsinchu: { id: "hsinchu", label: "新竹申請&對帳", icon: "🏙️", color: "#82B4D4" },
};

// ─── Supabase 設定 ────────────────────────────────────────────────────────────
const SUPABASE_URL = "https://kyzyozfdjqlhpcxtzzfp.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imt5enlvemZkanFsaHBjeHR6emZwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ3NjU5OTAsImV4cCI6MjA5MDM0MTk5MH0.CEiFol3JeQLJekoKOceJhJYbUlGsGuc0UDpOG78Ru3Y";

const sb = {
  from: (table) => ({
    _table: table, _filters: [], _order: null, _select: "*",
    select(cols="*") { this._select=cols; return this; },
    eq(col,val) { this._filters.push(`${col}=eq.${encodeURIComponent(val)}`); return this; },
    order(col,{ascending=true}={}) { this._order=`${col}.${ascending?"asc":"desc"}`; return this; },
    async _req(method,body) {
      let url=`${SUPABASE_URL}/rest/v1/${this._table}`;
      const params=[...this._filters];
      if(this._order) params.push(`order=${this._order}`);
      if(method==="GET") params.push(`select=${this._select}`);
      if(params.length) url+="?"+params.join("&");
      const res=await fetch(url,{
        method,
        headers:{"apikey":SUPABASE_ANON_KEY,"Authorization":`Bearer ${SUPABASE_ANON_KEY}`,"Content-Type":"application/json","Prefer":"return=representation"},
        ...(body!==undefined?{body:JSON.stringify(body)}:{}),
      });
      const text=await res.text();
      const data=text?JSON.parse(text):null;
      if(!res.ok) throw new Error(data?.message||data?.error||`HTTP ${res.status}`);
      return data;
    },
    async get() { return this._req("GET"); },
    async insert(row) { return this._req("POST",row); },
    async update(row) {
      if(!this._filters.length) throw new Error("update requires filter");
      return this._req("PATCH",row);
    },
    async delete() {
      if(!this._filters.length) throw new Error("delete requires filter");
      return this._req("DELETE");
    },
  }),
};

// ─── DB 欄位轉換 ──────────────────────────────────────────────────────────────
const dbToProject = (r) => ({
  id: r.id, applicant: r.applicant, company: r.company, title: r.title,
  clientTitle: r.client_title||"", taxId: r.tax_id||"",
  amount: Number(r.amount)||0, taxAmount: Number(r.tax_amount)||0,
  status: r.status, invoiceDate: r.invoice_date||"",
  expectedPayDate: r.expected_pay_date||"", orderNo: r.order_no||"",
  invoiceNo: r.invoice_no||"", vendorId: r.vendor_id,
  paid: Number(r.paid)||0,
  paidAmount: r.paid_amount!=null?Number(r.paid_amount):undefined,
  paidDate: r.paid_date||undefined,
  bankFee: r.bank_fee!=null?Number(r.bank_fee):undefined,
  commission: r.commission!=null?Number(r.commission):undefined,
  commissionNo: r.commission_no||"", expenseNo: r.expense_no||"",
  voucherNo: r.voucher_no||"", theme: r.theme||"invoice",
  items: r.items_json ? (() => { try { return JSON.parse(r.items_json); } catch { return []; } })() : [],
});
const projectToDb = (p) => ({
  applicant: p.applicant, company: p.company, title: p.title,
  client_title: p.clientTitle||"", tax_id: p.taxId||"",
  amount: p.amount||0, tax_amount: p.taxAmount||0,
  status: p.status, invoice_date: p.invoiceDate||null,
  expected_pay_date: p.expectedPayDate||null, order_no: p.orderNo||"",
  invoice_no: p.invoiceNo||"", vendor_id: p.vendorId||null,
  paid: p.paid||0, paid_amount: p.paidAmount??null,
  paid_date: p.paidDate||null, bank_fee: p.bankFee??null,
  commission: p.commission??null, commission_no: p.commissionNo||"",
  expense_no: p.expenseNo||"", voucher_no: p.voucherNo||"",
  items_json: p.items ? JSON.stringify(p.items) : null,
  theme: p.theme||"invoice",
});
const dbToVendor = (r) => ({
  id: r.id, name: r.name, taxId: r.tax_id||"", bank: r.bank||"",
  branch: r.branch||"", accountName: r.account_name||"", account: r.account||"",
  theme: r.theme||"invoice",
});
const vendorToDb = (v) => ({
  name: v.name, tax_id: v.taxId||"", bank: v.bank||"", branch: v.branch||"",
  account_name: v.accountName||"", account: v.account||"", theme: v.theme||"invoice",
});
const dbToUser = (r) => ({
  id: r.id, name: r.name, email: r.email||"", role: r.role,
  company: r.company, active: r.active, password: r.password||null,
});

// ─── 常數 ─────────────────────────────────────────────────────────────────────
const ITEMS_CATALOG = [
  {label:"製作費"},{label:"媒體費"},{label:"廣告行銷收入"},{label:"運費"},
  {label:"租金收入"},{label:"顧問費"},{label:"廣告服務收入"},{label:"稿費"},
  {label:"訂閱收入"},{label:"單次購買收入"},{label:"授權費"},
];
const COMPANIES = ["紳太","匯太","和和研"];
const ROLES = {applicant:"申請者",manager:"公司主管",finance:"財務部",admin:"管理員"};
const STATUS_COLOR = {
  "申請中":     {bg:"#FFF3CD",text:"#856404",dot:"#FFC107"},
  "已開立":     {bg:"#CCE5FF",text:"#004085",dot:"#007BFF"},
  "到期未付款": {bg:"#F8D7DA",text:"#721C24",dot:"#DC3545"},
  "已付款":     {bg:"#D4EDDA",text:"#155724",dot:"#28A745"},
};

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const fmt = (n) => new Intl.NumberFormat("zh-TW", { style: "currency", currency: "TWD", maximumFractionDigits: 0 }).format(n);
const toTax = (n) => Math.round(n * 1.05);
const toUntax = (n) => Math.round(n / 1.05);

// ─── COMPONENTS ───────────────────────────────────────────────────────────────

function Badge({ status }) {
  const c = STATUS_COLOR[status] || { bg: "#eee", text: "#333", dot: "#999" };
  return (
    <span style={{ background: c.bg, color: c.text, padding: "3px 10px", borderRadius: 20, fontSize: 12, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 5 }}>
      <span style={{ width: 6, height: 6, borderRadius: "50%", background: c.dot, display: "inline-block" }} />
      {status}
    </span>
  );
}

function Card({ children, style = {} }) {
  return <div style={{ background: "#fff", borderRadius: 16, padding: 24, boxShadow: "0 2px 12px rgba(0,0,0,0.06)", border: "1px solid #F0F0F0", ...style }}>{children}</div>;
}

function StatCard({ label, value, color, icon }) {
  return (
    <div style={{ background: "#fff", borderRadius: 16, padding: "20px 24px", boxShadow: "0 2px 12px rgba(0,0,0,0.06)", border: `1px solid ${color}22`, display: "flex", alignItems: "center", gap: 16 }}>
      <div style={{ width: 48, height: 48, borderRadius: 12, background: `${color}18`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22 }}>{icon}</div>
      <div>
        <div style={{ fontSize: 12, color: "#888", marginBottom: 4, fontWeight: 500 }}>{label}</div>
        <div style={{ fontSize: 20, fontWeight: 700, color: "#1a1a2e" }}>{value}</div>
      </div>
    </div>
  );
}

// ─── INVOICE FORM ─────────────────────────────────────────────────────────────
function InvoiceForm({ currentUser, vendors, onSubmit, prefillData, onClearPrefill }) {
  const [form, setForm] = useState({
    company: currentUser.company === "所有" ? "紳太" : currentUser.company,
    clientTitle: "", taxId: "", invoiceDate: "", expectedPayDate: "", vendorId: "", file: null,
    items: [{ label: "", qty: 1, unitPrice: 0, untaxAmt: 0, taxAmt: 0 }]
  });

  // ── 廠商欄位（自動帶入後可手動補齊）────────────────────────────────────
  const [vSearch, setVSearch] = useState("");          // 統一編號輸入框
  const [vSuggestions, setVSuggestions] = useState([]); // 下拉建議列表
  const [vSelected, setVSelected] = useState(null);    // 選中的廠商（原始物件）
  const [vFields, setVFields] = useState({             // 可編輯的廠商欄位
    name: "", taxId: "", bank: "", branch: "", accountName: "", account: ""
  });
  const [showSuggest, setShowSuggest] = useState(false);
  const [vExpanded, setVExpanded] = useState(false);   // 是否展開廠商資料欄位

  useEffect(() => {
    if (!prefillData) return;
    const pv = vendors.find(v => v.id === prefillData.vendorId);
    setForm(f => ({
      ...f,
      company: prefillData.company || f.company,
      clientTitle: prefillData.clientTitle || "",
      taxId: prefillData.taxId || "",
      expectedPayDate: prefillData.expectedPayDate || "",
      orderNo: prefillData.orderNo || "",
      vendorId: prefillData.vendorId || "",
      invoiceDate: "",
      items: prefillData.items?.length > 0
        ? prefillData.items.map(i => ({ ...i }))
        : [{ label: "", qty: 1, unitPrice: 0, untaxAmt: 0, taxAmt: 0 }],
    }));
    if (pv) {
      setVSearch(pv.taxId || "");
      setVSelected(pv);
      setVFields({ name: pv.name||"", taxId: pv.taxId||"", bank: pv.bank||"", branch: pv.branch||"", accountName: pv.accountName||"", account: pv.account||"" });
      setVExpanded(true);
    }
    if (onClearPrefill) onClearPrefill();
  }, [prefillData]);

  const handleVSearchChange = (val) => {
    setVSearch(val);
    setVSelected(null);
    // 一旦輸入就展開廠商資料欄位讓使用者填寫
    setVExpanded(val.trim().length > 0);
    if (val.trim().length === 0) {
      setVSuggestions([]); setShowSuggest(false);
      setVFields({ name:"", taxId:"", bank:"", branch:"", accountName:"", account:"" });
      setForm(f => ({ ...f, vendorId: "" }));
      return;
    }
    const q = val.trim();
    // 預填統編到 vFields
    setVFields(f => ({ ...f, taxId: q }));
    const matched = vendors
      .filter(v => v.taxId.startsWith(q) || v.taxId.includes(q) || v.name.includes(q))
      .sort((a, b) => (a.taxId.startsWith(q) ? 0 : 1) - (b.taxId.startsWith(q) ? 0 : 1))
      .slice(0, 8);
    setVSuggestions(matched);
    setShowSuggest(matched.length > 0);
  };

  const selectVendor = (v) => {
    setVSelected(v);
    setVSearch(v.taxId);
    setVFields({ name: v.name, taxId: v.taxId, bank: v.bank, branch: v.branch, accountName: v.accountName, account: v.account });
    setVSuggestions([]);
    setShowSuggest(false);
    setVExpanded(true);
    setForm(f => ({ ...f, vendorId: v.id }));
  };

  const missingFields = vExpanded
    ? [["bank","匯款銀行代號"],["branch","分行代號"],["accountName","戶名"],["account","帳號"]].filter(([k]) => !vFields[k])
    : [];

  const updateItem = (i, field, val) => {
    const items = [...form.items];
    items[i] = { ...items[i], [field]: val };
    if (field === "qty" || field === "unitPrice") {
      items[i].untaxAmt = items[i].unitPrice * items[i].qty;
      items[i].taxAmt = toTax(items[i].untaxAmt);
    }
    if (field === "untaxAmt") { items[i].taxAmt = toTax(val); }
    setForm({ ...form, items });
  };

  const total = form.items.reduce((s, i) => s + (i.taxAmt || 0), 0);

  return (
    <div style={{ maxWidth: 800, margin: "0 auto" }}>
      <h2 style={{ fontSize: 22, fontWeight: 700, marginBottom: 24, color: "#1a1a2e" }}>📄 發票開立申請</h2>
      <Card style={{ marginBottom: 16 }}>
        <h3 style={sectionTitle}>基本資訊</h3>
        <div className="rwd-grid2" style={grid2}>
          <div style={field}>
            <label style={lbl}>申請公司</label>
            {currentUser.company !== "所有"
              ? <input style={inp} value={form.company} readOnly />
              : <select style={inp} value={form.company} onChange={e => setForm({ ...form, company: e.target.value })}>
                  {COMPANIES.map(c => <option key={c}>{c}</option>)}
                </select>}
          </div>
          <div style={field}>
            <label style={lbl}>目標發票開立日期</label>
            <input style={inp} type="date" value={form.invoiceDate} onChange={e => setForm({ ...form, invoiceDate: e.target.value })} />
          </div>
          <div style={field}>
            <label style={lbl}>買方</label>
            <input style={inp} value={form.clientTitle} onChange={e => setForm({ ...form, clientTitle: e.target.value })} placeholder="公司抬頭/收款人姓名" />
          </div>
          <div style={field}>
            <label style={lbl}>統一編號/自然人免填</label>
            <input style={inp} value={form.taxId} onChange={e => setForm({ ...form, taxId: e.target.value })} placeholder="8碼統編，自然人免填" maxLength={8} />
          </div>
          <div style={field}>
            <label style={lbl}>預計收款日</label>
            <input style={inp} type="date" value={form.expectedPayDate} onChange={e => setForm({ ...form, expectedPayDate: e.target.value })} />
          </div>
          <div style={field}>
            <label style={lbl}>委刊單編號/其他備註</label>
            <input style={inp} value={form.orderNo || ""} onChange={e => setForm({ ...form, orderNo: e.target.value })} placeholder="請填入委刊單編號或其他備註" />
          </div>
        </div>
      </Card>

      <Card style={{ marginBottom: 16 }}>
        <h3 style={sectionTitle}>開立項目</h3>
        {form.items.map((item, i) => (
          <div key={i} style={{ display: "grid", gridTemplateColumns: "2fr 80px 120px 120px 120px 36px", gap: 8, marginBottom: 8, alignItems: "center" }}>
            <select style={inp} value={item.label} onChange={e => updateItem(i, "label", e.target.value)}>
              <option value="">選擇項目</option>
              {ITEMS_CATALOG.map(c => <option key={c.label}>{c.label}</option>)}
            </select>
            <input style={{ ...inp, textAlign: "center" }} type="text" inputMode="numeric" value={item.qty || ""} onChange={e => { const v = parseInt(e.target.value.replace(/\D/g,""),10); updateItem(i, "qty", isNaN(v) ? 0 : v); }} placeholder="數量" />
            <input style={{ ...inp, textAlign: "right" }} type="text" inputMode="numeric" value={item.unitPrice || ""} onChange={e => { const v = parseInt(e.target.value.replace(/\D/g,""),10); updateItem(i, "unitPrice", isNaN(v) ? 0 : v); }} placeholder="單價" />
            <div style={{ ...inp, background: "#F8F9FA", textAlign: "right", color: "#555", fontSize: 13 }}>{item.untaxAmt ? item.untaxAmt.toLocaleString() : "-"}</div>
            <div style={{ ...inp, background: "#E8F5E9", textAlign: "right", color: "#2E7D32", fontWeight: 700, fontSize: 13 }}>{item.taxAmt ? item.taxAmt.toLocaleString() : "-"}</div>
            <button onClick={() => setForm({ ...form, items: form.items.filter((_, j) => j !== i) })} style={{ border: "none", background: "#fee", color: "#c33", borderRadius: 8, cursor: "pointer", height: 38, fontSize: 16 }}>×</button>
          </div>
        ))}
        <div style={{ display: "grid", gridTemplateColumns: "2fr 80px 120px 120px 120px 36px", gap: 8, marginBottom: 12, paddingLeft: 0 }}>
          <div style={{ color: "#888", fontSize: 11, textAlign: "center" }}>項目名稱</div>
          <div style={{ color: "#888", fontSize: 11, textAlign: "center" }}>數量</div>
          <div style={{ color: "#888", fontSize: 11, textAlign: "center" }}>單價</div>
          <div style={{ color: "#888", fontSize: 11, textAlign: "center" }}>未稅金額</div>
          <div style={{ color: "#888", fontSize: 11, textAlign: "center" }}>含稅金額</div>
          <div />
        </div>
        <button onClick={() => setForm({ ...form, items: [...form.items, { label: "", qty: 1, unitPrice: 0, untaxAmt: 0, taxAmt: 0 }] })}
          style={{ border: "2px dashed #C0C0C0", background: "none", color: "#888", padding: "8px 20px", borderRadius: 8, cursor: "pointer", width: "100%", fontSize: 13 }}>
          ＋ 新增項目
        </button>
        <div style={{ textAlign: "right", marginTop: 12, fontSize: 18, fontWeight: 700, color: "#2E7D32" }}>
          合計（含稅）：{fmt(total)}
        </div>
      </Card>

      <Card style={{ marginBottom: 16 }}>
        <h3 style={sectionTitle}>廠商付款資料</h3>

        {/* ── 統一編號搜尋框 ── */}
        <div style={{ position: "relative", marginBottom: vExpanded ? 20 : 0 }}>
          <label style={lbl}>輸入統一編號搜尋廠商</label>
          <div style={{ position: "relative" }}>
            <input
              style={{ ...inp, paddingRight: 36 }}
              value={vSearch}
              onChange={e => handleVSearchChange(e.target.value)}
              onFocus={() => vSuggestions.length > 0 && setShowSuggest(true)}
              onBlur={() => setTimeout(() => setShowSuggest(false), 200)}
              placeholder="輸入統編數字或公司名稱…"
              autoComplete="off"
            />
            {vSearch && (
              <span
                onMouseDown={e => { e.preventDefault(); handleVSearchChange(""); setVExpanded(false); }}
                style={{ position:"absolute", right:10, top:"50%", transform:"translateY(-50%)", cursor:"pointer", color:"#aaa", fontSize:20, lineHeight:1, userSelect:"none" }}>×</span>
            )}
          </div>

          {/* ── 建議下拉 ── */}
          {showSuggest && vSuggestions.length > 0 && (
            <div style={{ position:"absolute", top:"calc(100% + 4px)", left:0, right:0, background:"#fff", border:"1.5px solid #C8DCF7", borderRadius:10, boxShadow:"0 8px 28px rgba(0,0,0,0.12)", zIndex:100, maxHeight:280, overflowY:"auto" }}>
              <div style={{ padding:"8px 14px 6px", fontSize:11, color:"#999", fontWeight:600, borderBottom:"1px solid #F0F0F0" }}>
                找到 {vSuggestions.length} 筆廠商資料
              </div>
              {vSuggestions.map(v => (
                <div key={v.id}
                  onMouseDown={e => { e.preventDefault(); selectVendor(v); }}
                  style={{ padding:"10px 14px", cursor:"pointer", borderBottom:"1px solid #F8F8F8" }}
                  onMouseEnter={e => e.currentTarget.style.background="#EEF6FF"}
                  onMouseLeave={e => e.currentTarget.style.background="#fff"}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                    <span style={{ fontWeight:600, fontSize:14, color:"#1a1a2e" }}>{v.name}</span>
                    <span style={{ fontSize:12, color:"#007BFF", fontWeight:700 }}>{v.taxId}</span>
                  </div>
                  <div style={{ fontSize:12, color:"#888", marginTop:3 }}>
                    {v.bank ? `${v.bank}` : <span style={{color:"#ccc"}}>無銀行資料</span>}
                    {v.account ? `　帳號：${v.account}` : ""}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* ── 廠商資料欄位（輸入統編後一律顯示，可核對、可補齊） ── */}
        {vExpanded && (
          <div style={{ background: vSelected ? "#F8FBFF" : "#FAFAFA", border: `1.5px solid ${vSelected ? "#BFD9F7" : "#E8E8E8"}`, borderRadius:12, padding:"16px 18px" }}>
            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:14 }}>
              {vSelected
                ? <span style={{ fontSize:13, fontWeight:700, color:"#1565C0" }}>✓ 已從資料庫帶入，請核對以下資料</span>
                : <span style={{ fontSize:13, fontWeight:700, color:"#888" }}>查無此統編的廠商紀錄，請手動填入以下資料</span>
              }
            </div>
            <div className="rwd-grid2" style={grid2}>
              {[
                ["name","公司名稱", true],
                ["taxId","統一編號", false],
                ["bank","匯款銀行代號", true],
                ["branch","分行代號", true],
                ["accountName","戶名", true],
                ["account","帳號（勿加標點符號如：- / . ）", true],
              ].map(([k, l, editable]) => (
                <div key={k} style={field}>
                  <label style={{ ...lbl, color: "#666" }}>{l}</label>
                  <input
                    style={{
                      ...inp,
                      background: !editable ? "#F0F0F0" : "#fff",
                      color: !editable ? "#888" : "#1a1a2e",
                    }}
                    value={vFields[k]}
                    onChange={e => editable && setVFields(f => ({ ...f, [k]: e.target.value }))}
                    readOnly={!editable}
                    placeholder={editable ? `請填入${l}` : ""}
                  />
                </div>
              ))}
            </div>
          </div>
        )}

        {!vExpanded && (
          <div style={{ fontSize:13, color:"#bbb", paddingTop:4 }}>請輸入統一編號或公司名稱來搜尋廠商資料。</div>
        )}
      </Card>

      <Card style={{ marginBottom: 16 }}>
        <h3 style={sectionTitle}>上傳專案憑證</h3>
        <label style={{ display: "block", border: "2px dashed #C0C0C0", borderRadius: 10, padding: "24px", textAlign: "center", cursor: "pointer", color: "#888", fontSize: 13 }}>
          <div style={{ fontSize: 28, marginBottom: 8 }}>📎</div>
          點擊上傳或拖曳檔案（PDF, PNG, JPG）
          <input type="file" style={{ display: "none" }} onChange={e => setForm({ ...form, file: e.target.files[0] })} />
        </label>
        {form.file && <div style={{ marginTop: 8, fontSize: 13, color: "#2E7D32" }}>✓ 已選擇：{form.file.name}</div>}
      </Card>

      <button onClick={() => onSubmit({ ...form, total, vendorId: form.vendorId, vFields, vSelected })}
        style={{ width: "100%", padding: "14px", background: "linear-gradient(135deg,#1a1a2e,#16213e)", color: "#fff", border: "none", borderRadius: 12, fontSize: 16, fontWeight: 700, cursor: "pointer" }}>
        提交申請 →
      </button>
    </div>
  );
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({ projects, currentUser }) {
  const [selectedMonth, setSelectedMonth] = useState("all");

  const visible = projects.filter(p => {
    if (currentUser.role === "admin" || currentUser.role === "finance") return true;
    if (currentUser.role === "manager") return p.company === currentUser.company;
    return p.applicant === currentUser.name;
  });

  // ── 取出所有月份（YYYY-MM），排序由新到舊 ─────────────────────────────
  const months = useMemo(() => {
    const set = new Set(visible.map(p => (p.invoiceDate || "").slice(0, 7)).filter(Boolean));
    return ["all", ...Array.from(set).sort((a, b) => b.localeCompare(a))];
  }, [visible]);

  const filtered = selectedMonth === "all"
    ? visible
    : visible.filter(p => (p.invoiceDate || "").startsWith(selectedMonth));

  const stats = {
    pending:  filtered.filter(p => p._effectiveStatus === "申請中").reduce((s, p) => s + p.taxAmount, 0),
    issued:   filtered.filter(p => ["已開立","已付款"].includes(p._effectiveStatus)).reduce((s, p) => s + p.taxAmount, 0),
    paid:     filtered.filter(p => p._effectiveStatus === "已付款").reduce((s, p) => s + p.taxAmount, 0),
    overdue:  filtered.filter(p => p._effectiveStatus === "到期未付款").reduce((s, p) => s + p.taxAmount, 0),
  };

  const fmtMonth = (m) => {
    if (m === "all") return "全部月份";
    const [y, mo] = m.split("-");
    return `${y} 年 ${parseInt(mo)} 月`;
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24, flexWrap: "wrap", gap: 12 }}>
        <h2 style={{ fontSize: 22, fontWeight: 700, color: "#1a1a2e" }}>📊 儀表板總覽</h2>
        {/* 月份切換 */}
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
          <span style={{ fontSize: 12, color: "#888", marginRight: 4 }}>發票開立月份：</span>
          {months.map(m => (
            <button key={m} onClick={() => setSelectedMonth(m)}
              style={{
                padding: "5px 14px", borderRadius: 20, border: "none", cursor: "pointer", fontSize: 13,
                background: selectedMonth === m ? "#1a1a2e" : "#F0F0F0",
                color: selectedMonth === m ? "#fff" : "#555",
                fontWeight: selectedMonth === m ? 700 : 400,
              }}>
              {fmtMonth(m)}
            </button>
          ))}
        </div>
      </div>

      {/* 統計卡 */}
      <div className="rwd-stat-grid" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 16, marginBottom: 32 }}>
        <StatCard label="預計開立金額" value={fmt(stats.pending)} color="#FF9800" icon="📋" />
        <StatCard label="已開立發票金額" value={fmt(stats.issued)} color="#2196F3" icon="🧾" />
        <StatCard label="已付款金額" value={fmt(stats.paid)} color="#4CAF50" icon="✅" />
        <StatCard label="到期未付款金額" value={fmt(stats.overdue)} color="#F44336" icon="⚠️" />
      </div>

      {/* 進度追蹤 */}
      <Card>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h3 style={{ ...sectionTitle, marginBottom: 0 }}>專案進度追蹤</h3>
          <span style={{ fontSize: 12, color: "#888" }}>
            {selectedMonth === "all" ? `共 ${filtered.length} 筆` : `${fmtMonth(selectedMonth)} · ${filtered.length} 筆`}
          </span>
        </div>
        {filtered.length === 0 && <div style={{ color: "#aaa", textAlign: "center", padding: 32 }}>此月份無資料</div>}
        {filtered.map(p => {
          const es = p._effectiveStatus;
          const pct = es === "已付款" ? 100 : es === "到期未付款" ? 66 : es === "已開立" ? 66 : es === "申請中" ? 33 : 10;
          const colors = { "申請中": "#FF9800", "已開立": "#2196F3", "到期未付款": "#DC3545", "已付款": "#28A745" };
          return (
            <div key={p.id} style={{ marginBottom: 20, paddingBottom: 20, borderBottom: "1px solid #F0F0F0" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <div>
                  <span style={{ fontWeight: 600, color: "#1a1a2e" }}>{p.title}</span>
                  <span style={{ marginLeft: 10, fontSize: 12, color: "#888" }}>
                    {p.company} · {p.applicant} · {p.invoiceDate}
                    {p.expectedPayDate && <span style={{ color: es === "到期未付款" ? "#DC3545" : "#aaa", marginLeft: 6 }}>
                      （預計收款：{p.expectedPayDate}）
                    </span>}
                  </span>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <span style={{ fontWeight: 700, color: "#1a1a2e" }}>{fmt(p.taxAmount)}</span>
                  <Badge status={es} />
                </div>
              </div>
              <div style={{ background: "#F0F0F0", borderRadius: 100, height: 8, overflow: "hidden" }}>
                <div style={{ width: `${pct}%`, height: "100%", background: colors[es] || "#888", borderRadius: 100, transition: "width 0.5s" }} />
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 4, fontSize: 11, color: "#AAA" }}>
                <span>申請中</span><span>發票開立</span><span>款項收取</span>
              </div>
            </div>
          );
        })}
      </Card>
    </div>
  );
}

// ─── BANK RECONCILE ───────────────────────────────────────────────────────────
function CSVReconcile({ projects, vendors, setProjects }) {
  const isMobile = useIsMobile();
  // ── 當前工作區（正在處理的檔案批次）────────────────────────────────────
  const [workingBatch, setWorkingBatch] = useState([]);
  // workingBatch = [{ fileName, bankType, matched: [...] }, ...]

  // ── 歷史資料庫 ──────────────────────────────────────────────────────────
  const [history, setHistory] = useState([]);           // 從 storage 載入的歷史批次
  const [activeHistory, setActiveHistory] = useState(null); // 正在查看的歷史 id
  const [storageReady, setStorageReady] = useState(false);

  const [loading, setLoading] = useState(false);
  const [loadingFiles, setLoadingFiles] = useState([]);  // 正在解析的檔名列表
  const [error, setError] = useState("");
  const [activeTab, setActiveTab] = useState("working"); // "working" | "history"

  // ── 從 storage 載入歷史 ───────────────────────────────────────────────
  useEffect(() => {
    const load = async () => {
      try {
        const batchesRaw = await sb.from("reconcile_batches").select("*").order("saved_at",{ascending:false}).get();
        if (batchesRaw?.length > 0) {
          const rowsRaw = await sb.from("reconcile_rows").select("*").get();
          const rowsByBatch = {};
          (rowsRaw||[]).forEach(r => {
            if (!rowsByBatch[r.batch_id]) rowsByBatch[r.batch_id]=[];
            rowsByBatch[r.batch_id].push(r);
          });
          const batches = batchesRaw.map(b => ({
            id: b.id, fileName: b.file_name, bankType: b.bank_type,
            savedAt: b.saved_at, savedDate: b.saved_date, parseError: b.parse_error,
            matched: (rowsByBatch[b.id]||[]).sort((a,z)=>a.row_index-z.row_index).map(r => ({
              date: r.date_val, amount: r.amount!=null?Number(r.amount):null,
              summary: r.summary||"", diff: r.diff!=null?Number(r.diff):null,
              diffNote: r.diff_note||"", fee: r.fee!=null?Number(r.fee):null,
              feeConfirmed: r.fee_confirmed||false,
              otherDiff: r.other_diff!=null?Number(r.other_diff):null,
              voucherNo: r.voucher_no||"", expenseNo: r.expense_no||"",
              match: r.match_id ? {id:r.match_id,title:r.match_title||"",taxAmount:r.match_tax_amount!=null?Number(r.match_tax_amount):null,clientTitle:r.match_client_title||"",orderNo:r.match_order_no||""} : null,
            })),
          }));
          setHistory(batches);
        }
      } catch(err) { console.warn("載入對帳歷史失敗:",err.message); }
      setStorageReady(true);
    };
    load();
  }, []);

  // ── 比對邏輯（不變）──────────────────────────────────────────────────
  const reconcile = (txRows, type) => {
    return txRows.map(tx => {
      let matchedProject = null;
      let matchReason = "";

      if (type === "yushan") {
        if (!matchedProject && tx.note) {
          const noteClean = tx.note.replace(/\s/g, "");
          matchedProject = projects.find(p => {
            const title = (p.clientTitle || "").replace(/\s/g, "");
            return title && noteClean.includes(title);
          });
          if (matchedProject) matchReason = "備註×公司抬頭";
        }
        if (!matchedProject && tx.bankAccount) {
          const [txBank, txAcc] = tx.bankAccount.split("/").map(s => s?.trim().replace(/^0+/, ""));
          matchedProject = projects.find(p => {
            const v = vendors.find(v => v.id === p.vendorId);
            if (!v) return false;
            const vBank = (v.bank || "").replace(/[^0-9]/g, "").replace(/^0+/, "");
            const vAcc = (v.account || "").replace(/\D/g, "").replace(/^0+/, "");
            const txAccClean = (txAcc || "").replace(/^0+/, "");
            return (vBank && txBank && vBank === txBank) || (vAcc && txAccClean && vAcc === txAccClean);
          });
          if (matchedProject) matchReason = "銀行代號/帳號";
        }
      }

      if (type === "cathay") {
        if (!matchedProject && tx.note) {
          const noteClean = tx.note.replace(/\s/g, "").replace(/[Ａ-Ｚａ-ｚ０-９]/g, c =>
            String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
          matchedProject = projects.find(p => {
            const title = (p.clientTitle || "").replace(/\s/g, "");
            return title && noteClean.includes(title);
          });
          if (matchedProject) matchReason = "備註×公司抬頭";
        }
        if (!matchedProject && tx.bankAccount) {
          const rawAcc = tx.bankAccount.replace(/\s/g, "");
          matchedProject = projects.find(p => {
            const v = vendors.find(v => v.id === p.vendorId);
            if (!v) return false;
            const vBank = (v.bank || "").replace(/[^0-9]/g, "").replace(/^0+/, "").padStart(3, "0");
            const vAcc = (v.account || "").replace(/\D/g, "");
            return (vBank && rawAcc.includes(vBank)) || (vAcc && rawAcc.includes(vAcc.slice(-6)));
          });
          if (matchedProject) matchReason = "對方帳號";
        }
      }

      let closestProject = null;
      let closestDiff = Infinity;
      if (!matchedProject) {
        projects.forEach(p => {
          const diff = Math.abs(p.taxAmount - tx.amount);
          if (diff < closestDiff) { closestDiff = diff; closestProject = p; }
        });
        if (closestDiff <= 500) {
          matchedProject = closestProject;
          matchReason = `金額最近似（差 ${closestDiff}，請人工確認）`;
        }
      }

      const diff = matchedProject ? tx.amount - matchedProject.taxAmount : null;
      return { ...tx, match: matchedProject, diff, matchReason };
    });
  };

  // ── SheetJS 載入 ──────────────────────────────────────────────────────
  const loadSheetJS = () => new Promise((res, rej) => {
    if (window.XLSX) return res();
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s.onload = res;
    s.onerror = () => rej(new Error("無法載入 SheetJS，請確認網路連線"));
    document.head.appendChild(s);
  });

  const readAsArrayBuffer = (file) => new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = (e) => res(e.target.result);
    reader.onerror = () => rej(new Error("檔案讀取失敗"));
    reader.readAsArrayBuffer(file);
  });

  const parseYushan = async (file) => {
    const ab = await readAsArrayBuffer(file);
    const wb = window.XLSX.read(new Uint8Array(ab), { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const headerIdx = rows.findIndex(r => r && String(r[0]).trim() === "序號");
    if (headerIdx < 0) throw new Error("找不到玉山對帳單標題列（序號）");
    const txRows = [];
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r[1] == null) continue;
      const credit = parseFloat(r[6]) || 0;
      if (credit <= 0) continue;
      txRows.push({ date: String(r[1] || ""), desc: String(r[4] || ""), note: String(r[8] || ""), bankAccount: String(r[9] || ""), amount: credit });
    }
    if (txRows.length === 0) throw new Error("玉山對帳單中找不到任何存入記錄");
    return txRows;
  };

  const parseCathay = async (file) => {
    const ab = await readAsArrayBuffer(file);
    const wb = window.XLSX.read(new Uint8Array(ab), { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const headerIdx = rows.findIndex(r => r && String(r[0]).trim() === "序號");
    if (headerIdx < 0) throw new Error("找不到國泰對帳單標題列（序號）");
    const txRows = [];
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r[0] == null || isNaN(parseFloat(r[0]))) continue;
      const credit = parseFloat(r[3]) || 0;
      if (credit <= 0) continue;
      txRows.push({ date: String(r[1] || ""), desc: String(r[5] || ""), note: String(r[7] || ""), bankAccount: String(r[6] || ""), amount: credit });
    }
    if (txRows.length === 0) throw new Error("國泰對帳單中找不到任何存入記錄");
    return txRows;
  };

  // ── 多檔上傳 ─────────────────────────────────────────────────────────
  const handleFiles = async (e) => {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    setError(""); setLoading(true);
    setLoadingFiles(files.map(f => f.name));
    try {
      await loadSheetJS();
      const newBatches = [];
      for (const file of files) {
        try {
          const name = file.name.toLowerCase();
          const isYushan = name.includes("玉山") || name.endsWith(".xls");
          const type = isYushan ? "yushan" : "cathay";
          const txRows = isYushan ? await parseYushan(file) : await parseCathay(file);
          newBatches.push({ fileName: file.name, bankType: type, matched: reconcile(txRows, type) });
        } catch (err) {
          newBatches.push({ fileName: file.name, bankType: null, matched: [], parseError: err.message });
        }
      }
      setWorkingBatch(prev => [...prev, ...newBatches]);
      setActiveTab("working");
    } catch (err) {
      setError("❌ 解析失敗：" + err.message);
    }
    setLoading(false);
    setLoadingFiles([]);
    e.target.value = "";
  };

  // ── 歷史編輯模式 ─────────────────────────────────────────────────────
  const [editingHistoryId, setEditingHistoryId] = useState(null); // 正在編輯的歷史 id
  const [selectedRows, setSelectedRows] = useState({});           // { batchKey_ri: true }

  // ── 儲存至資料庫（新增 / 覆寫）──────────────────────────────────────
  const saveBatch = async (batch) => {
    const id = batch.id || `reconcile:${Date.now()}_${Math.random().toString(36).slice(2,7)}`;
    // 序列化前精簡 match 物件（避免儲存完整 project 造成資料過大）
    const safeMatched = batch.matched.map(r => ({
      ...r,
      match: r.match ? {
        id: r.match.id,
        title: r.match.title,
        taxAmount: r.match.taxAmount,
        clientTitle: r.match.clientTitle || "",
        orderNo: r.match.orderNo || "",
      } : null,
    }));
    const record = {
      ...batch,
      matched: safeMatched,
      id,
      savedAt: Date.now(),
      savedDate: new Date().toLocaleDateString("zh-TW"),
    };

    // ── 先回寫申請紀錄（不依賴 storage 是否成功）────────────────────────
    const updatedProjects = [];
    setProjects(prev => {
      const updated = [...prev];
      batch.matched.forEach(r => {
        if (!r.match?.id) return;
        const idx = updated.findIndex(p => p.id === r.match.id);
        if (idx < 0) return;
        const p = { ...updated[idx] };
        const feeAmt = r.feeConfirmed ? (r.fee || 0) : 0;
        const otherAmt = r.otherDiff || 0;
        const absDiff = r.diff !== null ? Math.abs(r.diff) : null;
        const diffExplained = absDiff !== null && absDiff <= feeAmt + otherAmt;
        if (r.diff === 0 || diffExplained || r.diff !== null) {
          p.status = "已付款";
          p.paid = r.amount;
        }
        p.paidAmount = r.amount;
        if (r.date) p.paidDate = parseRocDate(r.date);
        if (r.feeConfirmed && r.fee) p.bankFee = r.fee;
        if (r.otherDiff) p.commission = r.otherDiff;
        if (r.diffNote) p.commissionNo = r.diffNote;
        if (r.voucherNo) p.voucherNo = r.voucherNo;
        if (r.expenseNo) p.expenseNo = r.expenseNo;
        updated[idx] = p;
        updatedProjects.push(p);
      });
      return updated;
    });
    for (const p of updatedProjects) {
      try { const {_effectiveStatus,...clean}=p; await sb.from("projects").eq("id",clean.id).update(projectToDb(clean)); }
      catch(err) { console.warn("回寫 project 失敗:", err.message); }
    }

    // ── 更新本地歷史 state（無論 storage 是否成功都先更新 UI）────────────
    setHistory(prev => {
      const exists = prev.find(h => h.id === id);
      return exists ? prev.map(h => h.id === id ? record : h) : [record, ...prev];
    });
    setWorkingBatch(prev => prev.filter(b => b !== batch));
    setEditingHistoryId(null);
    if (workingBatch.filter(b => b !== batch).length === 0) setActiveTab("history");

    // ── 寫入 Supabase ────────────────────────────────────────────────────
    try {
      await fetch(`${SUPABASE_URL}/rest/v1/reconcile_batches`, {
        method: "POST",
        headers: {"apikey":SUPABASE_ANON_KEY,"Authorization":`Bearer ${SUPABASE_ANON_KEY}`,"Content-Type":"application/json","Prefer":"resolution=merge-duplicates"},
        body: JSON.stringify({id:record.id,file_name:record.fileName,bank_type:record.bankType,saved_at:record.savedAt,saved_date:record.savedDate,parse_error:record.parseError||null}),
      });
      await sb.from("reconcile_rows").eq("batch_id",record.id).delete();
      if (record.matched.length > 0) {
        await sb.from("reconcile_rows").insert(record.matched.map((r,i) => ({
          batch_id:record.id, row_index:i, date_val:r.date||null,
          amount:r.amount??null, summary:r.summary||"", diff:r.diff??null,
          diff_note:r.diffNote||"", fee:r.fee??null, fee_confirmed:r.feeConfirmed||false,
          other_diff:r.otherDiff??null, voucher_no:r.voucherNo||"", expense_no:r.expenseNo||"",
          match_id:r.match?.id??null, match_title:r.match?.title||null,
          match_tax_amount:r.match?.taxAmount??null, match_client_title:r.match?.clientTitle||null,
          match_order_no:r.match?.orderNo||null,
        })));
      }
    } catch(err) { console.warn("Supabase reconcile write failed:", err.message); }
  };

  // ── 歷史記錄進入編輯模式 ─────────────────────────────────────────────
  const editHistory = (record) => {
    setEditingHistoryId(record.id);
    setActiveHistory(null);
  };

  // ── 更新歷史中某列 ────────────────────────────────────────────────────
  const updateHistoryRow = (recordId, rowIdx, patch) => {
    setHistory(prev => prev.map(h => h.id !== recordId ? h : {
      ...h, matched: h.matched.map((r, ri) => ri !== rowIdx ? r : { ...r, ...patch })
    }));
  };
  const deleteHistoryRow = (recordId, rowIdx) => {
    setHistory(prev => prev.map(h => h.id !== recordId ? h : {
      ...h, matched: h.matched.filter((_, ri) => ri !== rowIdx)
    }));
  };

  // ── 刪除歷史 ─────────────────────────────────────────────────────────
  const deleteHistory = async (record) => {
    try {
      await sb.from("reconcile_batches").eq("id", record.id).delete();
      setHistory(prev => prev.filter(h => h.id !== record.id));
      if (activeHistory?.id === record.id) setActiveHistory(null);
      if (editingHistoryId === record.id) setEditingHistoryId(null);
    } catch(err) { console.warn("刪除對帳歷史失敗:", err.message); }
  };

  // ── 更新 workingBatch 列 ──────────────────────────────────────────────
  const updateRow = (batchIdx, rowIdx, patch) => {
    setWorkingBatch(prev => prev.map((b, bi) => bi !== batchIdx ? b : {
      ...b, matched: b.matched.map((r, ri) => ri !== rowIdx ? r : { ...r, ...patch })
    }));
  };
  const deleteRow = (batchIdx, rowIdx) => {
    setWorkingBatch(prev => prev.map((b, bi) => bi !== batchIdx ? b : {
      ...b, matched: b.matched.filter((_, ri) => ri !== rowIdx)
    }));
  };

  // ── 勾選列 ───────────────────────────────────────────────────────────
  const toggleRow = (key) => setSelectedRows(prev => ({ ...prev, [key]: !prev[key] }));
  const toggleAll = (rows, batchKey) => {
    const keys = rows.map((_, i) => `${batchKey}_${i}`);
    const allSelected = keys.every(k => selectedRows[k]);
    setSelectedRows(prev => {
      const next = { ...prev };
      keys.forEach(k => { next[k] = !allSelected; });
      return next;
    });
  };

  // ── 匯出 Excel ───────────────────────────────────────────────────────
  const exportExcel = async (batch, batchKey) => {
    try {
      await loadSheetJS();
      const rows = batch.matched
        .filter((_, i) => selectedRows[`${batchKey}_${i}`])
        .map(r => ({
          "帳務日期": r.date,
          "摘要": r.desc,
          "備註": r.note !== "null" && r.note !== "--" ? r.note : "",
          "銀行帳號": r.bankAccount !== "null" ? r.bankAccount : "",
          "存入金額": r.amount,
          "比對方式": r.matchReason || "",
          "比對專案": r.match?.title || "",
          "應收金額": r.match?.taxAmount || "",
          "差異": r.diff ?? "",
          "其他差異": r.otherDiff ?? "",
          "差異說明": r.diffNote || "",
          "支出申請單編號": r.expenseNo || "",
          "傳票編號": r.voucherNo || "",
          "狀態": r.diff === 0 ? "已付款" : r.match ? "金額差異" : "查無發票",
        }));
      if (rows.length === 0) { alert("請先勾選要匯出的列"); return; }
      const ws = window.XLSX.utils.json_to_sheet(rows);
      const wb = window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(wb, ws, "對帳記錄");
      window.XLSX.writeFile(wb, `${batch.fileName}_對帳.xlsx`);
    } catch (err) {
      alert("匯出失敗：" + err.message);
    }
  };

  const matchColor = (r) => {
    if (!r.match) return { bg: "#F3E8FF", tag: <span style={{ background: "#EDE0FF", color: "#6B21A8", padding: "3px 10px", borderRadius: 20, fontSize: 12, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 5 }}><span style={{ width: 6, height: 6, borderRadius: "50%", background: "#9333EA", display: "inline-block" }} />查無發票</span> };
    if (r.diff === 0) return { bg: "#F0FFF4", tag: <Badge status="已付款" /> };
    if (r.matchReason?.includes("人工確認")) return { bg: "#FFFBEA", tag: <span style={{ color: "#B7791F", fontWeight: 600, fontSize: 12 }}>🔍 請人工確認</span> };
    return { bg: "#FFF8E1", tag: <span style={{ color: "#E65100", fontWeight: 600, fontSize: 12 }}>⚠ 金額差異</span> };
  };

  // ── 渲染比對結果表格（共用）─────────────────────────────────────────
  const renderTable = (batch, batchKey, isHistory = false) => {
    const { matched, bankType, fileName } = batch;
    const isEditingHistory = isHistory && editingHistoryId === batch.id;
    const readonly = isHistory && !isEditingHistory;
    // 手機版不使用 sticky（直接傳回普通樣式）
    const mTh = isMobile
      ? { padding: "9px 10px", textAlign: "left", color: "#555", fontWeight: 600, borderBottom: "2px solid #F0F0F0", whiteSpace: "nowrap", background: "#F8F9FA" }
      : { ...stickyTh, top: 0, zIndex: 5 };
    const mTd = isMobile ? { ...td } : { ...stickyTd };
    const unpaidProjects = projects.filter(p => {
      const es = p._effectiveStatus;
      // 已開立或到期未付款
      if (["已開立","到期未付款"].includes(es)) return true;
      // 或者目前對帳列中有比對到這個 project（保留在選單中以便顯示已帶入的值）
      if (matched.some(r => r.match?.id === p.id)) return true;
      return false;
    });

    const onUpdateRow = isHistory
      ? (ri, patch) => updateHistoryRow(batch.id, ri, patch)
      : (ri, patch) => updateRow(batchKey, ri, patch);
    const onDeleteRow = isHistory
      ? (ri) => deleteHistoryRow(batch.id, ri)
      : (ri) => deleteRow(batchKey, ri);

    const selectedCount = matched.filter((_, i) => selectedRows[`${batchKey}_${i}`]).length;
    const hasDiff = (r) => r.match && r.diff !== null && r.diff !== 0;

    return (
      <Card key={fileName + batchKey} style={{ marginBottom: 16 }}>
        {/* ── 卡片標題列 ── */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
          <div>
            <span style={{ fontWeight: 700, fontSize: 15, color: "#1a1a2e" }}>
              {bankType === "yushan" ? "🏦 玉山銀行" : bankType === "cathay" ? "🏦 國泰世華" : "📄"} {fileName}
            </span>
            {batch.parseError && <div style={{ color: "#c33", fontSize: 12, marginTop: 2 }}>❌ {batch.parseError}</div>}
            {batch.savedDate && <div style={{ fontSize: 12, color: "#888", marginTop: 2 }}>
              儲存於 {batch.savedDate}{isEditingHistory && <span style={{ color: "#E65100", marginLeft: 8 }}>（編輯中）</span>}
            </div>}
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            {/* 摘要小標籤 */}
            {[
              { label: "✓ 吻合", count: matched.filter(r => r.match && r.diff === 0).length, color: "#4CAF50" },
              { label: "⚠ 差異", count: matched.filter(r => r.match && r.diff !== 0).length, color: "#E65100" },
              { label: "✗ 查無發票", count: matched.filter(r => !r.match).length, color: "#9333EA" },
            ].map(s => s.count > 0 && (
              <span key={s.label} style={{ fontSize: 12, background: "#F8F9FA", borderRadius: 20, padding: "3px 10px", color: s.color, fontWeight: 700, border: `1px solid ${s.color}33` }}>
                {s.label} {s.count}
              </span>
            ))}

            {/* 匯出 Excel 按鈕（勾選後顯示數量） */}
            <button onClick={() => exportExcel(batch, batchKey)}
              style={{ border: "1px solid #28A745", background: "#EAFBF0", color: "#155724", padding: "6px 14px", borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
              📥 匯出{selectedCount > 0 ? `（${selectedCount}筆）` : ""}
            </button>

            {/* 歷史記錄：編輯 / 儲存 / 取消 / 刪除 */}
            {isHistory && !isEditingHistory && (
              <>
                <button onClick={() => editHistory(batch)}
                  style={{ border: "1px solid #007BFF", background: "#EBF5FF", color: "#0056b3", padding: "6px 14px", borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
                  ✏️ 編輯
                </button>
                <button onClick={() => deleteHistory(batch)}
                  style={{ border: "1px solid #fcc", background: "#fff5f5", color: "#c33", padding: "6px 14px", borderRadius: 8, cursor: "pointer", fontSize: 13 }}>
                  🗑
                </button>
              </>
            )}
            {isEditingHistory && (
              <>
                <button onClick={() => saveBatch(batch)}
                  style={{ border: "none", background: "#1a1a2e", color: "#fff", padding: "6px 16px", borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
                  💾 儲存變更
                </button>
                <button onClick={() => setEditingHistoryId(null)}
                  style={{ border: "1px solid #ddd", background: "#fff", padding: "6px 12px", borderRadius: 8, cursor: "pointer", fontSize: 13 }}>
                  取消
                </button>
              </>
            )}
            {!isHistory && (
              <button onClick={() => saveBatch(batch)}
                style={{ border: "none", background: "#1a1a2e", color: "#fff", padding: "6px 16px", borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
                💾 儲存至資料庫
              </button>
            )}
          </div>
        </div>

        {matched.length > 0 && !batch.parseError && (
          <div className="rwd-scroll-container" style={{ overflowX: "auto", overflowY: "auto", maxHeight: 520, WebkitOverflowScrolling: "touch" }}>
            <table style={{ borderCollapse: "collapse", fontSize: 12, tableLayout: isMobile ? "auto" : "fixed", width: isMobile ? "auto" : "100%", minWidth: isMobile ? 700 : 1450 }}>
              <thead>
                <tr style={{ background: "#F8F9FA" }}>
                  {/* 全選 checkbox */}
                  <th style={{ padding: "9px 10px", borderBottom: "2px solid #F0F0F0", width: 36, minWidth: 36, position: isMobile ? "static" : "sticky", left: 0, top: 0, zIndex: isMobile ? 0 : 5, background: "#F8F9FA" }}>
                    <input type="checkbox"
                      checked={matched.length > 0 && matched.every((_, i) => selectedRows[`${batchKey}_${i}`])}
                      onChange={() => toggleAll(matched, batchKey)}
                    />
                  </th>
                  <th style={{ ...mTh, left: 36,  width: 90,  minWidth: 90  }}>帳務日期</th>
                  <th style={{ ...mTh, left: 126, width: 150, minWidth: 150 }}>摘要/備註</th>
                  <th style={{ ...mTh, left: 276, width: 90,  minWidth: 90,  borderRight: isMobile ? "none" : "2px solid #D0D0D0" }}>存入金額</th>
                  {[
                    ["比對方式", 120], ["比對專案", 200], ["應收金額", 90], ["差異", 80],
                    ["其他差異", 90], ["差異說明", 100], ["支出申請單編號", 130], ["傳票編號", 110],
                    ["狀態", 110],
                    ...(!readonly ? [["", 50]] : [])
                  ].map(([h, w]) => (
                    <th key={h} style={{ padding: "9px 10px", textAlign: "left", color: "#555", fontWeight: 600, borderBottom: "2px solid #F0F0F0", whiteSpace: "nowrap", position: "static", background: "#F8F9FA", width: w, minWidth: w }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {matched.map((r, ri) => {
                  const { bg, tag } = matchColor(r);
                  const showDiffFields = hasDiff(r);
                  const absDiff = r.diff !== null ? Math.abs(r.diff) : 0;
                  const rowKey = `${batchKey}_${ri}`;
                  const cellBg = selectedRows[rowKey] ? "#EEF6FF" : bg;
                  return (
                    <tr key={ri} style={{ borderBottom: "1px solid #F8F8F8", background: cellBg }}>
                      {/* 勾選 — sticky 左 */}
                      <td style={{ ...td, textAlign: "center", position: "sticky", left: 0, zIndex: 2, background: cellBg, width: 36, minWidth: 36 }}>
                        <input type="checkbox" checked={!!selectedRows[rowKey]} onChange={() => toggleRow(rowKey)} />
                      </td>
                      {/* 凍結欄 */}
                      <td style={{ ...mTd, left: 36,  width: 90,  minWidth: 90,  background: cellBg, whiteSpace: "nowrap" }}>{r.date}</td>
                      <td style={{ ...mTd, left: 126, width: 150, minWidth: 150, background: cellBg }}>
                        <div style={{ fontWeight: 500, fontSize: 12 }}>{r.desc}</div>
                        {r.note && r.note !== "null" && r.note !== "--" && <div style={{ fontSize: 11, color: "#888", marginTop: 2 }}>{r.note}</div>}
                      </td>
                      <td style={{ ...mTd, left: 276, width: 90,  minWidth: 90,  background: cellBg, borderRight: "2px solid #D0D0D0", fontWeight: 700, color: "#2E7D32", whiteSpace: "nowrap" }}>{r.amount.toLocaleString()}</td>
                      <td style={{ ...td, fontSize: 11, color: "#666" }}>{r.matchReason || "-"}</td>
                      {/* 比對專案 - 無論是否已比對，都顯示下拉供修改 */}
                      <td style={{ ...td, minWidth: 180 }}>
                        {readonly ? (
                          <div style={{ fontSize: 12 }}>
                            {r.match ? (
                              <div>
                                <div style={{ fontWeight: 600 }}>{r.match.orderNo || r.match.title}</div>
                                <div style={{ color: "#888", fontSize: 11 }}>{r.match.title}</div>
                              </div>
                            ) : <span style={{ color: "#aaa" }}>查無發票</span>}
                          </div>
                        ) : (
                          <select
                            style={{ ...inp, padding: "4px 6px", fontSize: 11,
                              borderColor: r.match ? "#C8DCF7" : "#FFC0C0",
                              background: r.match ? "#F8FBFF" : "#FFF5F5" }}
                            value={r.match?.id || ""}
                            onChange={e => {
                              if (!e.target.value) {
                                onUpdateRow(ri, { match: null, diff: null, matchReason: "" });
                                return;
                              }
                              const proj = projects.find(p => p.id === +e.target.value);
                              if (proj) onUpdateRow(ri, {
                                match: { id: proj.id, title: proj.title, taxAmount: proj.taxAmount, clientTitle: proj.clientTitle || "", orderNo: proj.orderNo || "" },
                                diff: r.amount - proj.taxAmount,
                                matchReason: r.matchReason || "人工指定",
                              });
                            }}>
                            <option value="">— 手動選擇 —</option>
                            {unpaidProjects.map(p => (
                              <option key={p.id} value={p.id}>
                                {p.orderNo ? `${p.orderNo} ` : ""}{p.title}（{p.taxAmount.toLocaleString()}）
                              </option>
                            ))}
                          </select>
                        )}
                      </td>
                      <td style={td}>{r.match ? r.match.taxAmount.toLocaleString() : "-"}</td>
                      {/* 差異 */}
                      <td style={{ ...td, fontWeight: 700, whiteSpace: "nowrap", color: r.diff === 0 ? "#4CAF50" : r.diff !== null ? "#E65100" : "#aaa" }}>
                        {r.diff !== null ? (r.diff === 0 ? "✓ 0" : `${r.diff > 0 ? "+" : ""}${r.diff.toLocaleString()}`) : "-"}
                      </td>
                      {/* 其他差異 */}
                      <td style={{ ...td, minWidth: 90 }}>
                        {showDiffFields ? (
                          readonly ? (
                            <span style={{ fontSize: 12 }}>{r.otherDiff !== undefined && r.otherDiff !== "" ? (+r.otherDiff).toLocaleString() : "-"}</span>
                          ) : (
                            <input type="text" inputMode="numeric"
                              style={{ ...inp, padding: "4px 6px", fontSize: 11, width: 70, textAlign: "right" }}
                              value={r.otherDiff !== undefined ? r.otherDiff : ""}
                              onChange={e => { const v = parseInt(e.target.value.replace(/\D/g,""),10); onUpdateRow(ri, { otherDiff: isNaN(v) ? "" : v }); }}
                              placeholder="金額"
                            />
                          )
                        ) : <span style={{ color: "#ddd" }}>—</span>}
                      </td>
                      {/* 差異說明（下拉） */}
                      <td style={{ ...td, minWidth: 110 }}>
                        {showDiffFields ? (
                          readonly ? (
                            <span style={{ fontSize: 12, color: "#555" }}>{r.diffNote || "-"}</span>
                          ) : (
                            <select
                              style={{ ...inp, padding: "4px 6px", fontSize: 11, minWidth: 90 }}
                              value={r.diffNote || ""}
                              onChange={e => onUpdateRow(ri, { diffNote: e.target.value })}>
                              <option value="">— 選擇 —</option>
                              <option value="手續費">手續費</option>
                              <option value="佣金直扣">佣金直扣</option>
                            </select>
                          )
                        ) : <span style={{ color: "#ddd" }}>—</span>}
                      </td>
                      {/* 支出申請單編號 */}
                      <td style={{ ...td, minWidth: 130 }}>
                        {readonly ? (
                          <span style={{ fontSize: 12, color: "#555" }}>{r.expenseNo || "-"}</span>
                        ) : (
                          <input type="text"
                            style={{ ...inp, padding: "4px 6px", fontSize: 11, minWidth: 110 }}
                            value={r.expenseNo || ""}
                            onChange={e => onUpdateRow(ri, { expenseNo: e.target.value })}
                            placeholder="申請單編號"
                          />
                        )}
                      </td>
                      {/* 傳票編號 */}
                      <td style={{ ...td, minWidth: 110 }}>
                        {readonly ? (
                          <span style={{ fontSize: 12, color: "#555" }}>{r.voucherNo || "-"}</span>
                        ) : (
                          <input type="text"
                            style={{ ...inp, padding: "4px 6px", fontSize: 11, minWidth: 90 }}
                            value={r.voucherNo || ""}
                            onChange={e => onUpdateRow(ri, { voucherNo: e.target.value })}
                            placeholder="傳票編號"
                          />
                        )}
                      </td>
                      <td style={td}>{tag}</td>
                      {!readonly && (
                        <td style={{ ...td, textAlign: "center" }}>
                          <button onClick={() => onDeleteRow(ri)}
                            style={{ border: "1px solid #fcc", background: "#fff5f5", color: "#c33", padding: "3px 8px", borderRadius: 6, cursor: "pointer", fontSize: 12 }}>🗑</button>
                        </td>
                      )}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    );
  };

  return (
    <div>
      <h2 style={{ fontSize: 22, fontWeight: 700, marginBottom: 24, color: "#1a1a2e" }}>🏦 銀行對帳核對</h2>

      {/* ── 上傳區 ── */}
      <Card style={{ marginBottom: 16 }}>
        <h3 style={sectionTitle}>上傳對帳單</h3>
        <div style={{ display: "flex", gap: 12, alignItems: "flex-start", flexWrap: "wrap", marginBottom: 14 }}>
          <div style={{ background: "#F0F7FF", border: "1px solid #BFD9F7", borderRadius: 10, padding: "10px 16px", fontSize: 13, color: "#1565C0", lineHeight: 1.7 }}>
            <b>玉山銀行</b>：支援 .xls　　<b>國泰世華</b>：支援 .xlsx<br/>
            可一次選擇多份檔案同時上傳
          </div>
        </div>
        <label style={{ display: "inline-block", border: "none", background: "#1a1a2e", color: "#fff", padding: "11px 22px", borderRadius: 8, cursor: "pointer", fontSize: 14, fontWeight: 600 }}>
          📂 選擇對帳單（可多選）
          <input type="file" accept=".xls,.xlsx" multiple style={{ display: "none" }} onChange={handleFiles} />
        </label>
        {loading && (
          <div style={{ marginTop: 12, fontSize: 13, color: "#888" }}>
            ⏳ 解析中：{loadingFiles.join("、")}
          </div>
        )}
        {error && <div style={{ marginTop: 10, color: "#c33", fontSize: 13 }}>{error}</div>}
      </Card>

      {/* ── Tab 切換 ── */}
      {(workingBatch.length > 0 || history.length > 0) && (
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, borderBottom: "2px solid #F0F0F0" }}>
          <div style={{ display: "flex", gap: 0 }}>
            {[
              { id: "working", label: `📋 待確認（${workingBatch.length}）`, show: workingBatch.length > 0 },
              { id: "history", label: `🗄 歷史資料庫（${history.length}）`, show: true },
            ].filter(t => t.show).map(t => (
              <button key={t.id} onClick={() => setActiveTab(t.id)}
                style={{ padding: "10px 22px", border: "none", background: "none", fontSize: 14, fontWeight: activeTab === t.id ? 700 : 400, color: activeTab === t.id ? "#1a1a2e" : "#888", borderBottom: activeTab === t.id ? "2px solid #1a1a2e" : "2px solid transparent", cursor: "pointer", marginBottom: -2 }}>
                {t.label}
              </button>
            ))}
          </div>
          {/* 歷史資料庫彙總匯出 */}
          {activeTab === "history" && history.length > 0 && (
            <button onClick={async () => {
              try {
                if (!window.XLSX) {
                  await new Promise((res, rej) => {
                    const s = document.createElement("script");
                    s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
                    s.onload = res; s.onerror = rej;
                    document.head.appendChild(s);
                  });
                }
                const wb = window.XLSX.utils.book_new();
                history.forEach(batch => {
                  const rows = batch.matched.map(r => ({
                    "帳務日期": r.date,
                    "摘要": r.desc,
                    "備註": r.note !== "null" && r.note !== "--" ? r.note : "",
                    "銀行帳號": r.bankAccount !== "null" ? r.bankAccount : "",
                    "存入金額": r.amount,
                    "比對方式": r.matchReason || "",
                    "比對專案": r.match?.title || "",
                    "應收金額": r.match?.taxAmount || "",
                    "差異": r.diff ?? "",
                    "其他差異": r.otherDiff ?? "",
                    "差異說明": r.diffNote || "",
                    "支出申請單編號": r.expenseNo || "",
                    "傳票編號": r.voucherNo || "",
                    "狀態": r.diff === 0 ? "已付款" : r.match ? "金額差異" : "查無發票",
                  }));
                  // 工作表名稱限31字元，去除非法字元
                  const sheetName = batch.fileName.replace(/[\\\/\?\*\[\]]/g, "").slice(0, 31);
                  const ws = window.XLSX.utils.json_to_sheet(rows);
                  window.XLSX.utils.book_append_sheet(wb, ws, sheetName);
                });
                const date = new Date().toLocaleDateString("zh-TW").replace(/\//g, "-");
                window.XLSX.writeFile(wb, `銀行對帳彙總_${date}.xlsx`);
              } catch(err) { alert("匯出失敗：" + err.message); }
            }}
              style={{ border: "1px solid #28A745", background: "#EAFBF0", color: "#155724", padding: "7px 16px", borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600, marginBottom: 2 }}>
              📥 匯出全部歷史記錄
            </button>
          )}
        </div>
      )}

      {/* ── 待確認（工作區）── */}
      {activeTab === "working" && workingBatch.length > 0 && (
        <div>
          <div style={{ fontSize: 13, color: "#888", marginBottom: 12 }}>
            確認比對結果後，點「💾 儲存至資料庫」保存至歷史記錄。
          </div>
          {workingBatch.map((batch, bi) => renderTable(batch, bi, false))}
        </div>
      )}

      {/* ── 歷史資料庫 ── */}
      {activeTab === "history" && (
        <div>
          {history.length === 0 ? (
            <Card><div style={{ textAlign: "center", color: "#aaa", padding: 32 }}>尚無已儲存的對帳記錄</div></Card>
          ) : (
            <div>
              {/* 歷史列表索引 */}
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
                {history.map(h => (
                  <button key={h.id} onClick={() => setActiveHistory(activeHistory?.id === h.id ? null : h)}
                    style={{ padding: "7px 14px", border: `1.5px solid ${activeHistory?.id === h.id ? "#1a1a2e" : "#E0E0E0"}`, borderRadius: 8, background: activeHistory?.id === h.id ? "#1a1a2e" : "#fff", color: activeHistory?.id === h.id ? "#fff" : "#444", cursor: "pointer", fontSize: 13 }}>
                    {h.bankType === "yushan" ? "🏦" : "🏦"} {h.fileName}
                    <span style={{ marginLeft: 6, fontSize: 11, opacity: 0.7 }}>{h.savedDate}</span>
                  </button>
                ))}
              </div>
              {/* 展開檢視 */}
              {/* 展開檢視 或 編輯中 */}
              {activeHistory && !editingHistoryId && renderTable(activeHistory, `hist_${activeHistory.id}`, true)}
              {editingHistoryId && (() => {
                const rec = history.find(h => h.id === editingHistoryId);
                return rec ? renderTable(rec, `hist_${rec.id}`, true) : null;
              })()}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── VENDOR MANAGEMENT ────────────────────────────────────────────────────────
function VendorManagement({ vendors, setVendors, canEdit }) {
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState({ name: "", taxId: "", bank: "", branch: "", accountName: "", account: "" });
  const [showForm, setShowForm] = useState(false);

  const save = () => {
    if (editing !== null) {
      setVendors(vendors.map((v, i) => i === editing ? { ...v, ...form } : v));
    } else {
      setVendors([...vendors, { id: Date.now(), ...form }]);
    }
    setEditing(null); setShowForm(false); setForm({ name: "", taxId: "", bank: "", branch: "", accountName: "", account: "" });
  };

  const fields = [["name","公司名稱"],["taxId","統一編號"],["bank","匯款銀行代號"],["branch","分行代號"],["accountName","戶名"],["account","帳號（勿加標點符號如：- / . ）"]];

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
        <h2 style={{ fontSize: 22, fontWeight: 700, color: "#1a1a2e" }}>🏢 廠商資料管理</h2>
        {canEdit && <button onClick={() => { setShowForm(true); setEditing(null); setForm({ name: "", taxId: "", bank: "", branch: "", accountName: "", account: "" }); }}
          style={{ border: "none", background: "#1a1a2e", color: "#fff", padding: "10px 20px", borderRadius: 10, cursor: "pointer", fontSize: 14 }}>＋ 新增廠商</button>}
      </div>

      {showForm && (
        <Card style={{ marginBottom: 20, borderColor: "#1a1a2e33" }}>
          <h3 style={sectionTitle}>{editing !== null ? "編輯廠商" : "新增廠商"}</h3>
          <div className="rwd-grid2" style={grid2}>
            {fields.map(([k, l]) => (
              <div key={k} style={field}>
                <label style={lbl}>{l}</label>
                <input style={inp} value={form[k]} onChange={e => setForm({ ...form, [k]: e.target.value })} placeholder={l} />
              </div>
            ))}
          </div>
          <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
            <button onClick={save} style={{ border: "none", background: "#1a1a2e", color: "#fff", padding: "10px 20px", borderRadius: 8, cursor: "pointer" }}>儲存</button>
            <button onClick={() => { setShowForm(false); setEditing(null); }} style={{ border: "1px solid #ddd", background: "#fff", padding: "10px 20px", borderRadius: 8, cursor: "pointer" }}>取消</button>
          </div>
        </Card>
      )}

      <div style={{ display: "grid", gap: 12 }}>
        {vendors.map((v, i) => (
          <Card key={v.id} style={{ padding: "16px 20px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 15, color: "#1a1a2e", marginBottom: 4 }}>{v.name}</div>
                <div style={{ fontSize: 12, color: "#888", lineHeight: 2 }}>
                  統編：{v.taxId}　｜　{v.bank} {v.branch}　｜　戶名：{v.accountName}　｜　帳號：{v.account}
                </div>
              </div>
              {canEdit && <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => { setEditing(i); setForm({ name: v.name, taxId: v.taxId, bank: v.bank, branch: v.branch, accountName: v.accountName, account: v.account }); setShowForm(true); }}
                  style={{ border: "1px solid #ddd", background: "#F8F9FA", padding: "6px 12px", borderRadius: 8, cursor: "pointer", fontSize: 12 }}>✏️ 編輯</button>
                <button onClick={() => setVendors(vendors.filter((_, j) => j !== i))}
                  style={{ border: "1px solid #fcc", background: "#fff5f5", color: "#c33", padding: "6px 12px", borderRadius: 8, cursor: "pointer", fontSize: 12 }}>刪除</button>
              </div>}
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
}

// ─── PROJECT LIST ─────────────────────────────────────────────────────────────
function ProjectList({ projects, setProjects, vendors, currentUser, users, onExport, onCopyProject }) {
  const isMobile = useIsMobile();
  const sTh = isMobile ? { padding: "12px 14px", textAlign: "left", color: "#555", fontWeight: 600, borderBottom: "2px solid #F0F0F0", whiteSpace: "nowrap", background: "#F8F9FA" } : stickyTh;
  const sTd = isMobile ? { padding: "12px 14px", color: "#444", verticalAlign: "middle" } : stickyTd;
  const visible = projects.filter(p => {
    if (currentUser.role === "admin" || currentUser.role === "finance") return true;
    if (currentUser.role === "manager") return p.company === currentUser.company;
    return p.applicant === currentUser.name;
  });

  const [filter, setFilter] = useState("");
  const [selectedMonth, setSelectedMonth] = useState("all");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [detailModal, setDetailModal] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [deleteConfirm, setDeleteConfirm] = useState(null);
  const [issueModal, setIssueModal] = useState(null);
  const [invoiceNo, setInvoiceNo] = useState("");

  const canEdit = currentUser.role === "finance" || currentUser.role === "admin";

  // 月份列表（依發票開立日期，由新到舊）
  const months = useMemo(() => {
    const set = new Set(visible.map(p => (p.invoiceDate || "").slice(0, 7)).filter(Boolean));
    return ["all", ...Array.from(set).sort((a, b) => b.localeCompare(a))];
  }, [visible]);

  const fmtMonth = (m) => {
    if (m === "all") return "全部";
    const [y, mo] = m.split("-");
    return `${y}年${parseInt(mo)}月`;
  };

  const filtered = visible
    .filter(p => selectedMonth === "all" || (p.invoiceDate || "").startsWith(selectedMonth))
    .filter(p => !dateFrom || (p.invoiceDate || "") >= dateFrom)
    .filter(p => !dateTo   || (p.invoiceDate || "") <= dateTo)
    .filter(p => !filter || p._effectiveStatus === filter);

  const startEdit = (p) => {
    setEditingId(p.id);
    setEditForm({ ...p });
  };

  const saveEdit = () => {
    const { _effectiveStatus, ...cleanForm } = editForm; // 排除計算屬性
    setProjects(prev => prev.map(p => p.id === editingId ? { ...cleanForm } : p));
    setEditingId(null);
  };

  const cancelEdit = () => setEditingId(null);

  const confirmMarkIssued = () => {
    if (!invoiceNo.trim()) return;
    setProjects(projects.map(p => p.id === issueModal.id ? { ...p, status: "已開立", invoiceNo: invoiceNo.trim() } : p));
    setIssueModal(null);
    setInvoiceNo("");
  };

  const deleteProject = (id) => {
    setProjects(projects.filter(p => p.id !== id));
    setDeleteConfirm(null);
  };

  const ef = editForm;
  const setEf = (k, v) => setEditForm(prev => ({ ...prev, [k]: v }));

  return (
    <div>
      {/* ── 明細 Modal ── */}
      {detailModal && (() => {
        const dm = detailModal;
        const dv = vendors.find(x => x.id === dm.vendorId);
        const taxAmt = dm.taxAmount || 0;
        const untaxAmt = dm.amount || Math.round(taxAmt / 1.05);
        const taxOnly = taxAmt - untaxAmt;
        const isEditingModal = detailModal._editing || false;
        const ef2 = detailModal._editForm || {};
        const setEf2 = (k, v) => setDetailModal(prev => ({ ...prev, _editForm: { ...prev._editForm, [k]: v } }));
        const editItems = ef2.items && ef2.items.length > 0 ? ef2.items : [{label:"",qty:1,unitPrice:0,untaxAmt:0,taxAmt:0}];
        const updateItem = (i, field, val) => {
          const items = editItems.map((it,idx) => idx!==i ? it : { ...it, [field]: val });
          if (field==="qty" || field==="unitPrice") {
            items[i].untaxAmt = (items[i].unitPrice||0) * (items[i].qty||1);
            items[i].taxAmt = Math.round(items[i].untaxAmt * 1.05);
          }
          if (field==="untaxAmt") items[i].taxAmt = Math.round(val * 1.05);
          const totalTax = items.reduce((s,it)=>s+(it.taxAmt||0),0);
          const totalUntax = items.reduce((s,it)=>s+(it.untaxAmt||0),0);
          setEf2("items", items); setEf2("taxAmount", totalTax); setEf2("amount", totalUntax);
        };

        const Row = ({label, value, mono, highlight}) => value!=null && value!=="" ? (
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", padding:"11px 0", borderBottom:"1px solid #F0F0F0", gap:12 }}>
            <span style={{ fontSize:13, color:"#888", whiteSpace:"nowrap", flexShrink:0 }}>{label}</span>
            <span style={{ fontSize:13, fontWeight:highlight?700:600, color:highlight?"#F2A0A0":"#1a1a2e", textAlign:"right", fontFamily:mono?"monospace":undefined }}>{value}</span>
          </div>
        ) : null;
        const Section = ({title, color="#F2A0A0", children}) => (
          <div style={{ marginBottom:4 }}>
            <div style={{ fontSize:12, fontWeight:700, color, marginBottom:4, paddingBottom:6, borderBottom:`2px solid ${color}` }}>{title}</div>
            {children}
          </div>
        );
        const ERow = ({label, field, type="text", options}) => (
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"8px 0", borderBottom:"1px solid #F0F0F0", gap:12 }}>
            <span style={{ fontSize:13, color:"#888", whiteSpace:"nowrap", flexShrink:0, width:90 }}>{label}</span>
            {options ? (
              <select style={{ ...inp, padding:"4px 8px", fontSize:13, flex:1 }} value={ef2[field]||""} onChange={e=>setEf2(field,e.target.value)}>
                {options.map(o=><option key={o}>{o}</option>)}
              </select>
            ) : (
              <input style={{ ...inp, padding:"4px 8px", fontSize:13, flex:1, textAlign:type==="number"?"right":"left" }}
                type={type==="number"?"text":type} inputMode={type==="number"?"numeric":undefined}
                value={ef2[field]??""} onChange={e=>{ if(type==="number"){ const v=parseInt(e.target.value.replace(/\D/g,""),10); setEf2(field,isNaN(v)?"":v); } else { setEf2(field,e.target.value); } }}
              />
            )}
          </div>
        );
        const saveModalEdit = async () => {
          const nullIfEmpty = (v) => (v===""||v===null||v===undefined)?null:v;
          const numOrNull = (v) => { const n=Number(v); return (v===""||v===null||v===undefined||isNaN(n))?null:n; };
          const merged = { ...dm, ...ef2 };
          const clean = {
            id:merged.id, applicant:merged.applicant, company:merged.company, title:merged.title,
            clientTitle:merged.clientTitle||"", taxId:merged.taxId||"", orderNo:merged.orderNo||"",
            invoiceNo:merged.invoiceNo||"", invoiceDate:nullIfEmpty(merged.invoiceDate),
            expectedPayDate:nullIfEmpty(merged.expectedPayDate), status:merged.status,
            amount:numOrNull(merged.amount), taxAmount:numOrNull(merged.taxAmount),
            paidDate:nullIfEmpty(merged.paidDate), paidAmount:numOrNull(merged.paidAmount),
            bankFee:numOrNull(merged.bankFee), commission:numOrNull(merged.commission),
            commissionNo:merged.commissionNo||"", expenseNo:merged.expenseNo||"",
            voucherNo:merged.voucherNo||"", vendorId:merged.vendorId||null,
            paid:merged.paid||0, theme:merged.theme||"invoice",
            items:merged.items||[],
          };
          try {
            await sb.from("projects").eq("id",clean.id).update(projectToDb(clean));
            setProjects(prev=>prev.map(p=>p.id===clean.id?clean:p));
            setDetailModal({...clean,_effectiveStatus:dm._effectiveStatus});
          } catch(err) { alert("儲存失敗："+err.message); }
        };

        return (
          <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.45)", zIndex:200, display:"flex", alignItems:"flex-end", justifyContent:"center", padding:0 }}
            onClick={e=>{ if(e.target===e.currentTarget) setDetailModal(null); }}>
            <div style={{ background:"#fff", borderRadius:"20px 20px 0 0", padding:"0 0 32px", width:"100%", maxWidth:560, maxHeight:"92vh", overflowY:"auto", boxShadow:"0 -8px 40px rgba(0,0,0,0.18)" }}>
              <div style={{ display:"flex", justifyContent:"center", padding:"12px 0 4px" }}>
                <div style={{ width:40, height:4, background:"#E0E0E0", borderRadius:2 }}/>
              </div>
              <div style={{ padding:"12px 24px 16px", borderBottom:"1px solid #F0F0F0" }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
                  <div style={{ flex:1, marginRight:12 }}>
                    <div style={{ fontSize:11, color:"#aaa", marginBottom:4, fontFamily:"monospace" }}>{dm.orderNo||dm.invoiceNo||""}</div>
                    <div style={{ fontSize:18, fontWeight:700, color:"#1a1a2e", lineHeight:1.3, marginBottom:8 }}>{dm.title}</div>
                    <Badge status={dm._effectiveStatus}/>
                  </div>
                  <div style={{ display:"flex", gap:8, flexShrink:0 }}>
                    {canEdit && !isEditingModal && (
                      <button onClick={()=>setDetailModal(prev=>({...prev,_editing:true,_editForm:{
                        clientTitle:prev.clientTitle, taxId:prev.taxId, orderNo:prev.orderNo,
                        invoiceNo:prev.invoiceNo, invoiceDate:prev.invoiceDate,
                        expectedPayDate:prev.expectedPayDate, status:prev.status,
                        paidDate:prev.paidDate||"", paidAmount:prev.paidAmount||"",
                        bankFee:prev.bankFee||"", commission:prev.commission||"",
                        commissionNo:prev.commissionNo||"", expenseNo:prev.expenseNo||"",
                        voucherNo:prev.voucherNo||"", amount:prev.amount, taxAmount:prev.taxAmount,
                        items: prev.items&&prev.items.length>0 ? prev.items.map(i=>({...i})) : [{label:"",qty:1,unitPrice:0,untaxAmt:0,taxAmt:0}],
                      }}))}
                        style={{ border:"1px solid #ddd", background:"#F8F9FA", color:"#444", borderRadius:8, padding:"6px 14px", fontSize:13, cursor:"pointer", fontWeight:600 }}>
                        ✏️ 編輯
                      </button>
                    )}
                    {isEditingModal && (
                      <>
                        <button onClick={saveModalEdit} style={{ border:"none", background:"#1a1a2e", color:"#fff", borderRadius:8, padding:"6px 14px", fontSize:13, cursor:"pointer", fontWeight:700 }}>💾 儲存</button>
                        <button onClick={()=>setDetailModal(prev=>({...prev,_editing:false,_editForm:{}}))} style={{ border:"1px solid #ddd", background:"#fff", color:"#666", borderRadius:8, padding:"6px 12px", fontSize:13, cursor:"pointer" }}>取消</button>
                      </>
                    )}
                    <button onClick={()=>setDetailModal(null)} style={{ border:"none", background:"#F5F5F5", borderRadius:"50%", width:36, height:36, cursor:"pointer", fontSize:18, color:"#666", display:"flex", alignItems:"center", justifyContent:"center" }}>×</button>
                  </div>
                </div>
              </div>
              <div style={{ margin:"16px 24px", background:"#FFF0F0", borderRadius:14, padding:"16px 20px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                <span style={{ fontSize:14, fontWeight:700, color:"#F2A0A0" }}>含稅金額</span>
                <span style={{ fontSize:24, fontWeight:900, color:"#F2A0A0" }}>NT$ {taxAmt.toLocaleString()}</span>
              </div>
              <div style={{ padding:"0 24px" }}>
                <Section title="申請人資訊">
                  <Row label="申請人" value={dm.applicant}/><Row label="公司別" value={dm.company}/>
                </Section>
                <Section title="款項資訊" color="#5C85D6">
                  {isEditingModal ? <>
                    <ERow label="公司抬頭" field="clientTitle"/><ERow label="統一編號" field="taxId"/>
                    <ERow label="委刊單編號" field="orderNo"/><ERow label="未稅金額" field="amount" type="number"/>
                    <ERow label="含稅金額" field="taxAmount" type="number"/>
                  </> : <>
                    <Row label="公司抬頭" value={dm.clientTitle}/><Row label="統一編號" value={dm.taxId} mono/>
                    <Row label="委刊單編號" value={dm.orderNo} mono/>
                    <Row label="未稅金額" value={untaxAmt?`NT$ ${untaxAmt.toLocaleString()}`:null}/>
                    <Row label="5% 稅額" value={taxOnly>0?`NT$ ${taxOnly.toLocaleString()}`:null}/>
                    <Row label="含稅金額" value={taxAmt?`NT$ ${taxAmt.toLocaleString()}`:null} highlight/>
                  </>}
                </Section>
                {isEditingModal && (
                  <Section title="品項明細" color="#00897B">
                    {editItems.map((it,i)=>(
                      <div key={i} style={{ borderBottom:"1px solid #F0F0F0", padding:"10px 0" }}>
                        <div style={{ display:"flex", gap:8, alignItems:"center", marginBottom:6 }}>
                          <select style={{ ...inp, flex:2, padding:"5px 8px", fontSize:13 }} value={it.label||""} onChange={e=>updateItem(i,"label",e.target.value)}>
                            <option value="">選擇品項</option>
                            {ITEMS_CATALOG.map(c=><option key={c.label} value={c.label}>{c.label}</option>)}
                          </select>
                          <button onClick={()=>{ const items=editItems.filter((_,j)=>j!==i); const tt=items.reduce((s,x)=>s+(x.taxAmt||0),0); const tu=items.reduce((s,x)=>s+(x.untaxAmt||0),0); setEf2("items",items); setEf2("taxAmount",tt); setEf2("amount",tu); }}
                            style={{ border:"none", background:"#fee", color:"#c33", borderRadius:6, padding:"5px 10px", cursor:"pointer", fontSize:14, flexShrink:0 }}>×</button>
                        </div>
                        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:6 }}>
                          <div><div style={{ fontSize:11, color:"#aaa", marginBottom:2 }}>數量</div>
                            <input style={{ ...inp, padding:"5px 8px", fontSize:13, textAlign:"right" }} type="text" inputMode="numeric"
                              value={it.qty||""} onChange={e=>{ const v=parseInt(e.target.value.replace(/\D/g,""),10); updateItem(i,"qty",isNaN(v)?1:v); }}/></div>
                          <div><div style={{ fontSize:11, color:"#aaa", marginBottom:2 }}>單價（未稅）</div>
                            <input style={{ ...inp, padding:"5px 8px", fontSize:13, textAlign:"right" }} type="text" inputMode="numeric"
                              value={it.unitPrice||""} onChange={e=>{ const v=parseInt(e.target.value.replace(/\D/g,""),10); updateItem(i,"unitPrice",isNaN(v)?0:v); }}/></div>
                          <div><div style={{ fontSize:11, color:"#aaa", marginBottom:2 }}>含稅小計</div>
                            <div style={{ ...inp, padding:"5px 8px", fontSize:13, textAlign:"right", background:"#F0FFF4", color:"#2E7D32", fontWeight:700 }}>{it.taxAmt?it.taxAmt.toLocaleString():"-"}</div></div>
                        </div>
                      </div>
                    ))}
                    <button onClick={()=>setEf2("items",[...editItems,{label:"",qty:1,unitPrice:0,untaxAmt:0,taxAmt:0}])}
                      style={{ width:"100%", marginTop:10, padding:"8px", border:"2px dashed #00897B", background:"#F0FDF4", color:"#00897B", borderRadius:8, cursor:"pointer", fontSize:13, fontWeight:600 }}>＋ 新增品項</button>
                    <div style={{ marginTop:8, textAlign:"right", fontSize:13, color:"#2E7D32", fontWeight:700 }}>合計（含稅）：NT$ {(ef2.taxAmount||0).toLocaleString()}</div>
                  </Section>
                )}
                {!isEditingModal && (() => {
                  const hasItems = dm.items && dm.items.length > 0;
                  const fallbackItems = !hasItems && dm.title ? dm.title.split("、").map(label => ({ label: label.trim() })) : [];
                  const displayItems = hasItems ? dm.items : fallbackItems;
                  if (!displayItems.length) return null;
                  return (
                    <Section title="品項明細" color="#00897B">
                      <div style={{ overflowX:"auto" }}>
                        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                          <thead><tr style={{ background:"#F5F5F5" }}>
                            {(hasItems?["品項","數量","單價","未稅小計","含稅小計"]:["品項"]).map(h=>(
                              <th key={h} style={{ padding:"6px 8px", textAlign:h==="品項"?"left":"right", color:"#888", fontWeight:600, borderBottom:"1px solid #E0E0E0", whiteSpace:"nowrap", fontSize:12 }}>{h}</th>
                            ))}</tr></thead>
                          <tbody>{displayItems.map((it,i)=>(
                            <tr key={i} style={{ borderBottom:"1px solid #F5F5F5" }}>
                              <td style={{ padding:"7px 8px", fontWeight:500 }}>{it.label||"-"}</td>
                              {hasItems && <>
                                <td style={{ padding:"7px 8px", textAlign:"right", color:"#666" }}>{it.qty||1}</td>
                                <td style={{ padding:"7px 8px", textAlign:"right", color:"#666" }}>{it.unitPrice?it.unitPrice.toLocaleString():"-"}</td>
                                <td style={{ padding:"7px 8px", textAlign:"right" }}>{it.untaxAmt?it.untaxAmt.toLocaleString():"-"}</td>
                                <td style={{ padding:"7px 8px", textAlign:"right", fontWeight:600, color:"#2E7D32" }}>{it.taxAmt?it.taxAmt.toLocaleString():"-"}</td>
                              </>}
                            </tr>
                          ))}</tbody>
                        </table>
                        {!hasItems && <div style={{ fontSize:11, color:"#bbb", padding:"6px 8px" }}>此筆為舊資料，僅顯示品項名稱</div>}
                      </div>
                    </Section>
                  );
                })()}
                <Section title="發票資訊" color="#43A047">
                  {isEditingModal ? <>
                    <ERow label="發票號碼" field="invoiceNo"/><ERow label="傳票號碼" field="voucherNo"/>
                    <ERow label="發票日期" field="invoiceDate" type="date"/><ERow label="預計收款日" field="expectedPayDate" type="date"/>
                    <ERow label="付款狀態" field="status" options={["申請中","已開立","到期未付款","已付款"]}/>
                  </> : <>
                    <Row label="發票號碼" value={dm.invoiceNo} mono/><Row label="傳票號碼" value={dm.voucherNo} mono/>
                    <Row label="發票日期" value={dm.invoiceDate}/><Row label="預計收款日" value={dm.expectedPayDate}/>
                  </>}
                </Section>
                <Section title="付款資訊" color="#F57C00">
                  {isEditingModal ? <>
                    <ERow label="已收款金額" field="paidAmount" type="number"/><ERow label="實際收款日" field="paidDate" type="date"/>
                    <ERow label="手續費" field="bankFee" type="number"/><ERow label="佣金" field="commission" type="number"/>
                    <ERow label="說明" field="commissionNo"/><ERow label="支出申請單" field="expenseNo"/>
                  </> : <>
                    <Row label="已收款金額" value={dm.paidAmount?`NT$ ${dm.paidAmount.toLocaleString()}`:null}/>
                    <Row label="實際收款日" value={dm.paidDate}/><Row label="手續費" value={dm.bankFee?`NT$ ${dm.bankFee.toLocaleString()}`:null}/>
                    <Row label="佣金" value={dm.commission?`NT$ ${dm.commission.toLocaleString()}`:null}/>
                    <Row label="說明" value={dm.commissionNo}/><Row label="支出申請單" value={dm.expenseNo} mono/>
                  </>}
                </Section>
                {dv && (
                  <Section title="廠商付款資訊" color="#7B1FA2">
                    <Row label="廠商名稱" value={dv.name}/>
                    <Row label="銀行" value={dv.bank&&dv.branch?`${dv.bank} ${dv.branch}`:dv.bank}/>
                    <Row label="戶名" value={dv.accountName}/><Row label="帳號" value={dv.account} mono/>
                  </Section>
                )}
              </div>
              <div style={{ padding:"20px 24px 0", display:"flex", gap:10 }}>
                {isEditingModal ? (
                  <>
                    <button onClick={saveModalEdit} style={{ flex:2, padding:14, border:"none", background:"#1a1a2e", color:"#fff", borderRadius:12, fontSize:15, fontWeight:700, cursor:"pointer" }}>💾 儲存變更</button>
                    <button onClick={()=>setDetailModal(prev=>({...prev,_editing:false,_editForm:{}}))} style={{ flex:1, padding:14, border:"1px solid #ddd", background:"#fff", color:"#666", borderRadius:12, fontSize:14, cursor:"pointer" }}>取消</button>
                  </>
                ) : (
                  <button onClick={()=>setDetailModal(null)} style={{ flex:1, padding:14, border:"none", background:"#1a1a2e", color:"#fff", borderRadius:12, fontSize:15, fontWeight:700, cursor:"pointer" }}>關閉</button>
                )}
              </div>
            </div>
          </div>
        );
      })()}

            {/* 輸入發票號碼 modal */}
      {issueModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div style={{ background: "#fff", borderRadius: 16, padding: 24, width: 400, boxShadow: "0 8px 40px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 28, textAlign: "center", marginBottom: 10 }}>🧾</div>
            <div style={{ fontWeight: 700, fontSize: 16, textAlign: "center", marginBottom: 6 }}>標記已開立發票</div>
            <div style={{ fontSize: 13, color: "#666", textAlign: "center", marginBottom: 20 }}>
              {projects.find(p => p.id === issueModal.id)?.title}
            </div>
            <div style={{ marginBottom: 20 }}>
              <label style={lbl}>發票號碼</label>
              <input
                style={{ ...inp, fontSize: 15, letterSpacing: 1 }}
                value={invoiceNo}
                onChange={e => setInvoiceNo(e.target.value)}
                onKeyDown={e => e.key === "Enter" && confirmMarkIssued()}
                placeholder="例：AB12345678"
                autoFocus
              />
              {invoiceNo.trim() === "" && <div style={{ fontSize: 12, color: "#DC3545", marginTop: 5 }}>請填入發票號碼</div>}
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={confirmMarkIssued} disabled={!invoiceNo.trim()}
                style={{ flex: 1, border: "none", background: invoiceNo.trim() ? "#1a1a2e" : "#ccc", color: "#fff", padding: "11px", borderRadius: 8, cursor: invoiceNo.trim() ? "pointer" : "not-allowed", fontWeight: 700 }}>
                確認開立
              </button>
              <button onClick={() => { setIssueModal(null); setInvoiceNo(""); }}
                style={{ flex: 1, border: "1px solid #ddd", background: "#fff", padding: "11px", borderRadius: 8, cursor: "pointer" }}>取消</button>
            </div>
          </div>
        </div>
      )}

      {/* 刪除確認 modal */}
      {deleteConfirm && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div style={{ background: "#fff", borderRadius: 16, padding: 24, width: 360, boxShadow: "0 8px 40px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 28, textAlign: "center", marginBottom: 12 }}>🗑️</div>
            <div style={{ fontWeight: 700, fontSize: 16, textAlign: "center", marginBottom: 8 }}>確認刪除？</div>
            <div style={{ fontSize: 13, color: "#666", textAlign: "center", marginBottom: 24 }}>
              「{projects.find(p => p.id === deleteConfirm)?.title}」將永久刪除，無法復原。
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => deleteProject(deleteConfirm)}
                style={{ flex: 1, border: "none", background: "#DC3545", color: "#fff", padding: "11px", borderRadius: 8, cursor: "pointer", fontWeight: 700 }}>確認刪除</button>
              <button onClick={() => setDeleteConfirm(null)}
                style={{ flex: 1, border: "1px solid #ddd", background: "#fff", padding: "11px", borderRadius: 8, cursor: "pointer" }}>取消</button>
            </div>
          </div>
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 12 }}>
        <h2 style={{ fontSize: 22, fontWeight: 700, color: "#1a1a2e" }}>🧾 進度</h2>
        <div style={{ display: "flex", gap: 10 }}>
          <select style={{ ...inp, width: "auto" }} value={filter} onChange={e => setFilter(e.target.value)}>
            <option value="">全部狀態</option>
            {Object.keys(STATUS_COLOR).map(s => <option key={s}>{s}</option>)}
          </select>
          {(currentUser.role === "finance" || currentUser.role === "admin") &&
            <button onClick={onExport} style={{ border: "none", background: "#1a1a2e", color: "#fff", padding: "10px 18px", borderRadius: 10, cursor: "pointer", fontSize: 13 }}>
              📥 匯出Excel
            </button>}
        </div>
      </div>

      {/* ── 月份切換 Tab ── */}
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 12, overflowX: "auto", WebkitOverflowScrolling: "touch", paddingBottom: 4 }}>
        <span style={{ fontSize: 12, color: "#888", alignSelf: "center", marginRight: 4, whiteSpace: "nowrap" }}>發票開立月份：</span>
        {months.map(m => (
          <button key={m} onClick={() => { setSelectedMonth(m); setDateFrom(""); setDateTo(""); }}
            style={{
              padding: "5px 14px", borderRadius: 20, border: "none", cursor: "pointer", fontSize: 13, whiteSpace: "nowrap",
              background: selectedMonth === m && !dateFrom && !dateTo ? "#1a1a2e" : "#F0F0F0",
              color: selectedMonth === m && !dateFrom && !dateTo ? "#fff" : "#555",
              fontWeight: selectedMonth === m && !dateFrom && !dateTo ? 700 : 400,
            }}>
            {fmtMonth(m)}
          </button>
        ))}
      </div>
      {/* ── 自選日期區間 ── */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 20, flexWrap: "wrap", rowGap: 8 }}>
        <span style={{ fontSize: 12, color: "#888", whiteSpace: "nowrap" }}>自選期間：</span>
        <input type="date" style={{ ...inp, width: "auto", minWidth: 130, fontSize: 13, padding: "5px 10px" }}
          value={dateFrom} onChange={e => { setDateFrom(e.target.value); setSelectedMonth("all"); }} />
        <span style={{ fontSize: 13, color: "#aaa" }}>～</span>
        <input type="date" style={{ ...inp, width: "auto", minWidth: 130, fontSize: 13, padding: "5px 10px" }}
          value={dateTo} onChange={e => { setDateTo(e.target.value); setSelectedMonth("all"); }} />
        {(dateFrom || dateTo) && (
          <button onClick={() => { setDateFrom(""); setDateTo(""); setSelectedMonth("all"); }}
            style={{ border: "1px solid #ddd", background: "#fff", padding: "5px 12px", borderRadius: 8, cursor: "pointer", fontSize: 12, color: "#888" }}>
            清除
          </button>
        )}
      </div>

      <Card style={{ padding: 0, overflow: "visible" }}>
        <div className="rwd-scroll-container" style={{ overflowX: "auto", overflowY: "auto", maxHeight: "calc(100vh - 280px)", WebkitOverflowScrolling: "touch", borderRadius: 16 }}>
          <table style={{ borderCollapse: "collapse", fontSize: 13, tableLayout: isMobile ? "auto" : "fixed", width: isMobile ? "auto" : "100%", minWidth: isMobile ? 700 : 1590 }}>
            <thead>
              <tr style={{ background: "#F8F9FA" }}>
                <th style={{ ...sTh, left: 0,   width: 100, minWidth: 100, top: 0, zIndex: 4 }}>賣方</th>
                <th style={{ ...sTh, left: 100, width: 160, minWidth: 160, top: 0, zIndex: 4 }}>買方</th>
                <th style={{ ...sTh, left: 260, width: 100, minWidth: 100, borderRight: isMobile ? "none" : "2px solid #D0D0D0", top: 0, zIndex: 4 }}>含稅金額</th>
                {[
                  ["統編/身分證", 110], ["發票日期", 100], ["發票號碼", 120],
                  ["實際收款日", 100], ["預計收款日", 100], ["付款狀態", 120],
                  ["已收款金額", 100], ["手續費", 80], ["佣金", 80],
                  ["說明", 110], ["支出申請單編號", 130], ["傳票號碼", 110],
                  ["申請人", 90], ["操作", 130],
                ].map(([h, w]) => (
                  <th key={h} style={{ padding: "12px 14px", textAlign: "left", color: "#555", fontWeight: 600, borderBottom: "2px solid #F0F0F0", whiteSpace: "nowrap", background: "#F8F9FA", width: w, minWidth: w, position: isMobile ? "static" : "sticky", top: 0, zIndex: 3 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(p => {
                const isEditing = editingId === p.id;
                const vendor = vendors.find(v => v.id === p.vendorId);
                const rowBg = isEditing ? "#F0F7FF" : "white";
                return (
                  <tr key={p.id} style={{ borderBottom: "1px solid #F4F4F4", background: rowBg, cursor: isEditing ? "default" : "pointer" }} onClick={() => { if (!isEditing) setDetailModal(p); }}>
                    {isEditing ? (
                      <>
                        <td style={{ ...sTd, left: 0,   width: 100, minWidth: 100, background: rowBg }}>
                          <select style={{ ...inp, padding: "5px 8px", fontSize: 12 }} value={ef.company} onChange={e => setEf("company", e.target.value)}>
                            {COMPANIES.map(c => <option key={c}>{c}</option>)}
                          </select>
                        </td>
                        <td style={{ ...sTd, left: 100, width: 160, minWidth: 160, background: rowBg }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12 }} value={ef.clientTitle || ""} onChange={e => setEf("clientTitle", e.target.value)} placeholder="買方名稱" />
                        </td>
                        <td style={{ ...sTd, left: 260, width: 100, minWidth: 100, background: rowBg, borderRight: "2px solid #D0D0D0" }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12, width: 80 }} type="number" value={ef.taxAmount} onChange={e => setEf("taxAmount", +e.target.value)} />
                        </td>
                        <td style={{ ...td, width: 110, minWidth: 110 }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12 }} value={ef.taxId || ""} onChange={e => setEf("taxId", e.target.value)} placeholder="統編" />
                        </td>
                        <td style={{ ...td, width: 100, minWidth: 100 }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12, width: "100%" }} type="date" value={ef.invoiceDate} onChange={e => setEf("invoiceDate", e.target.value)} />
                        </td>
                        <td style={{ ...td, width: 120, minWidth: 120 }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12 }} value={ef.invoiceNo || ""} onChange={e => setEf("invoiceNo", e.target.value)} placeholder="發票號碼" />
                        </td>
                        <td style={{ ...td, width: 100, minWidth: 100 }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12, width: "100%" }} type="date" value={ef.paidDate || ""} onChange={e => setEf("paidDate", e.target.value)} />
                        </td>
                        <td style={{ ...td, width: 100, minWidth: 100 }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12, width: "100%" }} type="date" value={ef.expectedPayDate || ""} onChange={e => setEf("expectedPayDate", e.target.value)} />
                        </td>
                        <td style={{ ...td, width: 120, minWidth: 120 }}>
                          <select style={{ ...inp, padding: "5px 8px", fontSize: 12, width: "100%" }} value={ef.status} onChange={e => setEf("status", e.target.value)}>
                            {["申請中","已開立","到期未付款","已付款"].map(s => <option key={s}>{s}</option>)}
                          </select>
                        </td>
                        <td style={{ ...td, width: 100, minWidth: 100 }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12, width: "100%", textAlign: "right" }} type="text" inputMode="numeric"
                            value={ef.paidAmount || ""} onChange={e => { const v = parseInt(e.target.value.replace(/\D/g,""),10); setEf("paidAmount", isNaN(v) ? "" : v); }} placeholder="金額" />
                        </td>
                        <td style={{ ...td, width: 80, minWidth: 80 }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12, width: "100%", textAlign: "right" }} type="text" inputMode="numeric"
                            value={ef.bankFee || ""} onChange={e => { const v = parseInt(e.target.value.replace(/\D/g,""),10); setEf("bankFee", isNaN(v) ? "" : v); }} placeholder="金額" />
                        </td>
                        <td style={{ ...td, width: 80, minWidth: 80 }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12, width: "100%", textAlign: "right" }} type="text" inputMode="numeric"
                            value={ef.commission || ""} onChange={e => { const v = parseInt(e.target.value.replace(/\D/g,""),10); setEf("commission", isNaN(v) ? "" : v); }} placeholder="金額" />
                        </td>
                        <td style={{ ...td, width: 110, minWidth: 110 }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12, width: "100%" }} value={ef.commissionNo || ""} onChange={e => setEf("commissionNo", e.target.value)} placeholder="說明" />
                        </td>
                        <td style={{ ...td, width: 130, minWidth: 130 }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12, width: "100%" }} value={ef.expenseNo || ""} onChange={e => setEf("expenseNo", e.target.value)} placeholder="支出申請單編號" />
                        </td>
                        <td style={{ ...td, width: 110, minWidth: 110 }}>
                          <input style={{ ...inp, padding: "5px 8px", fontSize: 12, width: "100%" }} value={ef.voucherNo || ""} onChange={e => setEf("voucherNo", e.target.value)} placeholder="傳票號碼" />
                        </td>
                        <td style={{ ...td, width: 90, minWidth: 90 }}>
                          <select style={{ ...inp, padding: "4px 6px", fontSize: 12, width: "100%" }} value={ef.applicant} onChange={e => setEf("applicant", e.target.value)}>
                            {!users.some(u => u.name === ef.applicant) && (<option value={ef.applicant}>{ef.applicant}</option>)}
                            {users.map(u => <option key={u.id} value={u.name}>{u.name}</option>)}
                          </select>
                        </td>
                        <td style={{ ...td, width: 130, minWidth: 130, whiteSpace: "nowrap" }}>
                          <button onClick={saveEdit} style={{ border: "none", background: "#1a1a2e", color: "#fff", padding: "5px 12px", borderRadius: 6, cursor: "pointer", fontSize: 12, marginRight: 6 }}>儲存</button>
                          <button onClick={cancelEdit} style={{ border: "1px solid #ddd", background: "#fff", padding: "5px 10px", borderRadius: 6, cursor: "pointer", fontSize: 12 }}>取消</button>
                        </td>
                      </>
                    ) : (
                      // ── 一般列 ──────────────────────────────────────────
                      <>
                        <td style={{ ...sTd, left: 0,   width: 100, minWidth: 100, background: rowBg }}>
                          <span style={{ background: "#E8EAF6", color: "#3949AB", padding: "2px 8px", borderRadius: 6, fontSize: 12 }}>{p.company}</span>
                        </td>
                        <td style={{ ...sTd, left: 100, width: 160, minWidth: 160, background: rowBg, fontSize: 12 }}>{p.clientTitle || "-"}</td>
                        <td style={{ ...sTd, left: 260, width: 100, minWidth: 100, background: rowBg, borderRight: "2px solid #D0D0D0", fontWeight: 700, color: "#2E7D32", whiteSpace: "nowrap" }}>{fmt(p.taxAmount)}</td>
                        <td style={{ ...td, fontSize: 12, whiteSpace: "nowrap", color: "#666" }}>{p.taxId || "-"}</td>
                        <td style={{ ...td, whiteSpace: "nowrap" }}>{p.invoiceDate || "-"}</td>
                        <td style={{ ...td, fontSize: 12, fontFamily: "monospace", color: p.invoiceNo ? "#1a1a2e" : "#bbb" }}>{p.invoiceNo || "-"}</td>
                        <td style={{ ...td, whiteSpace: "nowrap", color: p.paidDate ? "#1a1a2e" : "#bbb" }}>{parseRocDate(p.paidDate) || "-"}</td>
                        <td style={{ ...td, whiteSpace: "nowrap", color: p.expectedPayDate ? "#1a1a2e" : "#bbb" }}>{p.expectedPayDate || "-"}</td>
                        <td style={td}><Badge status={p._effectiveStatus} /></td>
                        <td style={{ ...td, fontWeight: p.paidAmount ? 700 : 400, color: p.paidAmount ? "#2E7D32" : "#bbb", whiteSpace: "nowrap" }}>{p.paidAmount ? fmt(p.paidAmount) : "-"}</td>
                        <td style={{ ...td, color: p.bankFee ? "#E65100" : "#bbb", whiteSpace: "nowrap" }}>{p.bankFee ? fmt(p.bankFee) : "-"}</td>
                        <td style={{ ...td, color: p.commission ? "#E65100" : "#bbb", whiteSpace: "nowrap" }}>{p.commission ? fmt(p.commission) : "-"}</td>
                        <td style={{ ...td, fontSize: 12, color: p.commissionNo ? "#1a1a2e" : "#bbb" }}>{p.commissionNo || "-"}</td>
                        <td style={{ ...td, fontSize: 12, color: p.expenseNo ? "#1a1a2e" : "#bbb" }}>{p.expenseNo || "-"}</td>
                        <td style={{ ...td, fontSize: 12, fontFamily: "monospace", color: p.voucherNo ? "#1a1a2e" : "#bbb" }}>{p.voucherNo || "-"}</td>
                        <td style={{ ...td, fontSize: 12 }}>{p.applicant}</td>
                        <td style={{ ...td, whiteSpace: "nowrap" }}>
                          <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                            {/* 標記已開立：申請中才顯示 */}
                            {p._effectiveStatus === "申請中" && (
                              <button onClick={e => { e.stopPropagation(); setIssueModal({ id: p.id }); setInvoiceNo(""); }}
                                style={{ border: "1px solid #007BFF", background: "#EBF5FF", color: "#0056b3", padding: "4px 10px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 600, whiteSpace: "nowrap" }}>
                                ✓ 標記已開立
                              </button>
                            )}
                            {/* 標記已付款：已開立或到期未付款才顯示 */}
                            {(p._effectiveStatus === "已開立" || p._effectiveStatus === "到期未付款") && (
                              <button onClick={e => { e.stopPropagation(); setProjects(projects.map(x => x.id === p.id ? { ...x, status: "已付款" } : x)); }}
                                style={{ border: "1px solid #28A745", background: "#EAFBF0", color: "#155724", padding: "4px 10px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 600, whiteSpace: "nowrap" }}>
                                💰 標記已付款
                              </button>
                            )}
                            {/* 複製申請 */}
                            <button onClick={e => { e.stopPropagation(); onCopyProject(p); }}
                              style={{ border: "1px solid #F2A0A0", background: "#FFF5F5", color: "#c0566a", padding: "4px 10px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 600, whiteSpace: "nowrap" }}>
                              📋 複製
                            </button>
                            {/* 編輯：財務/管理員 */}
                            {canEdit && (
                              <button onClick={e => { e.stopPropagation(); startEdit(p); }}
                                style={{ border: "1px solid #ddd", background: "#F8F9FA", color: "#444", padding: "4px 10px", borderRadius: 6, cursor: "pointer", fontSize: 12 }}>
                                ✏️
                              </button>
                            )}
                            {/* 刪除：財務/管理員 */}
                            {canEdit && (
                              <button onClick={e => { e.stopPropagation(); setDeleteConfirm(p.id); }}
                                style={{ border: "1px solid #fcc", background: "#fff5f5", color: "#c33", padding: "4px 10px", borderRadius: 6, cursor: "pointer", fontSize: 12 }}>
                                🗑
                              </button>
                            )}
                          </div>
                        </td>
                      </>
                    )}
                  </tr>
                );
              })}
              {filtered.length === 0 && (
                <tr><td colSpan={16} style={{ textAlign: "center", color: "#aaa", padding: 32 }}>尚無符合條件的紀錄</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ─── USER MANAGEMENT ─────────────────────────────────────────────────────────
function UserManagement({ users, setUsers }) {
  const emptyForm = { name: "", email: "", role: "applicant", company: "紳太", active: true, password: "" };
  const [form, setForm] = useState(emptyForm);
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [pwModal, setPwModal] = useState(null); // { userId }
  const [newPw, setNewPw] = useState("");
  const [showPw, setShowPw] = useState(false);

  const openNew = () => { setEditingId(null); setForm(emptyForm); setShowForm(true); };
  const openEdit = (u) => {
    setEditingId(u.id);
    setForm({ name: u.name, email: u.email, role: u.role, company: u.company, active: u.active, password: u.password || "" });
    setShowForm(true);
  };
  const save = () => {
    if (editingId !== null) {
      setUsers(users.map(u => u.id === editingId ? { ...u, ...form } : u));
    } else {
      setUsers([...users, { id: Date.now(), ...form }]);
    }
    setShowForm(false); setEditingId(null); setForm(emptyForm);
  };
  const savePw = () => {
    if (!newPw.trim()) return;
    setUsers(users.map(u => u.id === pwModal.userId ? { ...u, password: newPw.trim() } : u));
    setPwModal(null); setNewPw(""); setShowPw(false);
  };

  return (
    <div>
      {/* 重設密碼 Modal */}
      {pwModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div style={{ background: "#fff", borderRadius: 16, padding: 24, width: 380, boxShadow: "0 8px 40px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 26, textAlign: "center", marginBottom: 10 }}>🔑</div>
            <div style={{ fontWeight: 700, fontSize: 16, textAlign: "center", marginBottom: 6 }}>重設密碼</div>
            <div style={{ fontSize: 13, color: "#888", textAlign: "center", marginBottom: 20 }}>
              {users.find(u => u.id === pwModal.userId)?.name}
            </div>
            <div style={{ marginBottom: 20 }}>
              <label style={lbl}>新密碼</label>
              <div style={{ position: "relative" }}>
                <input style={{ ...inp, paddingRight: 44 }}
                  type={showPw ? "text" : "password"}
                  value={newPw} onChange={e => setNewPw(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && savePw()}
                  placeholder="請輸入新密碼" autoFocus />
                <span onClick={() => setShowPw(v => !v)}
                  style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", cursor: "pointer", fontSize: 16, color: "#aaa", userSelect: "none" }}>
                  {showPw ? "🙈" : "👁"}
                </span>
              </div>
              {newPw.trim() === "" && <div style={{ fontSize: 12, color: "#DC3545", marginTop: 5 }}>請填入新密碼</div>}
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={savePw} disabled={!newPw.trim()}
                style={{ flex: 1, border: "none", background: newPw.trim() ? "#1a1a2e" : "#ccc", color: "#fff", padding: "11px", borderRadius: 8, cursor: newPw.trim() ? "pointer" : "not-allowed", fontWeight: 700 }}>
                確認更新
              </button>
              <button onClick={() => { setPwModal(null); setNewPw(""); setShowPw(false); }}
                style={{ flex: 1, border: "1px solid #ddd", background: "#fff", padding: "11px", borderRadius: 8, cursor: "pointer" }}>取消</button>
            </div>
          </div>
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
        <h2 style={{ fontSize: 22, fontWeight: 700, color: "#1a1a2e" }}>👥 帳號管理</h2>
        <button onClick={openNew} style={{ border: "none", background: "#1a1a2e", color: "#fff", padding: "10px 20px", borderRadius: 10, cursor: "pointer" }}>＋ 新增帳號</button>
      </div>

      {showForm && (
        <Card style={{ marginBottom: 20, borderColor: "#1a1a2e33" }}>
          <h3 style={{ ...sectionTitle, marginBottom: 16 }}>{editingId ? "✏️ 編輯帳號" : "＋ 新增帳號"}</h3>
          <div className="rwd-grid2" style={grid2}>
            <div style={field}><label style={lbl}>姓名</label><input style={inp} value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} /></div>
            <div style={field}><label style={lbl}>Email</label><input style={inp} value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} /></div>
            <div style={field}><label style={lbl}>角色</label>
              <select style={inp} value={form.role} onChange={e => setForm({ ...form, role: e.target.value })}>
                {Object.entries(ROLES).map(([k,v]) => <option key={k} value={k}>{v}</option>)}
              </select>
            </div>
            <div style={field}><label style={lbl}>公司</label>
              <select style={inp} value={form.company} onChange={e => setForm({ ...form, company: e.target.value })}>
                {COMPANIES.map(c => <option key={c}>{c}</option>)}
                <option value="所有">所有</option>
              </select>
            </div>
            <div style={field}>
              <label style={lbl}>密碼{editingId ? "（留空表示不變更）" : ""}</label>
              <div style={{ position: "relative" }}>
                <input style={{ ...inp, paddingRight: 44 }}
                  type={showPw ? "text" : "password"}
                  value={form.password} onChange={e => setForm({ ...form, password: e.target.value })}
                  placeholder={editingId ? "留空則不變更密碼" : "請設定密碼"} />
                <span onClick={() => setShowPw(v => !v)}
                  style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", cursor: "pointer", fontSize: 16, color: "#aaa", userSelect: "none" }}>
                  {showPw ? "🙈" : "👁"}
                </span>
              </div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
            <button onClick={save} style={{ border: "none", background: "#1a1a2e", color: "#fff", padding: "10px 20px", borderRadius: 8, cursor: "pointer" }}>
              {editingId ? "儲存變更" : "新增"}
            </button>
            <button onClick={() => { setShowForm(false); setEditingId(null); }} style={{ border: "1px solid #ddd", padding: "10px 20px", borderRadius: 8, cursor: "pointer" }}>取消</button>
          </div>
        </Card>
      )}

      <Card>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#F8F9FA" }}>
              {["姓名","Email","角色","公司","密碼","狀態","操作"].map(h => (
                <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: "#555", fontWeight: 600, borderBottom: "2px solid #F0F0F0" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {users.map((u, i) => (
              <tr key={u.id} style={{ borderBottom: "1px solid #F8F8F8" }}>
                <td style={{ ...td, fontWeight: 600 }}>{u.name}</td>
                <td style={td}>{u.email}</td>
                <td style={td}><span style={{ background: "#E8F5E9", color: "#2E7D32", padding: "2px 8px", borderRadius: 6, fontSize: 12 }}>{ROLES[u.role]}</span></td>
                <td style={td}>{u.company}</td>
                <td style={td}>
                  {u.password
                    ? <span style={{ fontFamily: "monospace", letterSpacing: 2, color: "#888" }}>{"●".repeat(Math.min(u.password.length, 8))}</span>
                    : <span style={{ color: "#ccc", fontSize: 12 }}>未設定</span>}
                </td>
                <td style={td}><span style={{ color: u.active ? "#4CAF50" : "#F44336", fontWeight: 700 }}>{u.active ? "● 啟用" : "● 停用"}</span></td>
                <td style={{ ...td, whiteSpace: "nowrap" }}>
                  <div style={{ display: "flex", gap: 6 }}>
                    <button onClick={() => openEdit(u)}
                      style={{ border: "1px solid #ddd", background: "#F8F9FA", color: "#444", padding: "4px 10px", borderRadius: 6, cursor: "pointer", fontSize: 12 }}>
                      ✏️ 編輯
                    </button>
                    <button onClick={() => { setPwModal({ userId: u.id }); setNewPw(""); setShowPw(false); }}
                      style={{ border: "1px solid #C8DCF7", background: "#EBF5FF", color: "#0056b3", padding: "4px 10px", borderRadius: 6, cursor: "pointer", fontSize: 12 }}>
                      🔑 密碼
                    </button>
                    <button onClick={() => setUsers(users.map((x, j) => j === i ? { ...x, active: !x.active } : x))}
                      style={{ border: `1px solid ${u.active ? "#fcc" : "#c8e6c9"}`, background: u.active ? "#fff5f5" : "#f0fff4", color: u.active ? "#c33" : "#2E7D32", padding: "4px 10px", borderRadius: 6, cursor: "pointer", fontSize: 12 }}>
                      {u.active ? "停用" : "啟用"}
                    </button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ─── 新竹物流 申請&對帳 ───────────────────────────────────────────────────────
// ─── 新竹物流部門清單 ─────────────────────────────────────────────────────────
const DEPTS=["管理部","財務部","和和研","廣告業務部","社群部","設計部","資訊部","匯太","董事長室","網友自付運費","其他"];

function HsinchuApp({ currentUser }) {
  const [tab,setTab]=useState("apply");
  const [toast,setToast]=useState(null);
  const [applications,setApplications]=useState([]);
  const [billRows,setBillRows]=useState([]);
  const [selectedMonth,setSelectedMonth]=useState("all");
  const emptyApply={dept:"",reason:"",month:"",fileNames:[],rows:[]};
  const [applyForm,setApplyForm]=useState(emptyApply);
  const [applyLoading,setApplyLoading]=useState(false);
  const [billLoading,setBillLoading]=useState(false);
  const nowYM=()=>{const n=new Date();return`${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,"0")}`;};
  const [billMonth,setBillMonth]=useState(nowYM);

  const showToast=(msg,type="success")=>{setToast({msg,type});setTimeout(()=>setToast(null),3000);};

  // 月份清單從手動標記的 billTag 來
  const months=useMemo(()=>{
    const set=new Set(billRows.map(r=>r.billTag).filter(Boolean));
    return["all",...Array.from(set).sort((a,b)=>b.localeCompare(a))];
  },[billRows]);

  const fmtMonth=m=>m==="all"?"全部":`${m.split("-")[0]}年${parseInt(m.split("-")[1])}月`;

  const filteredBills=useMemo(()=>{
    if(selectedMonth==="all")return billRows;
    return billRows.filter(r=>r.billTag===selectedMonth);
  },[billRows,selectedMonth]);

  const loadXLSX=()=>new Promise((res,rej)=>{
    if(window.XLSX)return res();
    const s=document.createElement("script");
    s.src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s.onload=res;s.onerror=rej;document.head.appendChild(s);
  });

  const readFile=async(file)=>{
    await loadXLSX();
    const ab=await new Promise((res,rej)=>{const r=new FileReader();r.onload=e=>res(e.target.result);r.onerror=rej;r.readAsArrayBuffer(file);});
    const isCSV=file.name.toLowerCase().endsWith(".csv");
    let wb;
    if(isCSV){
      let text=new TextDecoder("utf-8").decode(new Uint8Array(ab));
      if(!text.includes("碼")&&!text.includes("貨")&&!text.includes("收"))text=new TextDecoder("big5").decode(new Uint8Array(ab));
      wb=window.XLSX.read(text,{type:"string"});
    }else{wb=window.XLSX.read(new Uint8Array(ab),{type:"array"});}
    return window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:null,raw:true});
  };

  const norm=v=>String(v??"").replace(/[^\d]/g,"").trim();

  const parseManifest=async(file)=>{
    const data=await readFile(file);const hdr=data[0]||[];
    const col=(e,fb)=>{let i=hdr.findIndex(h=>h&&String(h).trim()===e);if(i<0&&fb)i=hdr.findIndex(h=>h&&String(h).includes(fb));return i;};
    const cNo=col("十碼貨號"),cRec=col("收貨人"),cNote=col("備註"),cPcs=col("件數"),cDate=col("發貨日期","建單時間");
    return data.slice(1).filter(r=>r&&r[cNo]!=null&&String(r[cNo]).trim()!=="").map(r=>({
      shipNo:norm(r[cNo]),recipient:String(r[cRec]??"").trim(),note:String(r[cNote]??"").trim(),pieces:parseInt(r[cPcs])||1,date:String(r[cDate]??"").trim()
    }));
  };

  const parseBill=async(file)=>{
    const data=await readFile(file);const hdr=data[0]||[];
    const col=(e,fb)=>{let i=hdr.findIndex(h=>h&&String(h).trim()===e);if(i<0&&fb)i=hdr.findIndex(h=>h&&String(h).includes(fb));return i;};
    const cShip=col("明細表號"),cApply=col("客戶用傳票_管理號碼","客戶用傳票"),cRec=col("收貨人"),cDate=col("發送日"),cAmt=col("請款金額"),cPcs=col("總件數");
    return data.slice(1).filter(r=>r&&r[cShip]!=null&&String(r[cShip]).trim()!=="").map(r=>({
      shipNo:norm(r[cShip]),applyId:String(r[cApply]??"").trim(),recipient:String(r[cRec]??"").trim(),
      sendDate:String(r[cDate]??"").trim(),amount:parseFloat(String(r[cAmt]??"0").replace(/,/g,""))||0,
      pieces:parseFloat(r[cPcs])||0,matchedApp:null,matchedDept:"",matchedReason:""
    }));
  };

  const handleManifestUpload=async(e)=>{
    const files=Array.from(e.target.files);if(!files.length)return;
    setApplyLoading(true);
    try{
      let all=[];for(const f of files)all=[...all,...await parseManifest(f)];
      setApplyForm(f=>({...f,fileNames:[...f.fileNames,...files.map(f=>f.name)],rows:[...f.rows,...all]}));
      showToast(`✅ 讀取 ${files.length} 個檔案，共 ${all.length} 筆`);
    }catch(err){showToast("❌ "+err.message,"error");}
    setApplyLoading(false);e.target.value="";
  };

  const submitApply=()=>{
    if(!applyForm.dept||!applyForm.reason){showToast("請填寫部門與使用原因","error");return;}
    if(!applyForm.month){showToast("請填寫托運月份","error");return;}
    if(!applyForm.rows.length){showToast("請先上傳托運總表","error");return;}
    const id=`HSIN-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    const a={id,dept:applyForm.dept,reason:applyForm.reason,month:applyForm.month,fileName:applyForm.fileNames.join("、"),rows:applyForm.rows,applicant:"使用者",createdAt:new Date().toLocaleDateString("zh-TW"),status:"待核對"};
    setApplications(prev=>[a,...prev]);setApplyForm(emptyApply);
    showToast(`✅ 申請單 ${id} 已建立，含 ${a.rows.length} 筆貨號`);
  };

  const handleBillUpload=async(e)=>{
    const file=e.target.files[0];if(!file)return;
    if(!billMonth){showToast("請先選擇帳單月份","error");return;}
    setBillLoading(true);
    try{
      const rows=await parseBill(file);
      const matched=rows.map(r=>{
        const key=norm(r.shipNo);
        if(r.applyId&&r.applyId!=="null"&&r.applyId.trim()!==""){const app=applications.find(a=>a.id===r.applyId.trim());if(app)return{...r,matchedApp:app.id,matchedDept:app.dept,matchedReason:app.reason};}
        for(const app of applications)if(app.rows.some(row=>norm(row.shipNo)===key))return{...r,matchedApp:app.id,matchedDept:app.dept,matchedReason:app.reason};
        return r;
      });
      const existing=new Set(billRows.map(r=>r.shipNo));
      // 標記帳單月份 billTag
      const newRows=matched.filter(r=>!existing.has(r.shipNo)).map(r=>({...r,billTag:billMonth}));
      setBillRows(prev=>[...prev,...newRows]);
      setSelectedMonth(billMonth); // 上傳後自動切換到該月份
      showToast(`✅ 匯入 ${newRows.length} 筆（${fmtMonth(billMonth)}），${matched.filter(r=>r.matchedApp).length} 筆自動核對`);
    }catch(err){showToast("❌ "+err.message,"error");}
    setBillLoading(false);e.target.value="";
  };

  const updateBillRow=(idx,patch)=>{
    const ship=filteredBills[idx]?.shipNo;
    setBillRows(prev=>prev.map(r=>r.shipNo===ship?{...r,...patch}:r));
  };

  const report=useMemo(()=>{
    const map={};
    filteredBills.forEach(r=>{
      const dept=r.matchedDept||"未分類",reason=r.matchedReason||r.applyId||"未指定";
      if(!map[dept])map[dept]={};map[dept][reason]=(map[dept][reason]||0)+r.amount;
    });
    return map;
  },[filteredBills]);

  const totalAmt=filteredBills.reduce((s,r)=>s+r.amount,0);

  const copyReport=()=>{
    const lines=[];
    Object.entries(report).forEach(([dept,reasons])=>{
      const t=Object.values(reasons).reduce((s,v)=>s+v,0);
      lines.push(`〈${dept}〉NT$ ${t.toLocaleString()}`);
      Object.entries(reasons).forEach(([r,a])=>lines.push(`-${r} -> NT$ ${a.toLocaleString()}`));
      lines.push("");
    });
    lines.push(`總金額 NT$ ${totalAmt.toLocaleString()}`);
    navigator.clipboard.writeText(lines.join("\n")).then(()=>showToast("✅ 已複製")).catch(()=>showToast("❌ 複製失敗","error"));
  };

  const tabStyle=t=>({padding:"10px 18px",border:"none",cursor:"pointer",fontWeight:tab===t?700:400,fontSize:14,background:"none",color:tab===t?THEME:"#888",borderBottom:tab===t?`2px solid ${THEME}`:"2px solid transparent",marginBottom:-2,whiteSpace:"nowrap"});

  const MonthBar=()=>months.length<=1?null:(
    <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:16,alignItems:"center"}}>
      <span style={{fontSize:12,color:"#888"}}>月份：</span>
      {months.map(m=>(
        <button key={m} onClick={()=>setSelectedMonth(m)} style={{padding:"4px 12px",borderRadius:20,border:"none",cursor:"pointer",fontSize:12,background:selectedMonth===m?"#1a1a2e":"#F0F0F0",color:selectedMonth===m?"#fff":"#555",fontWeight:selectedMonth===m?700:400}}>
          {fmtMonth(m)}
        </button>
      ))}
    </div>
  );

  return(
    <div style={{minHeight:"100vh",background:"#F5F6FA",fontFamily:"sans-serif"}}>
      <div style={{background:THEME,color:"#fff",padding:"0 16px",height:52,display:"flex",alignItems:"center",gap:10,position:"sticky",top:0,zIndex:10}}>
        <span style={{fontWeight:900,fontSize:18}}>媽咪拜</span>
        <span style={{opacity:0.5}}>|</span>
        <span style={{fontSize:13}}>🏙️ 新竹申請&對帳</span>
      </div>
      <div style={{maxWidth:860,margin:"0 auto",padding:"20px 16px 80px"}}>
        <div style={{display:"flex",borderBottom:"2px solid #F0F0F0",marginBottom:20,overflowX:"auto"}}>
          <button style={tabStyle("apply")} onClick={()=>setTab("apply")}>📝 使用申請</button>
          <button style={tabStyle("bill")} onClick={()=>setTab("bill")}>📄 全部帳單{billRows.length>0&&<span style={{marginLeft:6,background:THEME,color:"#fff",borderRadius:10,padding:"1px 7px",fontSize:11}}>{billRows.length}</span>}</button>
          <button style={tabStyle("report")} onClick={()=>setTab("report")}>📊 費用報告</button>
        </div>

        {/* 使用申請 */}
        {tab==="apply"&&<div>
          <_HCard style={{marginBottom:20}}>
            <div style={{fontSize:16,fontWeight:700,marginBottom:16}}>新增使用申請</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
              <div style={_field}><label style={_lbl}>部門</label>
                <select style={_inp} value={applyForm.dept} onChange={e=>setApplyForm(f=>({...f,dept:e.target.value}))}>
                  <option value="">— 請選擇部門 —</option>{DEPTS.map(d=><option key={d}>{d}</option>)}
                </select>
              </div>
              <div style={_field}><label style={_lbl}>專案編號 / 使用原因</label>
                <input style={_inp} value={applyForm.reason} onChange={e=>setApplyForm(f=>({...f,reason:e.target.value}))} placeholder="例：公關品寄送"/>
              </div>
              <div style={_field}><label style={_lbl}>托運月份</label>
                <input type="month" style={_inp} value={applyForm.month} onChange={e=>setApplyForm(f=>({...f,month:e.target.value}))}/>
              </div>
            </div>
            <div style={{marginBottom:16}}>
              <label style={_lbl}>上傳托運總表（可多選，.xlsx / .xls / .csv）</label>
              <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                <label style={{border:`2px dashed ${THEME}`,borderRadius:10,padding:"10px 18px",cursor:"pointer",color:THEME,fontSize:13,fontWeight:600,background:"#EBF4FB"}}>
                  📎 選擇檔案<input type="file" multiple accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleManifestUpload}/>
                </label>
                {applyForm.fileNames.length>0&&<button onClick={()=>setApplyForm(f=>({...f,fileNames:[],rows:[]}))} style={{border:"1px solid #fcc",background:"#fff5f5",color:"#c33",padding:"8px 14px",borderRadius:8,cursor:"pointer",fontSize:12}}>🗑 清除</button>}
              </div>
              {applyLoading&&<div style={{marginTop:8,color:"#888",fontSize:13}}>⏳ 讀取中...</div>}
              {applyForm.fileNames.length>0&&<div style={{marginTop:10,background:"#F0FFF4",border:"1px solid #C8E6C9",padding:"10px 14px",borderRadius:8,fontSize:13,lineHeight:1.9}}>
                <div style={{fontWeight:600,color:"#2E7D32"}}>✅ 共 {applyForm.rows.length} 筆</div>
                {applyForm.fileNames.map((n,i)=><div key={i} style={{fontSize:12,color:"#555"}}>📄 {n}</div>)}
              </div>}
            </div>
            <button onClick={submitApply} style={{border:"none",background:THEME,color:"#fff",padding:"12px 28px",borderRadius:10,fontWeight:700,cursor:"pointer",fontSize:14}}>提交申請單</button>
          </_HCard>
          {applications.length>0&&<_HCard>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <div style={{fontSize:16,fontWeight:700}}>申請記錄（{applications.length} 筆）</div>
              <button onClick={()=>setApplications([])} style={{border:"1px solid #fcc",background:"#fff5f5",color:"#c33",padding:"4px 10px",borderRadius:8,cursor:"pointer",fontSize:12}}>清除</button>
            </div>
            {applications.map(a=>(
              <div key={a.id} style={{borderBottom:"1px solid #F0F0F0",padding:"12px 0",display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
                <div>
                  <div style={{fontWeight:700,fontSize:13,fontFamily:"monospace",color:THEME}}>{a.id}</div>
                  <div style={{fontSize:13,color:"#555",marginTop:2}}><span style={{background:"#E3F2FD",color:"#1565C0",padding:"1px 8px",borderRadius:12,fontSize:12,marginRight:6}}>{a.dept}</span>{a.reason}</div>
                  <div style={{fontSize:12,color:"#aaa",marginTop:3}}>{a.createdAt}{a.month&&<span style={{marginLeft:6,background:"#F0F4FF",color:"#5C6BC0",padding:"1px 7px",borderRadius:10,fontSize:11}}>{a.month.replace("-","年")}月</span>} · {a.rows.length} 筆貨號</div>
                </div>
                <span style={{fontSize:12,background:"#E8F5E9",color:"#2E7D32",padding:"3px 10px",borderRadius:20}}>{a.status}</span>
              </div>
            ))}
          </_HCard>}
        </div>}

        {/* 全部帳單 */}
        {tab==="bill"&&<div>
          <_HCard style={{marginBottom:16}}>
            <div style={{fontSize:16,fontWeight:700,marginBottom:12}}>上傳新竹帳單</div>
            <div style={{fontSize:13,color:"#888",marginBottom:14}}>格式：第二步_帳單（.xlsx / .xls / .csv）</div>
            {/* 月份選擇 */}
            <div style={{marginBottom:16}}>
              <label style={_lbl}>帳單月份</label>
              <input type="month" style={{..._inp,width:"auto",minWidth:160}} value={billMonth} onChange={e=>setBillMonth(e.target.value)}/>
            </div>
            <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
              <label style={{background:THEME,color:"#fff",padding:"11px 22px",borderRadius:8,cursor:"pointer",fontWeight:600,fontSize:14}}>
                📂 選擇帳單檔案<input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleBillUpload}/>
              </label>
              {billRows.length>0&&<button onClick={()=>{setBillRows([]);setSelectedMonth("all");}} style={{border:"1px solid #fcc",background:"#fff5f5",color:"#c33",padding:"10px 16px",borderRadius:8,cursor:"pointer",fontSize:13}}>🗑 清除帳單</button>}
              {billLoading&&<span style={{color:"#888",fontSize:13}}>⏳ 解析中...</span>}
            </div>
          </_HCard>
          {billRows.length>0&&<_HCard style={{padding:0}}>
            <div style={{padding:"16px 20px 12px"}}>
              <div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:12}}>
                <div style={{fontWeight:700,fontSize:15}}>核對明細</div>
                <div style={{fontSize:13}}>已核對 <b style={{color:"#2E7D32"}}>{filteredBills.filter(r=>r.matchedDept).length}</b> ／ 待核對 <b style={{color:"#E65100"}}>{filteredBills.filter(r=>!r.matchedDept).length}</b>{selectedMonth!=="all"&&<span style={{color:"#aaa",marginLeft:8}}>（{fmtMonth(selectedMonth)}）</span>}</div>
              </div>
              <MonthBar/>
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,minWidth:700}}>
                <thead><tr style={{background:"#F8F9FA"}}>{["發送日","貨運編號","收貨人","金額","部門","使用原因"].map(h=><th key={h} style={{padding:"10px 12px",textAlign:"left",color:"#555",fontWeight:600,borderBottom:"2px solid #F0F0F0",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {filteredBills.map((r,i)=>(
                    <tr key={r.shipNo} style={{borderBottom:"1px solid #F8F8F8",background:r.matchedDept?"#F0FFF4":"#FFF5F5"}}>
                      <td style={{..._td,fontSize:12,whiteSpace:"nowrap"}}>{r.sendDate}</td>
                      <td style={{..._td,fontFamily:"monospace",fontSize:11}}>{r.shipNo}</td>
                      <td style={{..._td,fontSize:12}}>{r.recipient}</td>
                      <td style={{..._td,fontWeight:700,color:"#2E7D32",whiteSpace:"nowrap"}}>NT$ {r.amount.toLocaleString()}</td>
                      <td style={{..._td,minWidth:110}}>
                        {r.matchedApp?<span style={{background:"#E3F2FD",color:"#1565C0",padding:"3px 10px",borderRadius:12,fontSize:12,fontWeight:600,whiteSpace:"nowrap"}}>✓ {r.matchedDept}</span>
                        :<select style={{..._inp,padding:"4px 6px",fontSize:12,borderColor:"#FFC0C0",background:"#FFF5F5"}} value={r.matchedDept} onChange={e=>updateBillRow(i,{matchedDept:e.target.value})}><option value="">— 部門 —</option>{DEPTS.map(d=><option key={d}>{d}</option>)}</select>}
                      </td>
                      <td style={{..._td,minWidth:130}}>
                        {r.matchedApp?<span style={{fontSize:13,color:"#1565C0"}}>{r.matchedReason}</span>
                        :<input style={{..._inp,padding:"4px 6px",fontSize:12}} value={r.matchedReason} onChange={e=>updateBillRow(i,{matchedReason:e.target.value})} placeholder="使用原因"/>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </_HCard>}
        </div>}

        {/* 費用報告 */}
        {tab==="report"&&<div>
          {billRows.length===0?<_HCard><div style={{textAlign:"center",color:"#aaa",padding:32}}>請先上傳帳單並完成核對</div></_HCard>:<>
            <MonthBar/>
            {Object.entries(report).map(([dept,reasons])=>{
              const t=Object.values(reasons).reduce((s,v)=>s+v,0);
              return <_HCard key={dept} style={{marginBottom:12}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}><span style={{fontWeight:700,fontSize:15}}>〈{dept}〉</span><span style={{fontWeight:700,color:THEME,fontSize:15}}>NT$ {t.toLocaleString()}</span></div>
                {Object.entries(reasons).map(([r,a])=><div key={r} style={{display:"flex",justifyContent:"space-between",fontSize:13,color:"#555",marginBottom:4}}><span>－{r}</span><span style={{fontWeight:600}}>NT$ {a.toLocaleString()}</span></div>)}
              </_HCard>;
            })}
            <_HCard style={{background:"#1a1a2e",color:"#fff",marginBottom:16}}>
              <div style={{display:"flex",justifyContent:"space-between"}}>
                <span style={{fontWeight:700,fontSize:16}}>總金額{selectedMonth!=="all"&&<span style={{fontSize:12,fontWeight:400,marginLeft:8,opacity:0.7}}>{fmtMonth(selectedMonth)}</span>}</span>
                <span style={{fontWeight:900,fontSize:20}}>NT$ {totalAmt.toLocaleString()}</span>
              </div>
            </_HCard>
            <_HCard>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
                <span style={{fontWeight:700,color:THEME}}>│ 文字整理</span>
                <button onClick={copyReport} style={{border:`1px solid ${THEME}`,background:"#EBF4FB",color:"#1565C0",padding:"5px 14px",borderRadius:8,cursor:"pointer",fontSize:13}}>複製</button>
              </div>
              <div style={{background:"#F8F9FA",borderRadius:8,padding:"14px 16px",fontFamily:"monospace",fontSize:13,lineHeight:2}}>
                {Object.entries(report).map(([dept,reasons])=>{
                  const t=Object.values(reasons).reduce((s,v)=>s+v,0);
                  return <div key={dept} style={{marginBottom:8}}><div>〈{dept}〉NT$ {t.toLocaleString()}</div>{Object.entries(reasons).map(([r,a])=><div key={r} style={{color:"#555"}}>－{r} → NT$ {a.toLocaleString()}</div>)}</div>;
                })}
                <div style={{borderTop:"1px solid #E0E0E0",paddingTop:8,fontWeight:700}}>總金額 NT$ {totalAmt.toLocaleString()}</div>
              </div>
            </_HCard>
          </>}
        </div>}
      </div>
      {toast&&<div style={{position:"fixed",bottom:24,right:24,background:toast.type==="success"?"#1a1a2e":"#F44336",color:"#fff",padding:"14px 20px",borderRadius:12,boxShadow:"0 4px 20px rgba(0,0,0,0.2)",fontSize:14,fontWeight:500,zIndex:999}}>{toast.msg}</div>}
    </div>
  );
}


// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [loginError, setLoginError] = useState("");
  const [loginName, setLoginName] = useState("");
  const [loginPw, setLoginPw] = useState("");
  const [loginPwShow, setLoginPwShow] = useState(false);
  const [page, setPage] = useState("dashboard");
  const [theme, setTheme] = useState("invoice");

  const [invoiceProjects, setInvoiceProjects] = useState([]);
  const [invoiceVendors,  setInvoiceVendors]  = useState([]);
  const [hsinchuProjects, setHsinchuProjects] = useState([]);
  const [hsinchuVendors,  setHsinchuVendors]  = useState([]);
  const [users, setUsers] = useState([]);
  const [dbLoading, setDbLoading] = useState(true);
  const [dbError,   setDbError]   = useState("");
  const [toast,      setToast]      = useState(null);
  const [myPwModal,  setMyPwModal]  = useState(false);
  const [myPw,       setMyPw]       = useState("");
  const [myPwShow,   setMyPwShow]   = useState(false);
  const [prefillData, setPrefillData] = useState(null);

  const projects    = theme === "invoice" ? invoiceProjects : hsinchuProjects;
  const vendors     = theme === "invoice" ? invoiceVendors  : hsinchuVendors;
  const setProjects = theme === "invoice" ? setInvoiceProjects : setHsinchuProjects;
  const setVendors  = theme === "invoice" ? setInvoiceVendors  : setHsinchuVendors;

  useEffect(() => {
    const loadAll = async () => {
      setDbLoading(true); setDbError("");
      try {
        const [usersRaw, vendorsRaw, projectsRaw] = await Promise.all([
          sb.from("users").select("*").order("id",{ascending:true}).get(),
          sb.from("vendors").select("*").order("id",{ascending:true}).get(),
          sb.from("projects").select("*").order("id",{ascending:true}).get(),
        ]);
        setUsers((usersRaw||[]).map(dbToUser));
        const allVendors = (vendorsRaw||[]).map(dbToVendor);
        setInvoiceVendors(allVendors.filter(v => v.theme==="invoice"));
        setHsinchuVendors(allVendors.filter(v => v.theme==="hsinchu"));
        const allProjects = (projectsRaw||[]).map(dbToProject);
        setInvoiceProjects(allProjects.filter(p => p.theme==="invoice"));
        setHsinchuProjects(allProjects.filter(p => p.theme==="hsinchu"));
      } catch(err) {
        setDbError("⚠️ 無法連線至資料庫：" + err.message);
      } finally {
        setDbLoading(false);
      }
    };
    loadAll();
  }, []);

  const handleLogin = () => {
    const found = users.find(u => u.name === loginName && u.active);
    if (!found) { setLoginError("找不到此使用者或帳號已停用"); return; }
    if (found.password && found.password !== loginPw) { setLoginError("密碼錯誤，請重新輸入"); return; }
    setCurrentUser(found);
    setLoginError(""); setLoginName(""); setLoginPw("");
    setPage("dashboard");
  };
  const handleLogout = () => { setCurrentUser(null); setPage("dashboard"); };

  const effectiveStatus = (p) => {
    if (p.status === "已開立" && p.expectedPayDate) {
      const today = new Date(); today.setHours(0,0,0,0);
      if (new Date(p.expectedPayDate) < today) return "到期未付款";
    }
    return p.status;
  };
  const projectsWithStatus = projects.map(p => ({...p, _effectiveStatus: effectiveStatus(p)}));

  const showToast = (msg, type="success") => { setToast({msg, type}); setTimeout(() => setToast(null), 3000); };

  const handleSubmit = async (formData) => {
    try {
      let vendorId = +formData.vendorId || null;
      if (formData.vSelected && formData.vFields?.name) {
        const updatedVendorData = {
          name: formData.vFields.name||formData.vSelected.name,
          tax_id: formData.vFields.taxId||formData.vSelected.taxId,
          bank: formData.vFields.bank||formData.vSelected.bank,
          branch: formData.vFields.branch||formData.vSelected.branch,
          account_name: formData.vFields.accountName||formData.vSelected.accountName,
          account: formData.vFields.account||formData.vSelected.account,
        };
        await sb.from("vendors").eq("id", formData.vSelected.id).update(updatedVendorData);
        setVendors(prev => prev.map(v => v.id===formData.vSelected.id ? {...v,...dbToVendor({...v,...updatedVendorData,id:v.id})} : v));
        vendorId = formData.vSelected.id;
      } else if (!formData.vSelected && formData.vFields?.name) {
        const newVDb = {...vendorToDb({...formData.vFields,taxId:formData.vFields.taxId}), theme};
        const [created] = await sb.from("vendors").insert(newVDb);
        const newVendor = dbToVendor(created);
        setVendors(prev => [...prev, newVendor]);
        vendorId = newVendor.id;
      }
      const newProjectDb = {
        ...projectToDb({
          applicant: currentUser.name, company: formData.company,
          title: formData.items.map(i=>i.label).filter(Boolean).join("、")||"未命名專案",
          items: formData.items,
          clientTitle: formData.clientTitle||"", taxId: formData.taxId||"",
          amount: Math.round(formData.total/1.05), taxAmount: formData.total,
          status: "申請中",
          invoiceDate: formData.invoiceDate||new Date().toISOString().split("T")[0],
          expectedPayDate: formData.expectedPayDate||"",
          orderNo: formData.orderNo||"", vendorId, paid: 0,
        }),
        theme,
      };
      const [created] = await sb.from("projects").insert(newProjectDb);
      setProjects(prev => [dbToProject(created), ...prev]);
      showToast(formData.vSelected && formData.vFields?.name ? "✅ 發票申請已提交，廠商資料已同步更新！" : "✅ 發票申請已提交！");
      setPage("projects");
    } catch(err) { showToast("❌ 提交失敗：" + err.message, "error"); }
  };

  const handleExport = async () => {
    try {
      if (!window.XLSX) {
        await new Promise((res,rej) => {
          const s=document.createElement("script");
          s.src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
          s.onload=res; s.onerror=()=>rej(new Error("無法載入 SheetJS"));
          document.head.appendChild(s);
        });
      }
      const rows = projectsWithStatus.map(p => ({
        "申請人":p.applicant,"公司":p.company,"專案名稱":p.title,
        "委刊單編號":p.orderNo||"","公司抬頭":p.clientTitle||"",
        "統一編號":p.taxId||"","未稅金額":p.amount,"含稅金額":p.taxAmount,
        "發票日期":p.invoiceDate||"","預計收款日":p.expectedPayDate||"",
        "發票號碼":p.invoiceNo||"",
        "廠商名稱":vendors.find(v=>v.id===p.vendorId)?.name||"",
        "廠商帳號":vendors.find(v=>v.id===p.vendorId)?.account||"",
        "狀態":p._effectiveStatus,"已收款金額":p.paidAmount||"",
        "入帳日":p.paidDate||"","手續費":p.bankFee||"",
        "佣金":p.commission||"","說明":p.commissionNo||"",
        "支出申請單編號":p.expenseNo||"","傳票號碼":p.voucherNo||"",
      }));
      const ws=window.XLSX.utils.json_to_sheet(rows);
      const wb=window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(wb,ws,"進度");
      const date=new Date().toLocaleDateString("zh-TW").replace(/\//g,"-");
      window.XLSX.writeFile(wb,`進度_${date}.xlsx`);
      showToast("✅ 已匯出 Excel");
    } catch(err) { showToast("❌ 匯出失敗："+err.message,"error"); }
  };

  const setProjectsWithDb = (updaterOrArray) => {
    if (typeof updaterOrArray === "function") {
      setProjects(prev => {
        const next = updaterOrArray(prev);
        next.forEach(p => {
          const {_effectiveStatus,...clean} = p;
          sb.from("projects").eq("id",clean.id).update(projectToDb(clean)).catch(e=>console.warn("DB sync failed",e.message));
        });
        return next;
      });
    } else {
      setProjects(updaterOrArray);
      updaterOrArray.forEach(p => {
        const {_effectiveStatus,...clean} = p;
        sb.from("projects").eq("id",clean.id).update(projectToDb(clean)).catch(e=>console.warn("DB sync failed",e.message));
      });
    }
  };

  if (dbError) {
    return (
      <div style={{minHeight:"100vh",background:"#F5F6FA",display:"flex",alignItems:"center",justifyContent:"center",padding:24,fontFamily:"'Noto Sans TC','Microsoft JhengHei',sans-serif"}}>
        <div style={{background:"#fff",borderRadius:16,padding:32,maxWidth:500,width:"100%",boxShadow:"0 4px 24px rgba(0,0,0,0.08)"}}>
          <div style={{fontSize:40,textAlign:"center",marginBottom:16}}>🔌</div>
          <div style={{fontSize:18,fontWeight:700,textAlign:"center",marginBottom:12,color:"#c33"}}>資料庫連線失敗</div>
          <pre style={{background:"#FFF0F0",border:"1px solid #FCC",borderRadius:8,padding:16,fontSize:13,color:"#c33",whiteSpace:"pre-wrap",marginBottom:20}}>{dbError}</pre>
          <button onClick={()=>window.location.reload()} style={{width:"100%",padding:12,background:"#F2A0A0",color:"#fff",border:"none",borderRadius:10,fontSize:15,fontWeight:700,cursor:"pointer"}}>🔄 重新連線</button>
        </div>
      </div>
    );
  }

  if (dbLoading) {
    return (
      <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#FDE8E8 0%,#FFF5F5 100%)",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Noto Sans TC','Microsoft JhengHei',sans-serif"}}>
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:36,fontWeight:900,color:"#F2A0A0",marginBottom:12}}>媽咪拜</div>
          <div style={{fontSize:14,color:"#aaa"}}>正在連線資料庫⋯</div>
          <div style={{marginTop:16,width:40,height:40,border:"4px solid #F2A0A033",borderTop:"4px solid #F2A0A0",borderRadius:"50%",animation:"spin 0.8s linear infinite",margin:"16px auto 0"}}/>
          <style>{"@keyframes spin{to{transform:rotate(360deg);}}"}</style>
        </div>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#FDE8E8 0%,#FFF5F5 100%)",display:"flex",alignItems:"center",justifyContent:"center",padding:"16px",fontFamily:"'Noto Sans TC','Microsoft JhengHei',sans-serif"}}>
        <style>{"* { box-sizing: border-box; } body { margin: 0; } input, select { font-size: 16px; }"}</style>
        <div style={{background:"#fff",borderRadius:20,padding:"36px 28px",width:"100%",maxWidth:400,boxShadow:"0 16px 60px rgba(242,160,160,0.25)"}}>
          <div style={{textAlign:"center",marginBottom:32}}>
            <div style={{fontSize:32,fontWeight:900,color:"#F2A0A0",letterSpacing:-1,marginBottom:6}}>媽咪拜</div>
            <div style={{fontSize:13,color:"#aaa"}}>發票開立申請與收款追蹤系統</div>
          </div>
          <div style={{marginBottom:16}}>
            <label style={lbl}>使用者名稱</label>
            <select style={{...inp,fontSize:14}} value={loginName} onChange={e=>{setLoginName(e.target.value);setLoginError("");}}>
              <option value="">— 請選擇 —</option>
              {users.filter(u=>u.active).map(u=>(<option key={u.id} value={u.name}>{u.name}</option>))}
            </select>
          </div>
          <div style={{marginBottom:24}}>
            <label style={lbl}>密碼</label>
            <div style={{position:"relative"}}>
              <input style={{...inp,paddingRight:44,fontSize:14}} type={loginPwShow?"text":"password"} value={loginPw}
                onChange={e=>{setLoginPw(e.target.value);setLoginError("");}}
                onKeyDown={e=>e.key==="Enter"&&handleLogin()}
                placeholder={loginName&&!users.find(u=>u.name===loginName)?.password?"此帳號尚未設定密碼，可直接登入":"請輸入密碼"}
              />
              <span onClick={()=>setLoginPwShow(v=>!v)} style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",cursor:"pointer",fontSize:16,color:"#aaa",userSelect:"none"}}>{loginPwShow?"🙈":"👁"}</span>
            </div>
            {loginName&&!users.find(u=>u.name===loginName)?.password&&(
              <div style={{fontSize:11,color:"#aaa",marginTop:4}}>此帳號尚未設定密碼，點登入即可進入</div>
            )}
          </div>
          {loginError&&(<div style={{background:"#FFF0F0",border:"1px solid #FCC",borderRadius:8,padding:"10px 14px",fontSize:13,color:"#c33",marginBottom:16}}>⚠ {loginError}</div>)}
          <button onClick={handleLogin} disabled={!loginName}
            style={{width:"100%",padding:14,background:loginName?"#F2A0A0":"#eee",color:loginName?"#fff":"#aaa",border:"none",borderRadius:12,fontSize:16,fontWeight:700,cursor:loginName?"pointer":"not-allowed",transition:"background 0.2s"}}>
            登入
          </button>
        </div>
      </div>
    );
  }

  const canAccess = (p) => {
    const perms = {dashboard:true,"invoice-form":true,projects:true,vendors:true,
      csv: currentUser.role==="finance"||currentUser.role==="admin",
      users: currentUser.role==="admin"};
    return perms[p]!==false;
  };
  const navItems = [
    {id:"dashboard",label:"儀表板",icon:"📊"},{id:"invoice-form",label:"申請發票",icon:"📄"},
    {id:"projects",label:"進度",icon:"🧾"},{id:"vendors",label:"廠商管理",icon:"🏢"},
    ...(canAccess("csv")?[{id:"csv",label:"銀行對賬",icon:"🏦"}]:[]),
    ...(canAccess("users")?[{id:"users",label:"帳號管理",icon:"👥"}]:[]),
  ];

  return (
    <div style={{minHeight:"100vh",background:"#F5F6FA",fontFamily:"'Noto Sans TC','Microsoft JhengHei',sans-serif"}}>
      <GlobalStyles />
      {myPwModal&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:"#fff",borderRadius:16,padding:24,width:380,boxShadow:"0 8px 40px rgba(0,0,0,0.2)"}}>
            <div style={{fontSize:26,textAlign:"center",marginBottom:10}}>🔑</div>
            <div style={{fontWeight:700,fontSize:16,textAlign:"center",marginBottom:6}}>變更我的密碼</div>
            <div style={{fontSize:13,color:"#888",textAlign:"center",marginBottom:20}}>{currentUser.name}</div>
            <div style={{marginBottom:20}}>
              <label style={lbl}>新密碼</label>
              <div style={{position:"relative"}}>
                <input style={{...inp,paddingRight:44}} type={myPwShow?"text":"password"} value={myPw} onChange={e=>setMyPw(e.target.value)} placeholder="請輸入新密碼" autoFocus/>
                <span onClick={()=>setMyPwShow(v=>!v)} style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",cursor:"pointer",fontSize:16,color:"#aaa",userSelect:"none"}}>{myPwShow?"🙈":"👁"}</span>
              </div>
            </div>
            <div style={{display:"flex",gap:10}}>
              <button disabled={!myPw.trim()}
                onClick={async()=>{try{await sb.from("users").eq("id",currentUser.id).update({password:myPw.trim()});setUsers(prev=>prev.map(u=>u.id===currentUser.id?{...u,password:myPw.trim()}:u));setCurrentUser(u=>({...u,password:myPw.trim()}));setMyPwModal(false);setMyPw("");setMyPwShow(false);showToast("✅ 密碼已更新！");}catch(err){showToast("❌ 更新失敗："+err.message,"error");}}}
                style={{flex:1,border:"none",background:myPw.trim()?"#F2A0A0":"#ccc",color:"#fff",padding:"11px",borderRadius:8,cursor:myPw.trim()?"pointer":"not-allowed",fontWeight:700}}>確認更新</button>
              <button onClick={()=>{setMyPwModal(false);setMyPw("");setMyPwShow(false);}} style={{flex:1,border:"1px solid #ddd",background:"#fff",padding:"11px",borderRadius:8,cursor:"pointer"}}>取消</button>
            </div>
          </div>
        </div>
      )}
      <div style={{background:THEMES[theme].color,color:"#fff",padding:"0 16px",display:"flex",alignItems:"center",justifyContent:"space-between",height:56,position:"sticky",top:0,zIndex:100}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:18,fontWeight:900,letterSpacing:-0.5,color:"#fff"}}>媽咪拜</span>
          <div style={{display:"flex",gap:4,marginLeft:8}}>
            {Object.values(THEMES).map(t=>(
              <button key={t.id} onClick={()=>{setTheme(t.id);setPage("dashboard");}}
                style={{padding:"3px 10px",borderRadius:20,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,background:theme===t.id?"#fff":"#ffffff33",color:theme===t.id?THEMES[theme].color:"#fff",transition:"all 0.2s"}}>
                {t.icon} {t.label}
              </button>
            ))}
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:12,color:"#fff"}}>{currentUser.name}</span>
          <span style={{background:"#ffffff33",color:"#fff",padding:"2px 8px",borderRadius:20,fontSize:11}}>{ROLES[currentUser.role]}</span>
          <button className="navbar-pw-btn" onClick={()=>{setMyPwModal(true);setMyPw("");setMyPwShow(false);}} style={{background:"#ffffff33",color:"#fff",border:"none",borderRadius:8,padding:"4px 10px",fontSize:12,cursor:"pointer"}}>🔑 密碼</button>
          <button onClick={handleLogout} style={{background:"#ffffff33",color:"#fff",border:"none",borderRadius:8,padding:"4px 10px",fontSize:12,cursor:"pointer"}}>登出</button>
        </div>
      </div>
      <div style={{display:"flex",minHeight:"calc(100vh - 56px)"}}>
        <div className="sidebar-desktop" style={{flexDirection:"column",width:200,background:"#fff",borderRight:"1px solid #F0F0F0",padding:"16px 0",flexShrink:0}}>
          <div style={{padding:"8px 20px 12px",fontSize:11,color:"#aaa",fontWeight:600,borderBottom:"1px solid #F5F5F5",marginBottom:4}}>
            {THEMES[theme].icon} {THEMES[theme].label}
          </div>
          {navItems.map(item=>(
            <button key={item.id} onClick={()=>setPage(item.id)}
              style={{width:"100%",display:"flex",alignItems:"center",gap:10,padding:"12px 20px",border:"none",background:page===item.id?`${THEMES[theme].color}18`:"transparent",color:page===item.id?THEMES[theme].color:"#666",fontWeight:page===item.id?700:400,cursor:"pointer",fontSize:14,borderLeft:page===item.id?`3px solid ${THEMES[theme].color}`:"3px solid transparent"}}>
              <span>{item.icon}</span>{item.label}
            </button>
          ))}
        </div>
        <div className="content-area" style={{flex:1,padding:24,overflowY:"auto"}}>
          {theme==="hsinchu"?(
            <HsinchuApp currentUser={currentUser}/>
          ):(
            <>
              {page==="dashboard"    &&<Dashboard projects={projectsWithStatus} currentUser={currentUser}/>}
              {page==="invoice-form" &&<InvoiceForm currentUser={currentUser} vendors={vendors} onSubmit={handleSubmit} prefillData={prefillData} onClearPrefill={()=>setPrefillData(null)}/>}
              {page==="projects"     &&<ProjectList projects={projectsWithStatus} setProjects={setProjectsWithDb} vendors={vendors} currentUser={currentUser} users={users} onExport={handleExport} onCopyProject={p=>{ setPrefillData(p); setPage("invoice-form"); }}/>}
              {page==="vendors"      &&<VendorManagement vendors={vendors} setVendors={setVendors} canEdit={currentUser.role==="finance"||currentUser.role==="admin"}/>}
              {page==="csv"          &&canAccess("csv")&&<CSVReconcile projects={projectsWithStatus} vendors={vendors} setProjects={setProjectsWithDb}/>}
              {page==="users"        &&canAccess("users")&&<UserManagement users={users} setUsers={setUsers}/>}
            </>
          )}
        </div>
      </div>
      <div className="sidebar-mobile" style={{position:"fixed",bottom:0,left:0,right:0,zIndex:99,background:"#fff",borderTop:"1px solid #F0F0F0",flexDirection:"column",paddingBottom:"env(safe-area-inset-bottom,0px)",boxShadow:"0 -2px 12px rgba(0,0,0,0.06)"}}>
        <div style={{display:"flex",borderBottom:"1px solid #F0F0F0"}}>
          {Object.values(THEMES).map(t=>(
            <button key={t.id} onClick={()=>{setTheme(t.id);setPage("dashboard");}}
              style={{flex:1,padding:"6px 4px",border:"none",cursor:"pointer",fontSize:11,fontWeight:600,background:theme===t.id?`${t.color}18`:"#fff",color:theme===t.id?t.color:"#aaa",borderBottom:theme===t.id?`2px solid ${t.color}`:"2px solid transparent"}}>
              {t.icon} {t.label}
            </button>
          ))}
        </div>
        <div style={{display:"flex",justifyContent:"space-around",alignItems:"center",padding:"6px 0"}}>
          {navItems.map(item=>(
            <button key={item.id} onClick={()=>setPage(item.id)}
              style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:2,border:"none",background:"none",cursor:"pointer",padding:"4px 2px",color:page===item.id?THEMES[theme].color:"#aaa"}}>
              <span style={{fontSize:18}}>{item.icon}</span>
              <span style={{fontSize:10,fontWeight:page===item.id?700:400}}>{item.label}</span>
            </button>
          ))}
          <button onClick={()=>{setMyPwModal(true);setMyPw("");setMyPwShow(false);}}
            style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:2,border:"none",background:"none",cursor:"pointer",padding:"4px 2px",color:"#aaa"}}>
            <span style={{fontSize:18}}>🔑</span>
            <span style={{fontSize:10}}>密碼</span>
          </button>
        </div>
      </div>
      {toast&&(<div style={{position:"fixed",bottom:24,right:24,background:toast.type==="success"?"#1a1a2e":"#FF9800",color:"#fff",padding:"14px 20px",borderRadius:12,boxShadow:"0 4px 20px rgba(0,0,0,0.2)",fontSize:14,fontWeight:500,zIndex:999}}>{toast.msg}</div>)}
    </div>
  );
}

// ─── SHARED STYLES ────────────────────────────────────────────────────────────
const sectionTitle = { fontSize: 14, fontWeight: 700, color: "#888", textTransform: "uppercase", letterSpacing: 1, marginBottom: 16, marginTop: 0 };
const grid2 = { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 };
const field = { display: "flex", flexDirection: "column" };
const lbl = { fontSize: 12, fontWeight: 600, color: "#666", marginBottom: 6 };
const inp = { border: "1.5px solid #E8E8E8", borderRadius: 8, padding: "9px 12px", fontSize: 14, color: "#1a1a2e", outline: "none", width: "100%", boxSizing: "border-box", background: "#fff" };
const td = { padding: "12px 14px", color: "#444", verticalAlign: "middle" };
const stickyTh = { padding: "12px 14px", textAlign: "left", color: "#555", fontWeight: 600, borderBottom: "2px solid #F0F0F0", whiteSpace: "nowrap", background: "#F8F9FA", position: "sticky", top: 0, zIndex: 2 };
const stickyTd = { padding: "12px 14px", color: "#444", verticalAlign: "middle", position: "sticky", zIndex: 1, boxShadow: "inset -1px 0 0 #F0F0F0" };

// ── 民國日期 → 西元日期（7碼 YYYMMDD 或 YYY/MM/DD）────────────────────────
const parseRocDate = (val) => {
  if (!val) return val;
  const s = String(val).replace(/\//g, "").replace(/-/g, "").trim();
  // 7碼純數字：1150302 → 2026-03-02
  if (/^\d{7}$/.test(s)) {
    const y = parseInt(s.slice(0, 3), 10) + 1911;
    const m = s.slice(3, 5);
    const d = s.slice(5, 7);
    return `${y}-${m}-${d}`;
  }
  // 已是西元格式（含 -）直接回傳
  if (/^\d{4}-\d{2}-\d{2}$/.test(val)) return val;
  return val;
};
