// ─────────────────────────────────────────────────────────────────────────────
//  SAKER BAPTIST COLLEGE — Main App (Supabase backend)
// ─────────────────────────────────────────────────────────────────────────────
import { useState, useEffect, useRef, useCallback } from "react";
import { supabase } from "./supabase.js";
import * as XLSX from "xlsx";

// ─── Palette ──────────────────────────────────────────────────────────────────
const C = {
  navy:"#0D2340", navyMid:"#163558", gold:"#C9962A", goldLight:"#F0C050",
  white:"#FFFFFF", red:"#B91C1C", green:"#15803D",
  gray:"#6B7280", grayLight:"#E5E7EB", grayBg:"#F3F4F6", border:"#D1D5DB",
};

// ─── Constants ────────────────────────────────────────────────────────────────
const FORMS = ["Form 1","Form 2","Form 3","Form 4","Form 5"];
const SUBJECTS_BY_FORM = {
  "Form 1":["English language","French/ Français","Mathematics","Health science","Home management","Citizenship","Food and Nutrition","Chemistry","History","Economics","Geography","Biology","Human Biology","Physics","Computer Studies","Religious Studies","Hygiene","Sport/Physical education","Manual Labour"],
  "Form 2":["English language","French/ Français","Mathematics","Health science","Home management","Citizenship","Food and Nutrition","Chemistry","History","Economics","Geography","Biology","Human Biology","Physics","Computer Studies","Religious Studies","Hygiene","Sport/Physical education","Manual Labour"],
  "Form 3":["English language","French/ Français","Mathematics","Health science","Home management","Citizenship","Food and Nutrition","Chemistry","History","Economics","Geography","Biology","Human Biology","Physics","Computer Studies","Religious Studies","Hygiene","Sport/Physical education","Manual Labour"],
  "Form 4":["English language","French/ Français","Mathematics","Chemistry","History","Economics","Geography","Biology","Physics","Literature in English","Computer Studies","Religious Studies","Commerce","Accounts","Hygiene","Sport/Physical education"],
  "Form 5":["English language","French/ Français","Mathematics","Chemistry","History","Economics","Geography","Biology","Physics","Literature in English","Computer Studies","Religious Studies","Commerce","Accounts","Hygiene","Sport/Physical education"],
};
const SEQ_LABELS  = ["SEQ 1","SEQ 2","SEQ 3","SEQ 4","SEQ 5","SEQ 6"];
const TERM_SEQS   = {"First Term":["SEQ 1","SEQ 2"],"Second Term":["SEQ 3","SEQ 4"],"Third Term":["SEQ 5","SEQ 6"]};
const TERMS       = ["First Term","Second Term","Third Term"];
const ACAD_YEARS  = ["2026/2027","2027/2028","2028/2029"];
const TOTAL_FEE   = 35000;
const REG_NORMAL  = 500;
const REG_LATE    = 1000;
const LATE_CUTOFF = "2026-09-30";
const MOMO_MERCHANT_NUMBER = "PASTE_YOUR_MOMO_NUMBER_HERE"; // e.g. "677123456" — the school's MTN Mobile Money receiving number

const DEFAULT_COEFF = {
  "English language":4,"French/ Français":4,"Mathematics":4,
  "Biology":3,"Human Biology":3,"Chemistry":3,"Physics":3,
  "History":2,"Geography":2,"Computer Studies":2,
  "Religious Studies":1,"Sport/Physical education":1,"Hygiene":1,
  "Economics":3,"Commerce":3,"Accounts":3,"Literature in English":3,
  "Health science":2,"Home management":2,"Citizenship":1,
  "Food and Nutrition":2,"Manual Labour":1,
};
const getCoeff = sub => DEFAULT_COEFF[sub] || 2;

// ─── Helpers ──────────────────────────────────────────────────────────────────
const todayStr = () => new Date().toISOString().slice(0,10);
const fmtDate  = d => { if(!d) return "—"; const [y,m,day]=d.split("-"); return `${day}/${m}/${y}`; };
const isLate   = d => d && d > LATE_CUTOFF;

function scoreToGrade(v) {
  if (v===null||v===undefined||v==="") return { grade:"—", remark:"—" };
  const pct = Number(v)*5;
  if (pct>=85) return { grade:"A",  remark:"Excellent"   };
  if (pct>=75) return { grade:"B+", remark:"Very Good"   };
  if (pct>=65) return { grade:"B",  remark:"Good"        };
  if (pct>=55) return { grade:"C+", remark:"Fairly Good" };
  if (pct>=50) return { grade:"C",  remark:"Average"     };
  if (pct>=45) return { grade:"D",  remark:"Pass"        };
  return              { grade:"F",  remark:"Fail"        };
}
const gradeCol = g =>
  ["A","B+","B"].includes(g) ? C.green : g==="F" ? C.red : g==="—" ? C.gray : C.gold;

// Compress photo using canvas before storing in DB
function compressPhoto(base64) {
  return new Promise((resolve) => {
    if (!base64 || !base64.startsWith("data:")) { resolve(null); return; }
    const img = new Image();
    img.onload = () => {
      const MAX = 500; // passport photos print sharp at this size; still small enough for DB storage
      let w = img.width, h = img.height;
      const scale = Math.min(MAX/w, MAX/h, 1);
      w = Math.round(w*scale); h = Math.round(h*scale);
      const canvas = document.createElement("canvas");
      canvas.width = w; canvas.height = h;
      canvas.getContext("2d").drawImage(img, 0, 0, w, h);
      resolve(canvas.toDataURL("image/jpeg", 0.85));
    };
    img.onerror = () => resolve(null);
    img.src = base64;
  });
}
async function uploadPhoto(studentId, base64) {
  return await compressPhoto(base64);
}

// Print helper - works on mobile and desktop
function domPrint(id, html, size="A4 portrait", margin="8mm") {
  const fullHtml = `<!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
      *{box-sizing:border-box;margin:0;padding:0}
      @page{size:${size};margin:${margin}}
      @media print{ html,body{ width:${size.includes("A5")?"148mm":"210mm"}; } }
      body{font-family:'Segoe UI',Arial,sans-serif;background:#fff;
           -webkit-print-color-adjust:exact;print-color-adjust:exact}
    </style></head><body>${html}</body></html>`;

  // Method 1: Blob URL (best for mobile)
  try {
    const blob = new Blob([fullHtml], {type:"text/html"});
    const url  = URL.createObjectURL(blob);
    const win  = window.open(url, "_blank");
    if (win) {
      win.onload = () => { win.print(); URL.revokeObjectURL(url); };
      setTimeout(() => { try { win.print(); URL.revokeObjectURL(url); } catch(e){} }, 1000);
      return;
    }
    URL.revokeObjectURL(url);
  } catch(e) { console.log("Blob print failed:", e); }

  // Method 2: DOM inject fallback
  document.getElementById(id)?.remove();
  document.getElementById(id+"-s")?.remove();
  const st = document.createElement("style");
  st.id = id+"-s";
  st.textContent = `@media print{body>*:not(#${id}){display:none!important}#${id}{display:block!important;position:fixed;inset:0;background:#fff;z-index:99999}@page{size:${size};margin:${margin}}}#${id}{display:none}`;
  document.head.appendChild(st);
  const div = document.createElement("div");
  div.id = id; div.innerHTML = html;
  document.body.appendChild(div);
  setTimeout(()=>{
    window.print();
    setTimeout(()=>{ document.getElementById(id)?.remove(); document.getElementById(id+"-s")?.remove(); }, 2500);
  }, 300);
}

// Export an array of row-objects to a downloadable .xlsx file.
// rows: array of plain objects (keys become column headers)
// filename: e.g. "SBC-Students-2026-2027.xlsx"
// sheetName: tab name inside the workbook (max 31 chars, Excel limit)
function exportToExcel(rows, filename, sheetName="Sheet1") {
  if (!rows || !rows.length) { alert("Nothing to export."); return; }
  try {
    const ws = XLSX.utils.json_to_sheet(rows);
    // Auto-size columns roughly based on content length
    const colWidths = Object.keys(rows[0]).map(key => {
      const maxLen = Math.max(key.length, ...rows.map(r => String(r[key]??"").length));
      return { wch: Math.min(Math.max(maxLen+2, 8), 40) };
    });
    ws["!cols"] = colWidths;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0,31));
    XLSX.writeFile(wb, filename);
  } catch(e) {
    console.error("Excel export error:", e);
    alert("Export failed: " + e.message);
  }
}

// Export one workbook with multiple sheets — one sheet per group.
// sheets: array of { name: string, rows: array of row-objects }
// Sheet names are Excel-sanitized and de-duplicated (Excel forbids
// \ / ? * [ ] and caps names at 31 characters).
function exportToExcelMultiSheet(sheets, filename) {
  const usable = (sheets||[]).filter(s => s.rows && s.rows.length);
  if (!usable.length) { alert("Nothing to export."); return; }
  try {
    const wb = XLSX.utils.book_new();
    const usedNames = new Set();
    usable.forEach(({ name, rows }) => {
      let safeName = String(name||"Sheet").replace(/[\\/?*[\]:]/g,"").slice(0,31) || "Sheet";
      let finalName = safeName;
      let n = 2;
      while (usedNames.has(finalName)) { finalName = safeName.slice(0,28)+"_"+n; n++; }
      usedNames.add(finalName);

      const ws = XLSX.utils.json_to_sheet(rows);
      const colWidths = Object.keys(rows[0]).map(key => {
        const maxLen = Math.max(key.length, ...rows.map(r => String(r[key]??"").length));
        return { wch: Math.min(Math.max(maxLen+2, 8), 40) };
      });
      ws["!cols"] = colWidths;
      XLSX.utils.book_append_sheet(wb, ws, finalName);
    });
    XLSX.writeFile(wb, filename);
  } catch(e) {
    console.error("Excel export error:", e);
    alert("Export failed: " + e.message);
  }
}

// Print a downloadable payment slip for an approved fee payment submission.
// sub: the payment_submissions row (amount, momo_number, submitted_at, reviewed_at, reviewed_by, id)
// student: the student record (name, id, form)
function printPaymentSlip(sub, student) {
  if (!sub || !student) return;
  const html = `<div style="font-family:Segoe UI,sans-serif;padding:14px;color:#0D2340">
    <div style="max-width:360px;margin:0 auto;border:2px solid #0D2340;border-radius:10px;overflow:hidden">
      <div style="background:#0D2340;color:#fff;padding:14px;text-align:center">
        <div style="font-size:26px">🎓</div>
        <h2 style="font-size:14px;font-weight:900;margin:4px 0">SAKER BAPTIST COLLEGE</h2>
        <p style="font-size:9px;opacity:.65;margin:0">NGEPTANG · NONI · NW REGION</p>
        <div style="background:#15803D;display:inline-block;padding:2px 12px;border-radius:10px;font-size:10px;font-weight:800;margin-top:7px">FEE PAYMENT SLIP</div>
      </div>
      <div style="padding:14px">
        <div style="font-family:monospace;font-weight:900;font-size:13px;text-align:center;margin-bottom:9px;padding-bottom:7px;border-bottom:1px dashed #e5e7eb">PMT-${String(sub.id).padStart(6,"0")}</div>
        ${[
          ["Student Name", student.name],
          ["Matricule", student.id],
          ["Form", student.form],
          ["Amount Paid", `${sub.amount.toLocaleString()} FCFA`],
          ["Payment Method", "MTN Mobile Money"],
          ["Phone Used", sub.momo_number||"—"],
          ["Submitted", fmtDate(sub.submitted_at?.slice(0,10))],
          ["Confirmed By", sub.reviewed_by||"School Office"],
          ["Confirmed On", fmtDate(sub.reviewed_at?.slice(0,10))],
        ].map(([l,v])=>`<div style="display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #f0f0f0;font-size:11.5px"><span style="color:#6B7280">${l}</span><span style="font-weight:600">${v||"—"}</span></div>`).join("")}
        <div style="background:#f0fdf4;border-radius:8px;padding:12px;margin-top:11px;text-align:center">
          <div style="font-size:10px;color:#6b7280;margin-bottom:2px">AMOUNT CONFIRMED</div>
          <div style="font-size:26px;font-weight:900;color:#15803D">${sub.amount.toLocaleString()} FCFA</div>
          <div style="font-size:10px;color:#6b7280;margin-top:2px">✓ Approved Payment</div>
        </div>
      </div>
      <div style="text-align:center;font-size:9px;color:#9ca3af;border-top:1px dashed #e5e7eb;padding:8px 14px">This slip confirms a payment toward school fees for Academic Year 2026/2027.<br>Keep for your records.</div>
    </div>
  </div>`;
  domPrint("sbc-payment-slip", html, "A5 portrait", "7mm");
}

// ─── Nav ──────────────────────────────────────────────────────────────────────
const navItems = [
  { key:"dashboard",    label:"Dashboard",    icon:"📊", roles:["admin","teacher"] },
  { key:"registration", label:"Registration", icon:"📋", roles:["admin"] },
  { key:"students",     label:"Students",     icon:"🎒", roles:["admin","teacher"] },
  { key:"teachers",     label:"Teachers",     icon:"👩‍🏫", roles:["admin"] },
  { key:"marks",        label:"Enter Marks",  icon:"✏️", roles:["admin","teacher"] },
  { key:"attendance",   label:"Attendance",   icon:"🗓️", roles:["admin","teacher"] },
  { key:"reports",      label:"Report Cards", icon:"📄", roles:["admin","teacher"] },
  { key:"fees",         label:"School Fees",  icon:"💰", roles:["admin"] },
  { key:"notices",      label:"Notices",      icon:"📢", roles:["admin","teacher"] },
  { key:"calendar",     label:"Calendar",     icon:"📅", roles:["admin","teacher"] },
  { key:"profile",      label:"My Profile",   icon:"👤", roles:["admin","teacher"] },
];

// ─── Shared UI ────────────────────────────────────────────────────────────────
const lbl = { display:"block", fontSize:11, fontWeight:600, color:C.gray, marginBottom:3 };
const inp = { width:"100%", padding:"9px 11px", border:`1px solid ${C.border}`, borderRadius:7, fontSize:13, color:C.navy, background:C.white, boxSizing:"border-box", outline:"none", fontFamily:"inherit" };

const Modal = ({ title, onClose, children }) => (
  <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"flex-end",justifyContent:"center",zIndex:400}}>
    <div style={{background:C.white,borderRadius:"14px 14px 0 0",width:"100%",maxWidth:520,maxHeight:"90vh",overflow:"auto",boxShadow:"0 -8px 32px rgba(0,0,0,0.25)"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"13px 16px",borderBottom:`1px solid ${C.grayLight}`,position:"sticky",top:0,background:C.white,zIndex:1}}>
        <h3 style={{margin:0,fontSize:14,fontWeight:800,color:C.navy}}>{title}</h3>
        <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:22,color:C.gray}}>×</button>
      </div>
      <div style={{padding:16}}>{children}</div>
    </div>
  </div>
);
const Card  = ({ title, children, style }) => (
  <div style={{background:C.white,borderRadius:11,padding:13,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",...style}}>
    {title && <h3 style={{margin:"0 0 10px",fontSize:13,fontWeight:800,color:C.navy}}>{title}</h3>}
    {children}
  </div>
);
const Btn   = ({ children, onClick, outline, color, disabled }) => (
  <button onClick={onClick} disabled={disabled} style={{padding:"8px 16px",background:outline?"transparent":(color||C.navy),color:outline?(color||C.navy):C.white,border:`2px solid ${color||C.navy}`,borderRadius:8,fontWeight:700,fontSize:13,cursor:disabled?"not-allowed":"pointer",whiteSpace:"nowrap",opacity:disabled?0.5:1}}>{children}</button>
);
const SmBtn = ({ children, onClick, color, style }) => (
  <button onClick={onClick} style={{padding:"4px 10px",background:color+"18",color,border:`1px solid ${color}30`,borderRadius:6,fontWeight:700,fontSize:12,cursor:"pointer",whiteSpace:"nowrap",...style}}>{children}</button>
);
const Pill  = ({ children, color }) => <span style={{padding:"2px 8px",background:color+"20",color,borderRadius:10,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{children}</span>;
const Info  = ({ label, value }) => (
  <div><div style={{fontSize:9.5,color:C.gray}}>{label}</div><div style={{fontSize:11.5,fontWeight:600,color:C.navy}}>{value||"—"}</div></div>
);
const Fr    = ({ label, children }) => (
  <div style={{marginBottom:9}}><label style={lbl}>{label}</label>{children}</div>
);
const Empty = ({ text }) => <div style={{textAlign:"center",padding:"36px 0",color:C.gray,fontSize:13}}>{text}</div>;
const Spin  = ({ text }) => (
  <div style={{textAlign:"center",padding:"48px 0",color:C.gray}}>
    <div style={{fontSize:32,marginBottom:10}}>⏳</div>
    <div style={{fontSize:14}}>{text||"Loading…"}</div>
  </div>
);
const PhotoBox = ({ photo, size=[64,78] }) => photo
  ? <img src={photo} alt="passport" style={{width:size[0],height:size[1],objectFit:"cover",border:`2px solid ${C.navy}`,borderRadius:3,display:"block"}}/>
  : <div style={{width:size[0],height:size[1],border:`2px dashed ${C.border}`,borderRadius:3,background:C.grayBg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",color:C.gray}}>
      <span style={{fontSize:size[0]*0.3}}>👤</span>
      <span style={{fontSize:7,marginTop:2,textAlign:"center",lineHeight:1.3}}>PHOTO<br/>PASSEPORT</span>
    </div>;

// ══════════════════════════════════════════════════════════════════════════════
//  MAIN APP
// ══════════════════════════════════════════════════════════════════════════════
export default function App() {
  const [portalMode, setPortalMode] = useState(false); // true = showing the public Parent Portal instead of staff login
  const [session,    setSession]    = useState(undefined); // undefined=loading
  const [userRole,   setUserRole]   = useState(null);
  const [userProfile,setUserProfile]= useState(null);
  const [page,       setPage]       = useState("dashboard");
  const [menuOpen,   setMenuOpen]   = useState(false);
  const [recoveryMode, setRecoveryMode] = useState(false); // true when user arrived via password-reset email link

  // Data state
  const [students,   setStudents]   = useState([]);
  const [teachers,   setTeachers]   = useState([]);
  const [marksMap,   setMarksMap]   = useState({}); // key → mark record
  const [feesMap,    setFeesMap]    = useState({}); // studentId → {paid,total}
  const [attendanceMap, setAttendanceMap] = useState({}); // "studentId-date" → {status}
  const [calendarEvents, setCalendarEvents] = useState([]);
  const [notices,    setNotices]    = useState([]);
  const [loading,    setLoading]    = useState(true);
  const [dbError,    setDbError]    = useState(null);

  // ── Auth ───────────────────────────────────────────────────────────────────
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));
    const { data: { subscription } } = supabase.auth.onAuthStateChange((event, session) => {
      setSession(session);
      // Supabase fires this exact event when the user clicks the reset-password link in their email
      if (event === "PASSWORD_RECOVERY") setRecoveryMode(true);
    });
    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) { setUserRole(null); setUserProfile(null); return; }
    (async () => {
      const { data } = await supabase.from("users").select("*").eq("id", session.user.id).single();
      setUserRole(data?.role || "teacher");
      setUserProfile(data);
    })();
  }, [session]);

  // ── Load all data ──────────────────────────────────────────────────────────
  useEffect(() => {
    if (!session) return;
    loadAll();
  }, [session]);

  async function loadAll() {
    setLoading(true);
    try {
      const [
        { data: st }, { data: tc }, { data: mk },
        { data: fe }, { data: nt }, { data: at }, { data: cal }
      ] = await Promise.all([
        supabase.from("students").select("*").order("form").order("name"),
        supabase.from("teachers").select("*").order("name"),
        supabase.from("marks").select("*"),
        supabase.from("fees").select("*"),
        supabase.from("notices").select("*").order("posted_date", { ascending: false }),
        supabase.from("attendance").select("*"),
        supabase.from("calendar_events").select("*").order("event_date"),
      ]);
      setStudents(st || []);
      setTeachers(tc || []);
      // Build marks map: "studentId-subject-seq-year" → record
      const mm = {};
      (mk||[]).forEach(m => { mm[`${m.student_id}-${m.subject}-${m.seq}-${m.acad_year}`] = m; });
      setMarksMap(mm);
      // Build fees map: studentId → {paid, total}
      const fm = {};
      (fe||[]).forEach(f => { fm[f.student_id] = { paid: f.paid||0, total: f.total||TOTAL_FEE }; });
      setFeesMap(fm);
      // Build attendance map: "studentId-date" → record {status: present|absent|late}
      const am = {};
      (at||[]).forEach(a => { am[`${a.student_id}-${a.date}`] = a; });
      setAttendanceMap(am);
      setNotices(nt || []);
      setCalendarEvents(cal || []);
      setDbError(null);
    } catch(e) {
      setDbError(e.message);
    }
    setLoading(false);
  }

  // ── Real-time subscriptions ─────────────────────────────────────────────────
  useEffect(() => {
    if (!session) return;
    const sub = supabase.channel("db-changes")
      .on("postgres_changes", { event:"*", schema:"public", table:"students" }, () => loadAll())
      .on("postgres_changes", { event:"*", schema:"public", table:"marks"    }, () => loadAll())
      .on("postgres_changes", { event:"*", schema:"public", table:"fees"     }, () => loadAll())
      .on("postgres_changes", { event:"*", schema:"public", table:"notices"  }, () => loadAll())
      .on("postgres_changes", { event:"*", schema:"public", table:"attendance" }, () => loadAll())
      .on("postgres_changes", { event:"*", schema:"public", table:"calendar_events" }, () => loadAll())
      .on("postgres_changes", { event:"*", schema:"public", table:"teachers" }, () => loadAll())
      .subscribe();
    return () => supabase.removeChannel(sub);
  }, [session]);

  // ── Auth actions ───────────────────────────────────────────────────────────
  async function doLogin(email, password) {
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) throw error;
  }
  async function doLogout() {
    await supabase.auth.signOut();
    setMenuOpen(false); setPage("dashboard");
  }
  async function doRequestPasswordReset(email) {
    const { error } = await supabase.auth.resetPasswordForEmail(email, {
      redirectTo: window.location.origin, // Supabase appends its own recovery tokens to this URL
    });
    if (error) throw error;
  }
  async function doSetNewPassword(newPassword) {
    const { error } = await supabase.auth.updateUser({ password: newPassword });
    if (error) throw error;
    setRecoveryMode(false);
  }

  // ── Auto-logout after inactivity — protects unattended/shared devices ──
  // Two layers: (1) this timer catches an open tab left unattended,
  // (2) persistSession:false in supabase.js catches the tab/app being
  // closed entirely, so nothing lingers logged in for the next person
  // who opens the site on a shared phone.
  useEffect(() => {
    if (!session) return;
    let timer;
    const INACTIVITY_LIMIT = 10*60*1000; // 10 minutes
    const reset = () => {
      clearTimeout(timer);
      timer = setTimeout(() => {
        doLogout();
        alert("You were signed out after 10 minutes of inactivity for security.");
      }, INACTIVITY_LIMIT);
    };
    const events = ["mousedown","keydown","touchstart","scroll"];
    events.forEach(ev => window.addEventListener(ev, reset));
    // Also reset when the tab becomes visible again (e.g. switching back
    // from another app) — and immediately check if we've already been
    // away long enough that we should log out right now rather than wait.
    let hiddenAt = null;
    const onVisibility = () => {
      if (document.hidden) {
        hiddenAt = Date.now();
      } else if (hiddenAt && Date.now()-hiddenAt >= INACTIVITY_LIMIT) {
        doLogout();
        alert("You were signed out after being away for security.");
      } else {
        reset();
      }
    };
    document.addEventListener("visibilitychange", onVisibility);
    reset();
    return () => {
      clearTimeout(timer);
      events.forEach(ev => window.removeEventListener(ev, reset));
      document.removeEventListener("visibilitychange", onVisibility);
    };
  }, [session]);

  // ── DB write helpers (passed via ctx) ──────────────────────────────────────
  async function saveStudent(s) {
    // photo_url is already compressed base64 or null — save directly
    let photoUrl = s.photo_url || null;
    // Auto-generate a PIN for students who don't have one yet (e.g. registered before this feature existed)
    const existing = students.find(x=>x.id===s.id);
    const parentPin = s.parent_pin || existing?.parent_pin || String(Math.floor(1000 + Math.random()*9000));
    const row = {
      id: s.id, name: s.name, form: s.form, gender: s.gender,
      dob: s.dob||null, parent: s.parent, phone: s.phone, address: s.address,
      photo_url: photoUrl||null, active: s.active!==false,
      reg_status: s.reg_status||"pending", reg_date: s.reg_date||null,
      reg_fee: s.reg_fee||null, reg_receipt: s.reg_receipt||null,
      reg_paid_by: s.reg_paid_by||null, reg_cashier: s.reg_cashier||null,
      is_late_reg: s.is_late_reg||false, graduated: s.graduated||false,
      parent_pin: parentPin,
    };
    const { error } = await supabase.from("students").upsert(row, { onConflict:"id" });
    if (error) {
      console.error("saveStudent error:", error);
      alert("Save error: " + error.message + (error.details?(" — "+error.details):"") );
      throw error;
    }
    // Auto-create fee record if not exists
    await supabase.from("fees").upsert({ student_id: s.id, paid: feesMap[s.id]?.paid||0, total: TOTAL_FEE }, { onConflict:"student_id", ignoreDuplicates:true });
    await loadAll();
  }

  async function deleteStudent(id) {
    await supabase.from("students").update({ active: false }).eq("id", id);
    await loadAll();
  }

  async function saveTeacher(t) {
    const row = {
      id: t.id, name: t.name, email: t.email,
      subjects: t.subjects||[], forms: t.forms||[],
      active: t.active!==false, joined: t.joined||todayStr(),
    };
    const { error } = await supabase.from("teachers").upsert(row, { onConflict:"id" });
    if (error) throw error;
    await loadAll();
  }

  async function saveMark(studentId, subject, seq, year, score, coeff, teacherId) {
    const { error } = await supabase.from("marks").upsert({
      student_id: studentId, subject, seq, acad_year: year,
      score: Number(score), coeff: Number(coeff)||1,
      teacher_id: teacherId||null,
    }, { onConflict:"student_id,subject,seq,acad_year" });
    if (error) throw error;
    await loadAll(); // refresh marks immediately
  }

  async function saveFee(studentId, paid) {
    const { error } = await supabase.from("fees").upsert(
      { student_id: studentId, paid, total: TOTAL_FEE },
      { onConflict:"student_id" }
    );
    if (error) throw error;
    await loadAll();
  }

  // records: array of { studentId, status } where status is "present"|"absent"|"late"
  async function saveAttendanceBulk(date, records, markedBy) {
    const rows = records.map(r => ({
      student_id: r.studentId, date, status: r.status, marked_by: markedBy||null,
    }));
    const { error } = await supabase.from("attendance").upsert(rows, { onConflict:"student_id,date" });
    if (error) throw error;
    await loadAll();
  }

  async function saveNotice(n) {
    const row = { title: n.title, body: n.body, author: n.author, posted_date: n.posted_date||todayStr() };
    if (n.id) { await supabase.from("notices").update(row).eq("id", n.id); }
    else       { await supabase.from("notices").insert(row); }
    await loadAll();
  }

  async function deleteNotice(id) {
    await supabase.from("notices").delete().eq("id", id);
    await loadAll();
  }

  async function saveCalendarEvent(ev) {
    const row = {
      title: ev.title, description: ev.description||null,
      event_date: ev.event_date, end_date: ev.end_date||null,
      category: ev.category||"other", forms: ev.forms||[],
    };
    if (ev.id) { await supabase.from("calendar_events").update(row).eq("id", ev.id); }
    else       { await supabase.from("calendar_events").insert(row); }
    await loadAll();
  }

  async function deleteCalendarEvent(id) {
    await supabase.from("calendar_events").delete().eq("id", id);
    await loadAll();
  }

  // ── Current user info ───────────────────────────────────────────────────────
  const currentTeacher = teachers.find(t => t.email === session?.user?.email);
  const currentUser = {
    id:    session?.user?.id || "ADMIN",
    name:  currentTeacher?.name || userProfile?.name || session?.user?.email || "Admin",
    email: session?.user?.email || "",
  };

  // ── Loading / not-logged-in screens ────────────────────────────────────────
  // Parent Portal doesn't need staff auth at all — short-circuit everything else
  if (portalMode) return <ParentPortal onBack={()=>setPortalMode(false)}/>;

  if (session === undefined) return (
    <div style={{minHeight:"100vh",background:C.navy,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:14}}>
      <div style={{fontSize:48}}>🎓</div>
      <div style={{color:C.white,fontWeight:800,fontSize:18}}>Saker Baptist College</div>
      <div style={{color:C.goldLight,fontSize:13}}>Starting up…</div>
    </div>
  );
  // Password recovery gives a temporary session — intercept it before normal routing
  if (recoveryMode) return <SetNewPasswordScreen onSetPassword={doSetNewPassword} onCancel={doLogout}/>;
  if (!session) return <LoginScreen onLogin={doLogin} onRequestReset={doRequestPasswordReset} onOpenPortal={()=>setPortalMode(true)}/>;

  const ctx = {
    auth: { user: currentUser, role: userRole||"teacher" },
    students, teachers, marksMap, feesMap, notices, attendanceMap, calendarEvents,
    saveStudent, deleteStudent, saveTeacher, saveMark,
    saveFee, saveNotice, deleteNotice, loadAll, saveAttendanceBulk,
    saveCalendarEvent, deleteCalendarEvent,
  };

  return (
    <div style={{fontFamily:"'Segoe UI',system-ui,sans-serif",background:C.grayBg,minHeight:"100vh"}}>
      {/* ── Top bar ─────────────────────────────────────────────────────────── */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:C.navy,padding:"11px 14px",position:"sticky",top:0,zIndex:200}}>
        <div style={{display:"flex",alignItems:"center",gap:9}}>
          <span style={{fontSize:20}}>🎓</span>
          <div>
            <div style={{color:C.white,fontWeight:800,fontSize:13,lineHeight:1.2}}>Saker Baptist College</div>
            <div style={{color:C.goldLight,fontSize:9,letterSpacing:1}}>NGEPTANG · NONI</div>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{color:"rgba(255,255,255,0.6)",fontSize:11,textAlign:"right"}}>
            <div style={{color:C.white,fontWeight:600}}>{currentUser.name.split(" ").slice(-1)[0]}</div>
            <div style={{textTransform:"capitalize",fontSize:10}}>{ctx.auth.role}</div>
          </div>
          <button onClick={()=>setMenuOpen(o=>!o)} style={{background:"rgba(255,255,255,0.12)",border:"none",borderRadius:8,color:C.white,fontSize:20,width:38,height:38,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>☰</button>
        </div>
      </div>

      {/* ── Slide-out nav ───────────────────────────────────────────────────── */}
      {menuOpen && (
        <div style={{position:"fixed",inset:0,zIndex:300,display:"flex"}}>
          <div style={{background:C.navy,width:230,height:"100%",display:"flex",flexDirection:"column",overflowY:"auto"}}>
            <div style={{padding:"18px 16px 8px",borderBottom:"1px solid rgba(255,255,255,0.08)"}}>
              <div style={{color:C.goldLight,fontSize:10,marginBottom:3}}>Signed in as</div>
              <div style={{color:C.white,fontWeight:700,fontSize:13}}>{currentUser.name}</div>
              <div style={{color:"rgba(255,255,255,0.5)",fontSize:10,textTransform:"capitalize"}}>{ctx.auth.role}</div>
            </div>
            <nav style={{flex:1,paddingTop:6}}>
              {navItems.filter(n=>n.roles.includes(ctx.auth.role)).map(n => (
                <button key={n.key} onClick={()=>{setPage(n.key);setMenuOpen(false);}} style={{display:"flex",alignItems:"center",gap:11,width:"100%",padding:"12px 18px",border:"none",cursor:"pointer",background:page===n.key?"rgba(201,150,42,0.15)":"transparent",borderLeft:page===n.key?`3px solid ${C.gold}`:"3px solid transparent",color:page===n.key?C.goldLight:"rgba(255,255,255,0.7)",fontWeight:page===n.key?700:400,fontSize:14}}>
                  <span>{n.icon}</span>{n.label}
                </button>
              ))}
            </nav>
            <div style={{padding:14,borderTop:"1px solid rgba(255,255,255,0.08)"}}>
              <button onClick={doLogout} style={{width:"100%",padding:"9px",background:"rgba(255,255,255,0.08)",border:"1px solid rgba(255,255,255,0.15)",borderRadius:8,color:"rgba(255,255,255,0.8)",fontSize:13,cursor:"pointer",fontWeight:600}}>Sign Out</button>
            </div>
          </div>
          <div style={{flex:1,background:"rgba(0,0,0,0.5)"}} onClick={()=>setMenuOpen(false)}/>
        </div>
      )}

      {/* ── Page content ────────────────────────────────────────────────────── */}
      <div style={{padding:"14px 12px",maxWidth:900,margin:"0 auto"}}>
        <div style={{marginBottom:13}}>
          <h1 style={{margin:0,fontSize:17,fontWeight:800,color:C.navy}}>{navItems.find(n=>n.key===page)?.label||""}</h1>
        </div>

        {dbError && (
          <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:10,padding:"12px 16px",marginBottom:14,color:C.red}}>
            <strong>Database error:</strong> {dbError}
            <button onClick={loadAll} style={{marginLeft:12,fontSize:12,padding:"3px 10px",background:C.red,color:C.white,border:"none",borderRadius:6,cursor:"pointer"}}>Retry</button>
          </div>
        )}

        {loading ? <Spin text="Loading from Supabase…" /> : <>
          {page==="dashboard"    && <DashboardPage    ctx={ctx} setPage={setPage}/>}
          {page==="registration" && ctx.auth.role==="admin" && <RegistrationPage ctx={ctx}/>}
          {page==="students"     && <StudentsPage     ctx={ctx}/>}
          {page==="teachers"     && ctx.auth.role==="admin" && <TeachersPage    ctx={ctx}/>}
          {page==="marks"        && <MarksPage        ctx={ctx}/>}
          {page==="attendance"   && <AttendancePage   ctx={ctx}/>}
          {page==="reports"      && <ReportsPage      ctx={ctx}/>}
          {page==="fees"         && ctx.auth.role==="admin" && <FeesPage        ctx={ctx}/>}
          {page==="notices"      && <NoticesPage      ctx={ctx}/>}
          {page==="calendar"     && <CalendarPage     ctx={ctx}/>}
          {page==="profile"      && <ProfilePage      ctx={ctx}/>}
        </>}
      </div>
    </div>
  );
}

// ─── Login Screen ──────────────────────────────────────────────────────────────
function LoginScreen({ onLogin, onRequestReset, onOpenPortal }) {
  const [mode,  setMode]  = useState("login"); // "login" | "forgot"
  const [email, setEmail] = useState("");
  const [pass,  setPass]  = useState("");
  const [err,   setErr]   = useState("");
  const [info,  setInfo]  = useState("");
  const [busy,  setBusy]  = useState(false);
  const [attempts, setAttempts] = useState(0);
  const [lockedUntil, setLockedUntil] = useState(0);
  const [now, setNow] = useState(Date.now());

  // Tick every second while locked so the countdown updates
  useEffect(() => {
    if (!lockedUntil) return;
    const t = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(t);
  }, [lockedUntil]);

  const isLocked = lockedUntil > now;
  const secondsLeft = isLocked ? Math.ceil((lockedUntil-now)/1000) : 0;

  async function handle() {
    if (isLocked) return;
    if (!email.trim() || !pass) { setErr("Enter your email and password."); return; }
    setErr(""); setBusy(true);
    try {
      await onLogin(email.trim(), pass);
      setAttempts(0);
    } catch(e) {
      const next = attempts+1;
      setAttempts(next);
      if (next>=5) {
        // Lock out for 60 seconds after 5 failed attempts — slows down brute-force guessing
        setLockedUntil(Date.now()+60000);
        setErr("Too many failed attempts. Try again in 60 seconds.");
      } else {
        setErr(e.message.includes("Invalid login") ? `Incorrect email or password. (${next}/5 attempts)` : "Login error: "+e.message);
      }
    }
    setBusy(false);
  }

  async function handleReset() {
    if (!email.trim()) { setErr("Enter your email address first."); return; }
    setErr(""); setInfo(""); setBusy(true);
    try {
      await onRequestReset(email.trim());
      setInfo("Check your email — we sent a link to reset your password. It may take a minute to arrive.");
    } catch(e) {
      setErr("Could not send reset email: " + e.message);
    }
    setBusy(false);
  }

  if (mode === "forgot") {
    return (
      <div style={{minHeight:"100vh",background:`linear-gradient(160deg,${C.navy},${C.navyMid} 55%,#1a4a6e)`,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
        <div style={{width:"100%",maxWidth:380}}>
          <div style={{textAlign:"center",marginBottom:24}}>
            <div style={{width:70,height:70,borderRadius:"50%",background:C.gold,margin:"0 auto 12px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:32,boxShadow:"0 0 0 5px rgba(201,150,42,0.25)"}}>🔑</div>
            <h1 style={{color:C.white,fontSize:20,fontWeight:900,margin:0}}>Reset Password</h1>
            <p style={{color:C.goldLight,fontSize:11,margin:"4px 0 0",letterSpacing:1.5}}>SAKER BAPTIST COLLEGE</p>
          </div>
          <div style={{background:C.white,borderRadius:14,padding:22,boxShadow:"0 20px 60px rgba(0,0,0,0.35)"}}>
            <h2 style={{margin:"0 0 10px",fontSize:16,fontWeight:800,color:C.navy}}>Forgot your password?</h2>
            <p style={{fontSize:12,color:C.gray,margin:"0 0 16px"}}>Enter your email and we'll send you a link to set a new one.</p>
            <Fr label="Email Address">
              <input style={inp} type="email" autoComplete="username" placeholder="your@email.com" value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleReset()} disabled={busy}/>
            </Fr>
            {err  && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"8px 12px",color:C.red,fontSize:12,marginBottom:12}}>{err}</div>}
            {info && <div style={{background:"#f0fdf4",border:"1px solid #86efac",borderRadius:8,padding:"8px 12px",color:C.green,fontSize:12,marginBottom:12}}>{info}</div>}
            <Btn onClick={handleReset} disabled={busy}>{busy?"Sending…":"Send Reset Link →"}</Btn>
            <button onClick={()=>{setMode("login");setErr("");setInfo("");}} style={{width:"100%",marginTop:10,padding:"9px",background:"transparent",border:"none",color:C.navyMid,fontSize:12,fontWeight:700,cursor:"pointer"}}>← Back to Sign In</button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{minHeight:"100vh",background:`linear-gradient(160deg,${C.navy},${C.navyMid} 55%,#1a4a6e)`,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{width:"100%",maxWidth:380}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{width:70,height:70,borderRadius:"50%",background:C.gold,margin:"0 auto 12px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:32,boxShadow:"0 0 0 5px rgba(201,150,42,0.25)"}}>🎓</div>
          <h1 style={{color:C.white,fontSize:20,fontWeight:900,margin:0}}>Saker Baptist College</h1>
          <p style={{color:C.goldLight,fontSize:11,margin:"4px 0 0",letterSpacing:1.5}}>NGEPTANG · NONI</p>
          <p style={{color:"rgba(255,255,255,0.45)",fontSize:10,margin:"5px 0 0"}}>School Administrative Portal · 2026/2027</p>
        </div>
        <div style={{background:C.white,borderRadius:14,padding:22,boxShadow:"0 20px 60px rgba(0,0,0,0.35)"}}>
          <h2 style={{margin:"0 0 16px",fontSize:16,fontWeight:800,color:C.navy}}>Sign In</h2>
          <Fr label="Email Address">
            <input style={inp} type="email" autoComplete="username" placeholder="your@email.com" value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handle()} disabled={isLocked}/>
          </Fr>
          <Fr label="Password">
            <input style={inp} type="password" autoComplete="current-password" placeholder="••••••••" value={pass} onChange={e=>setPass(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handle()} disabled={isLocked}/>
          </Fr>
          {err && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"8px 12px",color:C.red,fontSize:12,marginBottom:12}}>{err}</div>}
          {isLocked
            ? <div style={{width:"100%",padding:"10px",background:C.grayBg,color:C.gray,borderRadius:8,fontWeight:700,fontSize:13,textAlign:"center"}}>🔒 Locked — try again in {secondsLeft}s</div>
            : <Btn onClick={handle} disabled={busy}>{busy?"Signing in…":"Sign In →"}</Btn>
          }
          <button onClick={()=>{setMode("forgot");setErr("");setInfo("");}} style={{width:"100%",marginTop:12,padding:"6px",background:"transparent",border:"none",color:C.navyMid,fontSize:12,fontWeight:700,cursor:"pointer",textDecoration:"underline"}}>Forgot your password?</button>
          <p style={{textAlign:"center",fontSize:10,color:C.gray,marginTop:10,marginBottom:0}}>Contact your administrator if you need a new account.</p>
          <p style={{textAlign:"center",fontSize:9.5,color:C.gray,marginTop:6,marginBottom:0}}>🔒 For security, you'll need to sign in again each time you open this site — this keeps your account safe on shared devices.</p>
        </div>
        <button onClick={onOpenPortal} style={{width:"100%",marginTop:14,padding:"11px",background:"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.2)",borderRadius:10,color:C.white,fontSize:13,fontWeight:700,cursor:"pointer"}}>
          👨‍👩‍👧 Parent? Check Your Child's Records →
        </button>
      </div>
    </div>
  );
}

// ─── Set New Password Screen (reached via email reset link) ────────────────────
function SetNewPasswordScreen({ onSetPassword, onCancel }) {
  const [pass1, setPass1] = useState("");
  const [pass2, setPass2] = useState("");
  const [err,   setErr]   = useState("");
  const [busy,  setBusy]  = useState(false);
  const [done,  setDone]  = useState(false);

  async function handle() {
    setErr("");
    if (pass1.length < 6) { setErr("Password must be at least 6 characters."); return; }
    if (pass1 !== pass2)  { setErr("Passwords do not match."); return; }
    setBusy(true);
    try {
      await onSetPassword(pass1);
      setDone(true);
    } catch(e) { setErr("Could not update password: " + e.message); }
    setBusy(false);
  }

  if (done) {
    return (
      <div style={{minHeight:"100vh",background:`linear-gradient(160deg,${C.navy},${C.navyMid} 55%,#1a4a6e)`,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
        <div style={{width:"100%",maxWidth:380,background:C.white,borderRadius:14,padding:26,boxShadow:"0 20px 60px rgba(0,0,0,0.35)",textAlign:"center"}}>
          <div style={{fontSize:40,marginBottom:10}}>✅</div>
          <h2 style={{margin:"0 0 8px",fontSize:16,fontWeight:800,color:C.navy}}>Password Updated</h2>
          <p style={{fontSize:12,color:C.gray,marginBottom:16}}>You're signed in with your new password.</p>
          <Btn onClick={onCancel}>Continue →</Btn>
        </div>
      </div>
    );
  }

  return (
    <div style={{minHeight:"100vh",background:`linear-gradient(160deg,${C.navy},${C.navyMid} 55%,#1a4a6e)`,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{width:"100%",maxWidth:380}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{width:70,height:70,borderRadius:"50%",background:C.gold,margin:"0 auto 12px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:32,boxShadow:"0 0 0 5px rgba(201,150,42,0.25)"}}>🔑</div>
          <h1 style={{color:C.white,fontSize:20,fontWeight:900,margin:0}}>Set New Password</h1>
          <p style={{color:C.goldLight,fontSize:11,margin:"4px 0 0",letterSpacing:1.5}}>SAKER BAPTIST COLLEGE</p>
        </div>
        <div style={{background:C.white,borderRadius:14,padding:22,boxShadow:"0 20px 60px rgba(0,0,0,0.35)"}}>
          <Fr label="New Password">
            <input style={inp} type="password" autoComplete="new-password" placeholder="At least 6 characters" value={pass1} onChange={e=>setPass1(e.target.value)} disabled={busy}/>
          </Fr>
          <Fr label="Confirm New Password">
            <input style={inp} type="password" autoComplete="new-password" placeholder="Re-type password" value={pass2} onChange={e=>setPass2(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handle()} disabled={busy}/>
          </Fr>
          {err && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"8px 12px",color:C.red,fontSize:12,marginBottom:12}}>{err}</div>}
          <Btn onClick={handle} disabled={busy}>{busy?"Updating…":"Update Password →"}</Btn>
          <button onClick={onCancel} style={{width:"100%",marginTop:10,padding:"9px",background:"transparent",border:"none",color:C.gray,fontSize:12,cursor:"pointer"}}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ─── Parent Portal ───────────────────────────────────────────────────────────
// Public, read-only lookup. No staff login required — parents authenticate with
// Matricule + 4-digit PIN, verified server-side by the `parent_portal_lookup`
// Postgres function (see PARENT_PORTAL.sql). The raw tables stay protected by
// RLS; only this narrow, rate-limitable function is reachable with the PIN.
function ParentPortal({ onBack }) {
  const [matricule, setMatricule] = useState("");
  const [pin, setPin] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const [data, setData] = useState(null); // { student, marks, attendance, fee }
  const [attempts, setAttempts] = useState(0);
  const [lockedUntil, setLockedUntil] = useState(0);
  const [now, setNow] = useState(Date.now());

  useEffect(() => {
    if (!lockedUntil) return;
    const t = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(t);
  }, [lockedUntil]);

  const isLocked = lockedUntil > now;
  const secondsLeft = isLocked ? Math.ceil((lockedUntil-now)/1000) : 0;

  async function lookup() {
    if (isLocked) return;
    if (!matricule.trim() || pin.length!==4) { setErr("Enter your child's Matricule and 4-digit PIN."); return; }
    setErr(""); setBusy(true);
    try {
      const { data: result, error } = await supabase.rpc("parent_portal_lookup", {
        p_matricule: matricule.trim().toUpperCase(),
        p_pin: pin.trim(),
      });
      if (error) throw error;
      if (!result || !result.student) {
        const next = attempts+1;
        setAttempts(next);
        if (next>=5) {
          setLockedUntil(Date.now()+60000);
          setErr("Too many failed attempts. Try again in 60 seconds.");
        } else {
          setErr(`Matricule or PIN not recognized. (${next}/5 attempts)`);
        }
      } else {
        setData(result);
        setAttempts(0);
      }
    } catch(e) {
      setErr("Lookup failed: " + e.message);
    }
    setBusy(false);
  }

  if (data) return <ParentPortalResults data={data} matricule={matricule.trim().toUpperCase()} pin={pin.trim()} onBack={()=>{setData(null);setMatricule("");setPin("");}} onExit={onBack}/>;

  return (
    <div style={{minHeight:"100vh",background:`linear-gradient(160deg,${C.navy},${C.navyMid} 55%,#1a4a6e)`,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{width:"100%",maxWidth:380}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{width:70,height:70,borderRadius:"50%",background:C.gold,margin:"0 auto 12px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:32,boxShadow:"0 0 0 5px rgba(201,150,42,0.25)"}}>👨‍👩‍👧</div>
          <h1 style={{color:C.white,fontSize:20,fontWeight:900,margin:0}}>Parent Portal</h1>
          <p style={{color:C.goldLight,fontSize:11,margin:"4px 0 0",letterSpacing:1.5}}>SAKER BAPTIST COLLEGE</p>
          <p style={{color:"rgba(255,255,255,0.45)",fontSize:10,margin:"5px 0 0"}}>View your child's marks, attendance & fees</p>
        </div>
        <div style={{background:C.white,borderRadius:14,padding:22,boxShadow:"0 20px 60px rgba(0,0,0,0.35)"}}>
          <Fr label="Student Matricule">
            <input style={inp} placeholder="e.g. SBC01001" value={matricule} onChange={e=>setMatricule(e.target.value)} disabled={isLocked} onKeyDown={e=>e.key==="Enter"&&lookup()}/>
          </Fr>
          <Fr label="4-Digit PIN">
            <input style={{...inp,letterSpacing:4,fontWeight:700,textAlign:"center"}} type="tel" maxLength={4} placeholder="••••" value={pin} onChange={e=>setPin(e.target.value.replace(/\D/g,"").slice(0,4))} disabled={isLocked} onKeyDown={e=>e.key==="Enter"&&lookup()}/>
          </Fr>
          {err && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"8px 12px",color:C.red,fontSize:12,marginBottom:12}}>{err}</div>}
          {isLocked
            ? <div style={{width:"100%",padding:"10px",background:C.grayBg,color:C.gray,borderRadius:8,fontWeight:700,fontSize:13,textAlign:"center"}}>🔒 Locked — try again in {secondsLeft}s</div>
            : <Btn onClick={lookup} disabled={busy}>{busy?"Looking up…":"View Records →"}</Btn>
          }
          <p style={{textAlign:"center",fontSize:10,color:C.gray,marginTop:14,marginBottom:0}}>The Matricule and PIN were printed on your child's registration receipt. Contact the school office if you've lost them.</p>
        </div>
        <button onClick={onBack} style={{width:"100%",marginTop:14,padding:"9px",background:"transparent",border:"none",color:"rgba(255,255,255,0.6)",fontSize:12,cursor:"pointer"}}>← Staff Sign In</button>
      </div>
    </div>
  );
}

function ParentPortalResults({ data, matricule, pin, onBack, onExit }) {
  const { student, marks, attendance, fee, recent_submissions } = data;
  const [tab, setTab] = useState("marks");
  const [refreshedData, setRefreshedData] = useState(data);

  // Group marks by subject, term
  const bySubject = {};
  (marks||[]).forEach(m => {
    if (!bySubject[m.subject]) bySubject[m.subject] = [];
    bySubject[m.subject].push(m);
  });

  const curFee = refreshedData.fee || fee;
  const paid = curFee?.paid||0;
  const total = curFee?.total||TOTAL_FEE;
  const balance = total-paid;
  const pct = Math.round(paid/total*100);
  const submissions = refreshedData.recent_submissions || recent_submissions || [];
  const pendingSubmission = submissions.find(s=>s.status==="pending");

  async function refresh() {
    try {
      const { data: result } = await supabase.rpc("parent_portal_lookup", { p_matricule: matricule, p_pin: pin });
      if (result?.student) setRefreshedData(result);
    } catch(e) { /* silent — refresh is best-effort */ }
  }

  return (
    <div style={{minHeight:"100vh",background:C.grayBg}}>
      <div style={{background:C.navy,padding:"14px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <PhotoBox photo={student.photo_url} size={[36,44]}/>
          <div>
            <div style={{color:C.white,fontWeight:800,fontSize:14}}>{student.name}</div>
            <div style={{color:C.goldLight,fontSize:10}}>{student.id} · {student.form}</div>
          </div>
        </div>
        <button onClick={onBack} style={{background:"rgba(255,255,255,0.12)",border:"none",borderRadius:8,color:C.white,fontSize:12,padding:"7px 12px",cursor:"pointer",fontWeight:700}}>← Back</button>
      </div>

      <div style={{padding:"14px 12px",maxWidth:600,margin:"0 auto"}}>
        <div style={{display:"flex",background:C.white,borderRadius:10,padding:3,marginBottom:13,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          {[["marks","📝 Marks"],["attendance","🗓️ Attendance"],["fees","💰 Fees"]].map(([k,l]) => (
            <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"9px 4px",borderRadius:7,border:"none",cursor:"pointer",fontWeight:700,fontSize:12,background:tab===k?C.navy:"transparent",color:tab===k?C.white:C.gray}}>{l}</button>
          ))}
        </div>

        {tab==="marks" && (
          <div style={{background:C.white,borderRadius:10,overflow:"hidden",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            {Object.keys(bySubject).length===0
              ? <Empty text="No marks recorded yet."/>
              : Object.entries(bySubject).map(([sub, entries], i) => (
                  <div key={sub} style={{padding:"10px 12px",borderBottom:`1px solid ${C.grayBg}`,background:i%2===0?C.white:C.grayBg}}>
                    <div style={{fontWeight:700,color:C.navy,fontSize:13,marginBottom:6}}>{sub}</div>
                    <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                      {entries.sort((a,b)=>a.seq.localeCompare(b.seq)).map(m=>{
                        const {grade}=scoreToGrade(m.score);
                        return <span key={m.seq} style={{fontSize:11,background:gradeCol(grade)+"18",color:gradeCol(grade),padding:"3px 8px",borderRadius:6,fontWeight:700}}>{m.seq.replace("SEQ ","S")}: {m.score}/20</span>;
                      })}
                    </div>
                  </div>
                ))
            }
          </div>
        )}

        {tab==="attendance" && (
          <div style={{background:C.white,borderRadius:10,padding:16,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",textAlign:"center"}}>
            {(()=>{
              let present=0,absent=0,late=0;
              (attendance||[]).forEach(a=>{ if(a.status==="present")present++; else if(a.status==="absent")absent++; else if(a.status==="late")late++; });
              return(
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
                  <div><div style={{fontSize:24,fontWeight:900,color:C.green}}>{present}</div><div style={{fontSize:11,color:C.gray}}>Present</div></div>
                  <div><div style={{fontSize:24,fontWeight:900,color:C.red}}>{absent}</div><div style={{fontSize:11,color:C.gray}}>Absent</div></div>
                  <div><div style={{fontSize:24,fontWeight:900,color:C.gold}}>{late}</div><div style={{fontSize:11,color:C.gray}}>Late</div></div>
                </div>
              );
            })()}
          </div>
        )}

        {tab==="fees" && (
          <div>
            <div style={{background:C.white,borderRadius:10,padding:16,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",marginBottom:11}}>
              <div style={{background:C.grayBg,borderRadius:8,height:10,overflow:"hidden",marginBottom:10}}>
                <div style={{width:`${pct}%`,height:"100%",background:pct>=100?C.green:pct>0?C.gold:C.red,borderRadius:8}}/>
              </div>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginBottom:4}}>
                <span style={{color:C.gray}}>Paid: <strong style={{color:C.navy}}>{paid.toLocaleString()} FCFA</strong></span>
                <span style={{color:C.gray}}>Total: <strong style={{color:C.navy}}>{total.toLocaleString()} FCFA</strong></span>
              </div>
              <div style={{textAlign:"center",marginTop:12,padding:"10px",background:balance>0?"#fef2f2":"#f0fdf4",borderRadius:8}}>
                <div style={{fontSize:11,color:C.gray}}>{balance>0?"Balance Owing":"Fully Paid"}</div>
                <div style={{fontSize:22,fontWeight:900,color:balance>0?C.red:C.green}}>{balance.toLocaleString()} FCFA</div>
              </div>
            </div>

            {balance>0 && (
              pendingSubmission
                ? <div style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:10,padding:14,textAlign:"center"}}>
                    <div style={{fontSize:24,marginBottom:6}}>⏳</div>
                    <div style={{fontWeight:800,color:"#92400e",fontSize:13,marginBottom:2}}>Payment Under Review</div>
                    <div style={{fontSize:11,color:"#92400e"}}>{pendingSubmission.amount.toLocaleString()} FCFA submitted on {fmtDate(pendingSubmission.submitted_at?.slice(0,10))}. The school office will confirm it shortly.</div>
                    <button onClick={refresh} style={{marginTop:10,padding:"6px 14px",background:"transparent",border:"1px solid #fde68a",borderRadius:6,color:"#92400e",fontSize:11,fontWeight:700,cursor:"pointer"}}>🔄 Check Status</button>
                  </div>
                : <PayFeesPanel matricule={matricule} pin={pin} balance={balance} onSubmitted={refresh}/>
            )}

            {submissions.filter(s=>s.status!=="pending").length>0 && (
              <div style={{marginTop:11}}>
                <div style={{fontSize:11,fontWeight:700,color:C.gray,marginBottom:6}}>Recent Payment History</div>
                {submissions.filter(s=>s.status!=="pending").map(s=>(
                  <div key={s.id} style={{background:C.white,borderRadius:8,padding:"9px 12px",marginBottom:6,display:"flex",justifyContent:"space-between",alignItems:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
                    <div>
                      <div style={{fontSize:12,fontWeight:700,color:C.navy}}>{s.amount.toLocaleString()} FCFA</div>
                      <div style={{fontSize:10,color:C.gray}}>{fmtDate(s.submitted_at?.slice(0,10))}</div>
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <Pill color={s.status==="approved"?C.green:C.red}>{s.status==="approved"?"✓ Approved":"✕ Rejected"}</Pill>
                      {s.status==="approved" && <SmBtn onClick={()=>printPaymentSlip(s, student)} color={C.navyMid}>📄 Slip</SmBtn>}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        <button onClick={onExit} style={{width:"100%",marginTop:16,padding:"10px",background:"transparent",border:"none",color:C.gray,fontSize:12,cursor:"pointer"}}>← Exit Parent Portal</button>
      </div>
    </div>
  );
}

// ─── Pay Fees Panel (MTN MoMo USSD + screenshot upload) ─────────────────────────
function PayFeesPanel({ matricule, pin, balance, onSubmitted }) {
  const [step, setStep] = useState("amount"); // "amount" | "instructions" | "upload" | "done"
  const [amount, setAmount] = useState(String(balance));
  const [momoNumber, setMomoNumber] = useState("");
  const [screenshot, setScreenshot] = useState(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const fileInputRef = useRef(null);

  const ussdCode = `*126*16*${MOMO_MERCHANT_NUMBER}*${amount}#`;

  function handleShot(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (ev) => {
      const compressed = await compressPhoto(ev.target.result);
      setScreenshot(compressed);
    };
    reader.readAsDataURL(file);
  }

  async function submit() {
    if (!screenshot) { setErr("Please attach a screenshot of your payment confirmation."); return; }
    setErr(""); setBusy(true);
    try {
      const { data: result, error } = await supabase.rpc("parent_submit_payment", {
        p_matricule: matricule, p_pin: pin,
        p_amount: Number(amount), p_momo_number: momoNumber||null,
        p_screenshot: screenshot,
      });
      if (error) throw error;
      if (!result?.success) throw new Error(result?.error || "Submission failed");
      setStep("done");
      onSubmitted?.();
    } catch(e) { setErr("Could not submit: " + e.message); }
    setBusy(false);
  }

  if (step==="done") {
    return (
      <div style={{background:"#f0fdf4",border:"1px solid #86efac",borderRadius:10,padding:16,textAlign:"center"}}>
        <div style={{fontSize:28,marginBottom:6}}>✅</div>
        <div style={{fontWeight:800,color:C.green,fontSize:14,marginBottom:4}}>Payment Submitted</div>
        <div style={{fontSize:12,color:C.gray}}>The school office will review your payment and confirm it. Check back here for the status.</div>
      </div>
    );
  }

  return (
    <div style={{background:C.white,borderRadius:10,padding:14,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
      {step==="amount" && (
        <div>
          <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:10}}>💰 Pay School Fees</div>
          <Fr label="Amount to Pay (FCFA)">
            <input style={inp} type="number" value={amount} onChange={e=>setAmount(e.target.value)} max={balance} min={1}/>
          </Fr>
          <div style={{fontSize:11,color:C.gray,marginBottom:10}}>Balance owing: {balance.toLocaleString()} FCFA. You can pay in full or in part.</div>
          <Btn onClick={()=>setStep("instructions")} disabled={!amount||Number(amount)<=0}>Continue →</Btn>
        </div>
      )}

      {step==="instructions" && (
        <div>
          <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:10}}>📞 Dial This on Your Phone</div>
          <div style={{background:"#1a4a2e",borderRadius:10,padding:16,textAlign:"center",marginBottom:12}}>
            <div style={{fontFamily:"monospace",fontWeight:900,fontSize:20,color:"#fff",letterSpacing:1}}>{ussdCode}</div>
          </div>
          <ol style={{fontSize:12,color:C.gray,paddingLeft:18,marginBottom:14,lineHeight:1.8}}>
            <li>Dial the code above on the phone paying via MTN Mobile Money</li>
            <li>Confirm the payment of <strong>{Number(amount).toLocaleString()} FCFA</strong> to Saker Baptist College</li>
            <li>Take a screenshot of the confirmation message</li>
            <li>Come back here and upload it</li>
          </ol>
          <div style={{display:"flex",gap:8}}>
            <Btn onClick={()=>setStep("amount")} outline>← Back</Btn>
            <Btn onClick={()=>setStep("upload")}>I've Paid — Upload Proof →</Btn>
          </div>
        </div>
      )}

      {step==="upload" && (
        <div>
          <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:10}}>📤 Upload Payment Proof</div>
          <Fr label="Phone Number Used (optional, helps us verify faster)">
            <input style={inp} type="tel" placeholder="e.g. 677123456" value={momoNumber} onChange={e=>setMomoNumber(e.target.value)}/>
          </Fr>
          <input ref={fileInputRef} type="file" accept="image/*" onChange={handleShot} style={{display:"none"}}/>
          <button onClick={()=>fileInputRef.current?.click()} style={{width:"100%",padding:"12px",background:C.grayBg,border:`2px dashed ${C.border}`,borderRadius:10,cursor:"pointer",fontSize:12,fontWeight:700,color:C.navyMid,marginBottom:10}}>
            {screenshot ? "✓ Screenshot attached — tap to change" : "📸 Tap to attach screenshot"}
          </button>
          {screenshot && <img src={screenshot} alt="preview" style={{width:"100%",maxHeight:180,objectFit:"contain",borderRadius:8,marginBottom:10,border:`1px solid ${C.grayLight}`}}/>}
          {err && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"8px 12px",color:C.red,fontSize:12,marginBottom:10}}>{err}</div>}
          <div style={{display:"flex",gap:8}}>
            <Btn onClick={()=>setStep("instructions")} outline disabled={busy}>← Back</Btn>
            <Btn onClick={submit} disabled={busy||!screenshot}>{busy?"Submitting…":"Submit for Review →"}</Btn>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Dashboard ─────────────────────────────────────────────────────────────────
function DashboardPage({ ctx, setPage }) {
  const { students, feesMap, notices, auth } = ctx;
  const active     = students.filter(s => s.active);
  const regCount   = active.filter(s => s.reg_status==="registered").length;
  const pendCount  = active.filter(s => s.reg_status==="pending").length;
  const collected  = active.reduce((a,s) => a+(feesMap[s.id]?.paid||0), 0);

  const cards = [
    { label:"Registered",      value:regCount,                   icon:"✅", color:C.green,    page:"students" },
    { label:"Pending Reg.",     value:pendCount,                  icon:"⏳", color:C.red,      page:"registration" },
    ...(auth.role==="admin" ? [{ label:"Fees Collected", value:`${Math.round(collected/1000)}k F`, icon:"💰", color:C.gold, page:"fees" }] : []),
  ];

  return (
    <div>
      <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,borderRadius:12,padding:"18px",marginBottom:16,color:C.white}}>
        <h2 style={{margin:0,fontSize:17,fontWeight:800}}>Welcome, {auth.user.name.split(" ")[0]}! 👋</h2>
        <p style={{margin:"4px 0 0",color:"rgba(255,255,255,0.6)",fontSize:12}}>Academic Year 2026/2027 · Saker Baptist College</p>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
        {cards.map(c => (
          <div key={c.label} onClick={()=>setPage(c.page)} style={{background:C.white,borderRadius:11,padding:13,borderLeft:`4px solid ${c.color}`,cursor:"pointer",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{fontSize:20,marginBottom:3}}>{c.icon}</div>
            <div style={{fontSize:22,fontWeight:900,color:c.color}}>{c.value}</div>
            <div style={{fontSize:11,color:C.gray,marginTop:1}}>{c.label}</div>
          </div>
        ))}
      </div>
      <Card title="Enrollment by Form">
        {FORMS.map(f => {
          const n = active.filter(s=>s.form===f).length;
          const max = Math.max(...FORMS.map(fm=>active.filter(s=>s.form===fm).length),1);
          return (
            <div key={f} style={{display:"flex",alignItems:"center",gap:8,marginBottom:7}}>
              <div style={{width:48,fontSize:11,color:C.gray,flexShrink:0}}>{f}</div>
              <div style={{flex:1,background:C.grayBg,borderRadius:4,height:18,overflow:"hidden"}}>
                <div style={{width:`${(n/max)*100}%`,background:C.navy,height:"100%",borderRadius:4,display:"flex",alignItems:"center",paddingLeft:6,transition:"width .5s"}}>
                  <span style={{color:C.white,fontSize:10,fontWeight:700}}>{n}</span>
                </div>
              </div>
            </div>
          );
        })}
      </Card>
      <div style={{marginTop:11}}>
        <Card title="Recent Notices">
          {notices.slice(0,3).map(n => (
            <div key={n.id} style={{borderLeft:`3px solid ${C.gold}`,paddingLeft:10,marginBottom:11}}>
              <div style={{fontWeight:700,color:C.navy,fontSize:12}}>{n.title}</div>
              <div style={{fontSize:11,color:C.gray,marginTop:1}}>{(n.body||"").slice(0,80)}</div>
              <div style={{fontSize:10,color:C.gold,marginTop:2}}>{fmtDate(n.posted_date)}</div>
            </div>
          ))}
          {!notices.length && <Empty text="No notices yet."/>}
        </Card>
      </div>
    </div>
  );
}

// ─── Registration ──────────────────────────────────────────────────────────────
function RegistrationPage({ ctx }) {
  const { students, saveStudent, auth } = ctx;
  const [tab,     setTab]    = useState("register");
  const [receipt, setReceipt]= useState(null);
  const [saving,  setSaving] = useState(false);
  const blank = { name:"", form:"Form 1", gender:"Female", dob:"", parent:"", phone:"", address:"", paidBy:"", isLate:false, photo_url:null };
  const [form, setForm] = useState(blank);
  const photoInputRef = useRef(null);

  const fee = form.isLate ? REG_LATE : REG_NORMAL;

  async function handleRegister() {
    if (!form.name.trim()||!form.parent.trim()||!form.phone.trim()) return;
    setSaving(true);
    const fNum = form.form.replace("Form ","");
    const n    = students.filter(s=>s.form===form.form).length + 1;
    const id   = `SBC0${fNum.padStart(2,"0")}${String(n).padStart(3,"0")}`;
    const rec  = `RCP-${new Date().getFullYear()}-${String(Date.now()).slice(-4)}`;
    const pin  = String(Math.floor(1000 + Math.random()*9000)); // 4-digit parent portal PIN
    try {
      // Photo was already compressed when selected — use as-is
      const studentData = {
        name:form.name, form:form.form, gender:form.gender,
        dob:form.dob||null, parent:form.parent, phone:form.phone,
        address:form.address||null, photo_url:form.photo_url||null,
        id, active:true,
        reg_status:"registered", reg_date:todayStr(),
        reg_fee:fee, reg_receipt:rec,
        reg_paid_by:form.paidBy||form.parent,
        reg_cashier:auth.user.name,
        is_late_reg:form.isLate,
        parent_pin: pin,
      };
      await saveStudent(studentData);
      setReceipt({ ...studentData });
      setTab("receipt");
      setForm(blank);
    } catch(e) { alert("Error saving: "+e.message); }
    setSaving(false);
  }

  function printReceipt() {
    if (!receipt) return;
    const photo  = receipt.photo_url;
    const photoHtml = photo
      ? `<img src="${photo}" style="width:70px;height:84px;object-fit:cover;border:2px solid #0D2340;border-radius:3px;display:block;margin:0 auto 10px">`
      : `<div style="width:70px;height:84px;border:2px dashed #ccc;border-radius:3px;background:#f9fafb;display:flex;align-items:center;justify-content:center;font-size:24px;margin:0 auto 10px">👤</div>`;
    const rows = [
      ["Student Name",receipt.name],["Student ID",receipt.id],["Form",receipt.form],
      ["Gender",receipt.gender],["Date of Birth",fmtDate(receipt.dob)],
      ["Address",receipt.address||"—"],["Parent/Guardian",receipt.parent],
      ["Paid By",receipt.reg_paid_by],["Registration Date",fmtDate(receipt.reg_date)],
      ["Cashier",receipt.reg_cashier],["Academic Year","2026/2027"],
    ].map(([l,v]) => `<div style="display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #f0f0f0;font-size:11.5px"><span style="color:#6B7280">${l}</span><span style="font-weight:600">${v||"—"}</span></div>`).join("");

    const amtCol = receipt.reg_fee===REG_LATE ? "#C9962A" : "#15803D";
    const html = `<div style="font-family:Segoe UI,sans-serif;padding:14px;color:#0D2340">
      <div style="max-width:360px;margin:0 auto;border:2px solid #0D2340;border-radius:10px;overflow:hidden">
        <div style="background:#0D2340;color:#fff;padding:14px;text-align:center">
          <div style="font-size:26px">🎓</div>
          <h2 style="font-size:14px;font-weight:900;margin:4px 0">SAKER BAPTIST COLLEGE</h2>
          <p style="font-size:9px;opacity:.65;margin:0">NGEPTANG · NONI · NW REGION</p>
          <div style="background:#C9962A;display:inline-block;padding:2px 12px;border-radius:10px;font-size:10px;font-weight:800;margin-top:7px">REGISTRATION RECEIPT</div>
          ${receipt.reg_fee===REG_LATE?'<div style="margin-top:4px;background:#fbbf24;color:#78350f;display:inline-block;padding:2px 9px;border-radius:9px;font-size:9px;font-weight:700">LATE REGISTRATION</div>':""}
        </div>
        <div style="padding:14px">
          ${photoHtml}
          <div style="font-family:monospace;font-weight:900;font-size:13px;text-align:center;margin-bottom:9px;padding-bottom:7px;border-bottom:1px dashed #e5e7eb">${receipt.reg_receipt}</div>
          ${rows}
          <div style="background:${receipt.reg_fee===REG_LATE?"#fffbeb":"#f0fdf4"};border-radius:8px;padding:12px;margin-top:11px;text-align:center">
            <div style="font-size:10px;color:#6b7280;margin-bottom:2px">AMOUNT PAID</div>
            <div style="font-size:26px;font-weight:900;color:${amtCol}">${(receipt.reg_fee||0).toLocaleString()} FCFA</div>
            <div style="font-size:10px;color:#6b7280;margin-top:2px">${receipt.reg_fee===REG_LATE?"Late":"Normal"} Registration Fee</div>
          </div>
          <div style="background:#eff6ff;border:1px dashed #93c5fd;border-radius:8px;padding:12px;margin-top:11px;text-align:center">
            <div style="font-size:10px;color:#1e40af;font-weight:700;margin-bottom:4px">📱 PARENT PORTAL ACCESS</div>
            <div style="font-size:10px;color:#374151;margin-bottom:6px">Check marks, attendance & fees online</div>
            <div style="display:flex;justify-content:center;gap:16px">
              <div><div style="font-size:8px;color:#6b7280">Matricule</div><div style="font-family:monospace;font-weight:900;font-size:14px;color:#1e40af">${receipt.id}</div></div>
              <div><div style="font-size:8px;color:#6b7280">PIN</div><div style="font-family:monospace;font-weight:900;font-size:14px;color:#1e40af">${receipt.parent_pin||"----"}</div></div>
            </div>
            <div style="font-size:8px;color:#6b7280;margin-top:6px">Keep this PIN safe — needed to view your child's records</div>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px">
            <div style="border-top:1px solid #999;padding-top:3px;font-size:8px;color:#6b7280;margin-top:24px;text-align:center">Cashier Signature</div>
            <div style="border-top:1px solid #999;padding-top:3px;font-size:8px;color:#6b7280;margin-top:24px;text-align:center">Parent/Guardian Signature</div>
          </div>
        </div>
        <div style="text-align:center;font-size:9px;color:#9ca3af;border-top:1px dashed #e5e7eb;padding:8px 14px">Academic Year 2026/2027 · Keep for your records.</div>
      </div>
    </div>`;
    domPrint("sbc-receipt", html, "A5 portrait", "7mm");
  }

  const registered = students.filter(s=>s.active&&s.reg_status==="registered");
  const pending    = students.filter(s=>s.active&&s.reg_status==="pending");

  return (
    <div>
      <div style={{display:"flex",background:C.white,borderRadius:10,padding:3,marginBottom:13,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        {[["register","➕ Register"],["bulk","📥 Bulk Import"],["list","📋 List"],["receipt","🧾 Receipt"]].map(([k,l]) => (
          <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"9px 4px",borderRadius:7,border:"none",cursor:"pointer",fontWeight:700,fontSize:11,background:tab===k?C.navy:"transparent",color:tab===k?C.white:C.gray}}>{l}</button>
        ))}
      </div>

      {tab==="register" && (
        <div>
          <div style={{background:form.isLate?"#fffbeb":"#f0fdf4",border:`1px solid ${form.isLate?"#fbbf24":C.green}`,borderRadius:10,padding:"8px 12px",marginBottom:11,fontSize:12}}>
            <strong style={{color:form.isLate?C.gold:C.green}}>{form.isLate?"⚠️ Late":"✅ Normal"} Registration</strong>
            <span style={{color:C.gray,marginLeft:6}}>Fee: <strong>{form.isLate?"1,000":"500"} FCFA</strong></span>
          </div>
          <Card title="Student Information">
            <div style={{display:"flex",alignItems:"flex-start",gap:13,marginBottom:11}}>
              <div style={{flexShrink:0,textAlign:"center"}}>
                <PhotoBox photo={form.photo_url} size={[78,94]}/>
                <input ref={photoInputRef} type="file" accept="image/*" style={{display:"none"}} onChange={async e=>{
                  const f=e.target.files[0]; if(!f) return;
                  const r=new FileReader();
                  r.onload=async ev=>{
                    const compressed = await compressPhoto(ev.target.result);
                    setForm(prev=>({...prev,photo_url:compressed}));
                  };
                  r.readAsDataURL(f);
                  e.target.value=""; // allow choosing the same file again after Remove
                }}/>
                <button type="button" onClick={()=>photoInputRef.current?.click()} style={{display:"block",width:"100%",marginTop:4,fontSize:10,color:C.navyMid,cursor:"pointer",fontWeight:700,background:C.grayBg,border:"none",borderRadius:4,padding:"5px 5px",textAlign:"center"}}>
                  📷 {form.photo_url?"Change":"Upload"}
                </button>
                {form.photo_url && <button type="button" onClick={()=>setForm(f=>({...f,photo_url:null}))} style={{marginTop:2,fontSize:9,color:C.red,background:"none",border:"none",cursor:"pointer"}}>✕ Remove</button>}
              </div>
              <div style={{flex:1}}>
                <Fr label="Full Name *"><input style={inp} value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></Fr>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9}}>
                  <Fr label="Form"><select style={inp} value={form.form} onChange={e=>setForm(f=>({...f,form:e.target.value}))}>{FORMS.map(f=><option key={f}>{f}</option>)}</select></Fr>
                  <Fr label="Gender"><select style={inp} value={form.gender} onChange={e=>setForm(f=>({...f,gender:e.target.value}))}><option>Female</option><option>Male</option></select></Fr>
                </div>
              </div>
            </div>
            <Fr label="Date of Birth"><input style={inp} type="date" value={form.dob} onChange={e=>setForm(f=>({...f,dob:e.target.value}))}/></Fr>
            <Fr label="Home Address"><input style={inp} value={form.address} onChange={e=>setForm(f=>({...f,address:e.target.value}))}/></Fr>
          </Card>

          <Card title="Parent / Guardian" style={{marginTop:10}}>
            <Fr label="Name *"><input style={inp} value={form.parent} onChange={e=>setForm(f=>({...f,parent:e.target.value}))}/></Fr>
            <Fr label="Phone *"><input style={inp} type="tel" value={form.phone} onChange={e=>setForm(f=>({...f,phone:e.target.value}))}/></Fr>
            <Fr label="Paid By (if different)"><input style={inp} value={form.paidBy} onChange={e=>setForm(f=>({...f,paidBy:e.target.value}))}/></Fr>
          </Card>

          <Card title="Registration Type" style={{marginTop:10}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9}}>
              {[[false,"✅ Normal","500 FCFA"],[true,"⚠️ Late","1,000 FCFA"]].map(([v,l,a]) => (
                <button key={String(v)} onClick={()=>setForm(f=>({...f,isLate:v}))} style={{padding:"11px",borderRadius:10,border:`2px solid ${form.isLate===v?(v?C.gold:C.green):C.grayLight}`,background:form.isLate===v?(v?"#fffbeb":"#f0fdf4"):C.white,fontWeight:700,fontSize:12,cursor:"pointer",color:form.isLate===v?(v?C.gold:C.green):C.gray,textAlign:"center"}}>
                  {l}<br/><span style={{fontSize:11}}>{a}</span>
                </button>
              ))}
            </div>
          </Card>

          <button onClick={handleRegister} disabled={saving||!form.name.trim()||!form.parent.trim()||!form.phone.trim()} style={{width:"100%",marginTop:13,padding:"13px",background:C.navy,color:C.white,border:"none",borderRadius:10,fontWeight:800,fontSize:14,cursor:"pointer",opacity:(saving||!form.name.trim()||!form.parent.trim()||!form.phone.trim())?0.45:1}}>
            {saving ? "Saving to Supabase…" : "Register & Generate Receipt →"}
          </button>
        </div>
      )}

      {tab==="bulk" && (
        <BulkImportPanel students={students} saveStudent={saveStudent} auth={auth} onDone={()=>setTab("list")}/>
      )}

      {tab==="list" && (
        <div>
          <div style={{fontSize:12,color:C.gray,marginBottom:9}}>
            <strong style={{color:C.green}}>{registered.length}</strong> registered · <strong style={{color:C.red}}>{pending.length}</strong> pending
          </div>
          {students.filter(s=>s.active).sort((a,b)=>a.reg_status.localeCompare(b.reg_status)).map(s => (
            <div key={s.id} style={{background:C.white,borderRadius:10,padding:"11px 12px",marginBottom:7,borderLeft:`4px solid ${s.reg_status==="registered"?C.green:C.red}`,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
                <div style={{display:"flex",gap:9,alignItems:"center"}}>
                  <PhotoBox photo={s.photo_url} size={[36,44]}/>
                  <div><div style={{fontWeight:700,color:C.navy,fontSize:13}}>{s.name}</div><div style={{fontSize:10,fontFamily:"monospace",color:C.gold}}>{s.id}</div><div style={{fontSize:10,color:C.gray}}>{s.form}</div></div>
                </div>
                <Pill color={s.reg_status==="registered"?C.green:C.red}>{s.reg_status==="registered"?"Registered":"Pending"}</Pill>
              </div>
              {s.reg_status==="registered" && (
                <div style={{display:"flex",gap:11,flexWrap:"wrap",marginTop:7}}>
                  <Info label="Fee"     value={`${(s.reg_fee||0).toLocaleString()} F`}/>
                  <Info label="Date"    value={fmtDate(s.reg_date)}/>
                  <Info label="Receipt" value={s.reg_receipt}/>
                </div>
              )}
              {s.reg_status==="registered" && (
                <SmBtn onClick={()=>{setReceipt({...s,reg_paid_by:s.reg_paid_by});setTab("receipt");}} color={C.navy} style={{marginTop:7}}>View Receipt</SmBtn>
              )}
            </div>
          ))}
        </div>
      )}

      {tab==="receipt" && (
        <div>
          {!receipt ? <Empty text="Register a student first to see their receipt."/> : (
            <div>
              <div style={{background:C.white,borderRadius:11,overflow:"hidden",border:`2px solid ${C.navy}`,maxWidth:360,margin:"0 auto 13px"}}>
                <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,color:C.white,padding:"14px",textAlign:"center"}}>
                  <div style={{fontSize:26}}>🎓</div>
                  <h2 style={{margin:"4px 0",fontSize:14,fontWeight:900}}>SAKER BAPTIST COLLEGE</h2>
                  <p style={{margin:0,opacity:.6,fontSize:9}}>NGEPTANG · NONI · NW REGION</p>
                  <div style={{marginTop:7,background:C.gold,display:"inline-block",borderRadius:11,padding:"2px 12px",fontSize:10,fontWeight:800}}>REGISTRATION RECEIPT</div>
                </div>
                <div style={{padding:"12px 13px"}}>
                  {receipt.photo_url && <img src={receipt.photo_url} alt="passport" style={{width:70,height:84,objectFit:"cover",border:`2px solid ${C.navy}`,borderRadius:3,display:"block",margin:"0 auto 10px"}}/>}
                  <div style={{fontFamily:"monospace",fontWeight:900,fontSize:13,textAlign:"center",marginBottom:9,color:C.navy}}>{receipt.reg_receipt}</div>
                  {[["Name",receipt.name],["ID",receipt.id],["Form",receipt.form],["Fee",`${(receipt.reg_fee||0).toLocaleString()} FCFA`],["Date",fmtDate(receipt.reg_date)],["Parent",receipt.parent]].map(([l,v])=>(
                    <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:`1px solid ${C.grayBg}`,fontSize:11}}>
                      <span style={{color:C.gray}}>{l}</span><span style={{fontWeight:600,color:C.navy}}>{v||"—"}</span>
                    </div>
                  ))}
                  <div style={{background:receipt.reg_fee===REG_LATE?"#fffbeb":"#f0fdf4",borderRadius:8,padding:"11px",margin:"11px 0",textAlign:"center"}}>
                    <div style={{fontSize:10,color:C.gray,marginBottom:1}}>AMOUNT PAID</div>
                    <div style={{fontSize:26,fontWeight:900,color:receipt.reg_fee===REG_LATE?C.gold:C.green}}>{(receipt.reg_fee||0).toLocaleString()} FCFA</div>
                  </div>
                  <div style={{background:"#eff6ff",border:"1px dashed #93c5fd",borderRadius:8,padding:"11px",textAlign:"center"}}>
                    <div style={{fontSize:10,color:"#1e40af",fontWeight:700,marginBottom:4}}>📱 PARENT PORTAL ACCESS</div>
                    <div style={{display:"flex",justifyContent:"center",gap:20}}>
                      <div><div style={{fontSize:9,color:C.gray}}>Matricule</div><div style={{fontFamily:"monospace",fontWeight:900,fontSize:15,color:"#1e40af"}}>{receipt.id}</div></div>
                      <div><div style={{fontSize:9,color:C.gray}}>PIN</div><div style={{fontFamily:"monospace",fontWeight:900,fontSize:15,color:"#1e40af"}}>{receipt.parent_pin||"----"}</div></div>
                    </div>
                  </div>
                </div>
              </div>
              <button onClick={printReceipt} style={{width:"100%",maxWidth:360,display:"block",margin:"0 auto",padding:"12px",background:C.navy,color:C.white,border:"none",borderRadius:9,fontWeight:800,fontSize:14,cursor:"pointer"}}>🖨 Print Receipt</button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── Bulk Import Panel ───────────────────────────────────────────────────────
// Expected spreadsheet columns (header row, case-insensitive, order doesn't matter):
// Name | Form | Gender | Date of Birth | Parent/Guardian | Phone | Address
function BulkImportPanel({ students, saveStudent, auth, onDone }) {
  const [rows,     setRows]     = useState([]);   // parsed + validated rows
  const [fileName, setFileName] = useState("");
  const [importing,setImporting]= useState(false);
  const [progress,  setProgress] = useState({done:0,total:0});
  const [results,   setResults]  = useState(null); // {success:[], failed:[]}
  const fileInputRef = useRef(null);

  function downloadTemplate() {
    const sample = [
      { "Name":"John Nkeng Tabi", "Form":"Form 1", "Gender":"Male",   "Date of Birth":"2013-04-15", "Parent/Guardian":"Mrs. Tabi", "Phone":"677123456", "Address":"Ngeptang" },
      { "Name":"Grace Fon Achu",  "Form":"Form 2", "Gender":"Female", "Date of Birth":"2012-09-02", "Parent/Guardian":"Mr. Achu",  "Phone":"677987654", "Address":"Noni" },
    ];
    exportToExcel(sample, "SBC-Bulk-Import-Template.xlsx", "Students");
  }

  function normalizeGender(v) {
    const s = String(v||"").trim().toLowerCase();
    if (s.startsWith("m")) return "Male";
    if (s.startsWith("f")) return "Female";
    return "";
  }
  function normalizeForm(v) {
    const s = String(v||"").trim();
    const match = FORMS.find(f => f.toLowerCase()===s.toLowerCase() || f.replace("Form ","")===s);
    return match || "";
  }
  function normalizeDate(v) {
    if (!v) return "";
    // Handle Excel serial dates
    if (typeof v === "number") {
      const d = XLSX.SSF.parse_date_code(v);
      if (d) return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
    }
    const s = String(v).trim();
    // Accept YYYY-MM-DD directly
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    // Try to parse common formats
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0,10);
    return "";
  }

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    setResults(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary", cellDates:false });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(sheet, { defval:"" });

        // Track how many students already exist per form, to continue ID numbering correctly
        const formCounts = {};
        FORMS.forEach(f => { formCounts[f] = students.filter(s=>s.form===f).length; });

        const parsed = raw.map((r, idx) => {
          // Case-insensitive column lookup
          const get = (...keys) => {
            for (const k of Object.keys(r)) {
              if (keys.some(want => k.toLowerCase().replace(/[^a-z]/g,"") === want)) return r[k];
            }
            return "";
          };
          const name   = String(get("name","fullname","studentname")||"").trim();
          const form   = normalizeForm(get("form","class"));
          const gender = normalizeGender(get("gender","sex"));
          const dob    = normalizeDate(get("dateofbirth","dob","birthdate"));
          const parent = String(get("parentguardian","parent","guardian")||"").trim();
          const phone  = String(get("phone","phonenumber","contact")||"").trim();
          const address= String(get("address","homeaddress")||"").trim();

          const errors = [];
          if (!name)   errors.push("Missing name");
          if (!form)   errors.push("Invalid/missing form");
          if (!gender) errors.push("Invalid/missing gender");
          if (!parent) errors.push("Missing parent/guardian");
          if (!phone)  errors.push("Missing phone");

          let id = "";
          if (form && !errors.length) {
            formCounts[form] = (formCounts[form]||0) + 1;
            const fNum = form.replace("Form ","");
            id = `SBC0${fNum.padStart(2,"0")}${String(formCounts[form]).padStart(3,"0")}`;
          }

          return { rowNum:idx+2, name, form, gender, dob, parent, phone, address, id, errors, valid: errors.length===0 };
        });

        setRows(parsed);
      } catch(err) {
        alert("Could not read file: " + err.message + "\n\nMake sure it's a valid .xlsx or .csv file with a header row.");
        setRows([]);
      }
    };
    reader.readAsBinaryString(file);
  }

  async function registerAll() {
    const valid = rows.filter(r => r.valid);
    if (!valid.length) { alert("No valid rows to import."); return; }
    if (!window.confirm(`Register ${valid.length} student(s)? This cannot be undone in bulk.`)) return;

    setImporting(true);
    setProgress({done:0, total:valid.length});
    const success = [];
    const failed = [];

    for (const r of valid) {
      try {
        const rec = `RCP-${new Date().getFullYear()}-${String(Date.now()).slice(-4)}${Math.floor(Math.random()*90+10)}`;
        await saveStudent({
          id: r.id, name: r.name, form: r.form, gender: r.gender,
          dob: r.dob||null, parent: r.parent, phone: r.phone, address: r.address||null,
          photo_url: null, active: true,
          reg_status: "registered", reg_date: todayStr(),
          reg_fee: REG_NORMAL, reg_receipt: rec,
          reg_paid_by: r.parent, reg_cashier: auth.user.name,
          is_late_reg: false,
        });
        success.push(r);
      } catch(e) {
        failed.push({ ...r, importError: e.message });
      }
      setProgress(p => ({ ...p, done: p.done+1 }));
    }

    setResults({ success, failed });
    setImporting(false);
    setRows([]);
  }

  const validCount   = rows.filter(r=>r.valid).length;
  const invalidCount = rows.length - validCount;

  return (
    <div>
      <div style={{background:"#eff6ff",border:"1px solid #bfdbfe",borderRadius:10,padding:"10px 12px",marginBottom:12,fontSize:12,color:"#1e40af"}}>
        ℹ️ Upload an Excel (.xlsx) or CSV file with columns: <strong>Name, Form, Gender, Date of Birth, Parent/Guardian, Phone, Address</strong>. Column order doesn't matter.
        <div style={{marginTop:6}}>
          <SmBtn onClick={downloadTemplate} color={C.navyMid}>📥 Download Template</SmBtn>
        </div>
      </div>

      {!results && (
        <Card title="1. Upload File">
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{display:"none"}}/>
          <button onClick={()=>fileInputRef.current?.click()} style={{width:"100%",padding:"14px",background:C.grayBg,border:`2px dashed ${C.border}`,borderRadius:10,cursor:"pointer",fontSize:13,fontWeight:700,color:C.navyMid}}>
            📁 {fileName || "Tap to choose a file…"}
          </button>
        </Card>
      )}

      {rows.length>0 && !results && (
        <Card title="2. Preview & Validate" style={{marginTop:10}}>
          <div style={{display:"flex",gap:10,marginBottom:10,flexWrap:"wrap"}}>
            <Pill color={C.green}>{validCount} valid</Pill>
            {invalidCount>0 && <Pill color={C.red}>{invalidCount} with errors</Pill>}
          </div>
          <div style={{maxHeight:300,overflowY:"auto",border:`1px solid ${C.grayLight}`,borderRadius:8}}>
            {rows.map(r => (
              <div key={r.rowNum} style={{padding:"8px 10px",borderBottom:`1px solid ${C.grayBg}`,background:r.valid?C.white:"#fef2f2"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div style={{fontWeight:700,fontSize:12,color:C.navy}}>Row {r.rowNum}: {r.name||"(no name)"}</div>
                  {r.valid
                    ? <Pill color={C.green}>✓ {r.id}</Pill>
                    : <Pill color={C.red}>✕ Error</Pill>
                  }
                </div>
                {r.valid
                  ? <div style={{fontSize:11,color:C.gray,marginTop:2}}>{r.form} · {r.gender} · {r.parent} · {r.phone}</div>
                  : <div style={{fontSize:11,color:C.red,marginTop:2}}>{r.errors.join(", ")}</div>
                }
              </div>
            ))}
          </div>
          {importing ? (
            <div style={{marginTop:12}}>
              <div style={{background:C.grayBg,borderRadius:6,height:10,overflow:"hidden"}}>
                <div style={{width:`${(progress.done/progress.total)*100}%`,height:"100%",background:C.green,transition:"width .3s"}}/>
              </div>
              <div style={{textAlign:"center",fontSize:12,color:C.gray,marginTop:6}}>Registering {progress.done} / {progress.total}…</div>
            </div>
          ) : (
            <button onClick={registerAll} disabled={validCount===0} style={{width:"100%",marginTop:12,padding:"13px",background:validCount?C.navy:C.grayLight,color:C.white,border:"none",borderRadius:10,fontWeight:800,fontSize:14,cursor:validCount?"pointer":"not-allowed"}}>
              Register {validCount} Valid Student{validCount!==1?"s":""} →
            </button>
          )}
        </Card>
      )}

      {results && (
        <Card title="Import Complete" style={{marginTop:10}}>
          <div style={{display:"flex",gap:10,marginBottom:12,flexWrap:"wrap"}}>
            <Pill color={C.green}>✓ {results.success.length} registered</Pill>
            {results.failed.length>0 && <Pill color={C.red}>✕ {results.failed.length} failed</Pill>}
          </div>
          {results.failed.length>0 && (
            <div style={{marginBottom:12}}>
              <div style={{fontSize:12,fontWeight:700,color:C.red,marginBottom:6}}>Failed rows:</div>
              {results.failed.map(r=>(
                <div key={r.rowNum} style={{fontSize:11,color:C.gray,padding:"4px 0",borderBottom:`1px solid ${C.grayBg}`}}>
                  Row {r.rowNum} ({r.name}): {r.importError}
                </div>
              ))}
            </div>
          )}
          <div style={{display:"flex",gap:8}}>
            <Btn onClick={()=>{setResults(null);setFileName("");}} outline>Import Another File</Btn>
            <Btn onClick={onDone}>View Student List →</Btn>
          </div>
        </Card>
      )}
    </div>
  );
}

// ─── Students ──────────────────────────────────────────────────────────────────
function StudentsPage({ ctx }) {
  const { students, saveStudent, deleteStudent, auth, feesMap } = ctx;
  const [filter, setFilter] = useState({ form:"", search:"" });
  const [modal,  setModal]  = useState(null);
  const [viewS,  setViewS]  = useState(null);
  const [saving, setSaving] = useState(false);
  const [showPromote, setShowPromote] = useState(false);
  const blank = { name:"", form:"Form 1", gender:"Female", dob:"", parent:"", phone:"", address:"", photo_url:null };
  const [form, setForm] = useState(blank);
  const photoInputRef = useRef(null);

  const filtered = students.filter(s => s.active
    && (!filter.form   || s.form===filter.form)
    && (!filter.search || s.name?.toLowerCase().includes(filter.search.toLowerCase()) || s.id?.includes(filter.search))
  );

  function studentExportRow(s) {
    return {
      "Matricule":    s.id,
      "Name":         s.name,
      "Form":         s.form,
      "Gender":       s.gender,
      "Date of Birth":fmtDate(s.dob),
      "Parent/Guardian": s.parent||"",
      "Phone":        s.phone||"",
      "Address":      s.address||"",
      "Registration Status": s.reg_status==="registered"?"Registered":"Pending",
      "Registration Date":   fmtDate(s.reg_date),
      "Fees Paid (FCFA)":    feesMap?.[s.id]?.paid ?? 0,
      "Fees Total (FCFA)":   feesMap?.[s.id]?.total ?? TOTAL_FEE,
      "Balance (FCFA)":      (feesMap?.[s.id]?.total ?? TOTAL_FEE) - (feesMap?.[s.id]?.paid ?? 0),
    };
  }

  function exportStudents() {
    if (filter.form) {
      // A single form is already selected — one sheet is enough
      const rows = filtered.map(studentExportRow);
      exportToExcel(rows, `SBC-Students-${filter.form.replace(" ","")}-${sel_year_safe()}.xlsx`, filter.form);
    } else {
      // No form filter — split into one sheet per class, all in one workbook
      const sheets = FORMS.map(f => ({
        name: f,
        rows: filtered.filter(s=>s.form===f).map(studentExportRow),
      }));
      exportToExcelMultiSheet(sheets, `SBC-Students-AllForms-${sel_year_safe()}.xlsx`);
    }
  }
  function sel_year_safe(){ return "2026-2027"; }

  async function save() {
    if (!form.name?.trim()) return;
    setSaving(true);
    try {
      if (modal==="add") {
        const fNum = form.form.replace("Form ","");
        const n    = students.filter(s=>s.form===form.form).length + 1;
        const id   = `SBC0${fNum.padStart(2,"0")}${String(n).padStart(3,"0")}`;
        await saveStudent({ ...form, id, active:true, reg_status:"pending" });
      } else {
        await saveStudent({ ...form });
      }
      setModal(null);
    } catch(e) { alert("Error: "+e.message); }
    setSaving(false);
  }

  const missingPinCount = students.filter(s=>s.active && !s.parent_pin).length;

  async function generateMissingPins() {
    const toFix = students.filter(s=>s.active && !s.parent_pin);
    if (!toFix.length) return;
    if (!window.confirm(`Generate parent portal PINs for ${toFix.length} student(s) who don't have one yet?`)) return;
    setSaving(true);
    for (const s of toFix) {
      try { await saveStudent(s); } catch(e) { console.error(`Failed for ${s.id}:`, e); }
    }
    setSaving(false);
    alert(`Done. Generated PINs for ${toFix.length} student(s).`);
  }

  return (
    <div>
      <div style={{display:"flex",gap:8,marginBottom:11,flexWrap:"wrap"}}>
        <input style={{...inp,flex:1,minWidth:130}} placeholder="Search name or ID…" value={filter.search} onChange={e=>setFilter(f=>({...f,search:e.target.value}))}/>
        <select style={{...inp,width:114}} value={filter.form} onChange={e=>setFilter(f=>({...f,form:e.target.value}))}>
          <option value="">All Forms</option>{FORMS.map(f=><option key={f}>{f}</option>)}
        </select>
        {auth.role==="admin" && <SmBtn onClick={exportStudents} color={C.green}>📊 Export Excel</SmBtn>}
        {auth.role==="admin" && <SmBtn onClick={()=>setShowPromote(true)} color={C.gold}>🎓 Promote Students</SmBtn>}
        {auth.role==="admin" && missingPinCount>0 && <SmBtn onClick={generateMissingPins} color={C.navyMid}>📱 Generate {missingPinCount} Missing PIN{missingPinCount!==1?"s":""}</SmBtn>}
        {auth.role==="admin" && <Btn onClick={()=>{setForm(blank);setModal("add");}}>+ Add</Btn>}
      </div>

      {filtered.map(s => (
        <div key={s.id} style={{background:C.white,borderRadius:10,padding:"11px 12px",marginBottom:7,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
            <div style={{display:"flex",gap:9,alignItems:"center"}}>
              <PhotoBox photo={s.photo_url} size={[38,46]}/>
              <div>
                <div style={{fontWeight:700,color:C.navy}}>{s.name}</div>
                <div style={{fontSize:10,fontFamily:"monospace",color:C.gold}}>{s.id}</div>
              </div>
            </div>
            <Pill color={C.navyMid}>{s.form}</Pill>
          </div>
          <div style={{display:"flex",gap:10,flexWrap:"wrap",marginTop:6}}>
            <Info label="Gender" value={s.gender}/>
            <Info label="Parent" value={s.parent}/>
            <Info label="Status" value={s.reg_status==="registered"?"✅ Registered":"⏳ Pending"}/>
          </div>
          <div style={{display:"flex",gap:6,marginTop:7}}>
            <SmBtn onClick={()=>setViewS(s)} color={C.navyMid}>View</SmBtn>
            {auth.role==="admin" && <SmBtn onClick={()=>{setForm({...s,photo_url:s.photo_url});setModal(s);}} color={C.green}>Edit</SmBtn>}
            {auth.role==="admin" && <SmBtn onClick={()=>{if(confirm(`Remove ${s.name}?`)) deleteStudent(s.id);}} color={C.red}>Remove</SmBtn>}
          </div>
        </div>
      ))}
      {!filtered.length && <Empty text="No students found."/>}

      {viewS && (
        <Modal title={viewS.id} onClose={()=>setViewS(null)}>
          <div style={{display:"flex",gap:12,alignItems:"flex-start",marginBottom:12}}>
            <PhotoBox photo={viewS.photo_url} size={[70,84]}/>
            <div>
              <div style={{fontWeight:800,fontSize:15,color:C.navy}}>{viewS.name}</div>
              <div style={{fontSize:11,color:C.gold,fontFamily:"monospace"}}>{viewS.id}</div>
              <div style={{fontSize:12,color:C.gray}}>{viewS.form} · {viewS.gender}</div>
            </div>
          </div>
          {[["DOB",fmtDate(viewS.dob)],["Parent",viewS.parent],["Phone",viewS.phone],["Address",viewS.address],["Status",viewS.reg_status],["Receipt",viewS.reg_receipt||"—"],["Reg Date",fmtDate(viewS.reg_date)]].map(([l,v])=>(
            <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.grayBg}`}}>
              <span style={{fontSize:12,color:C.gray}}>{l}</span>
              <span style={{fontSize:12,fontWeight:600,color:C.navy}}>{v||"—"}</span>
            </div>
          ))}
          <div style={{background:"#eff6ff",borderRadius:8,padding:"10px 12px",marginTop:10}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontSize:11,color:"#1e40af",fontWeight:700}}>📱 Parent Portal PIN</div>
                <div style={{fontFamily:"monospace",fontWeight:900,fontSize:18,color:"#1e40af",marginTop:2}}>{viewS.parent_pin||"—"}</div>
              </div>
              {auth.role==="admin" && (
                <SmBtn onClick={async()=>{
                  const newPin = String(Math.floor(1000+Math.random()*9000));
                  try { await saveStudent({ ...viewS, parent_pin:newPin }); setViewS(v=>({...v, parent_pin:newPin})); }
                  catch(e){ alert("Error: "+e.message); }
                }} color={C.navyMid}>Reset PIN</SmBtn>
              )}
            </div>
          </div>
        </Modal>
      )}

      {modal && (
        <Modal title={modal==="add"?"Add Student":"Edit Student"} onClose={()=>setModal(null)}>
          <div style={{display:"flex",gap:11,alignItems:"flex-start",marginBottom:10}}>
            <div style={{flexShrink:0}}>
              <PhotoBox photo={form.photo_url} size={[62,76]}/>
              <input ref={photoInputRef} type="file" accept="image/*" style={{display:"none"}} onChange={e=>{
                const f=e.target.files[0]; if(!f) return;
                const r=new FileReader();
                r.onload=async ev=>{
                  const compressed = await compressPhoto(ev.target.result);
                  setForm(prev=>({...prev,photo_url:compressed}));
                };
                r.readAsDataURL(f);
                e.target.value="";
              }}/>
              <button type="button" onClick={()=>photoInputRef.current?.click()} style={{display:"block",width:"100%",marginTop:3,fontSize:9,color:C.navyMid,cursor:"pointer",fontWeight:700,background:C.grayBg,border:"none",borderRadius:4,padding:"3px 5px",textAlign:"center"}}>
                📷 Photo
              </button>
            </div>
            <div style={{flex:1}}>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9}}>
                <div style={{gridColumn:"1/-1"}}><label style={lbl}>Full Name</label><input style={inp} value={form.name||""} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></div>
                <div><label style={lbl}>Form</label><select style={inp} value={form.form} onChange={e=>setForm(f=>({...f,form:e.target.value}))}>{FORMS.map(f=><option key={f}>{f}</option>)}</select></div>
                <div><label style={lbl}>Gender</label><select style={inp} value={form.gender} onChange={e=>setForm(f=>({...f,gender:e.target.value}))}><option>Female</option><option>Male</option></select></div>
                <div><label style={lbl}>DOB</label><input style={inp} type="date" value={form.dob||""} onChange={e=>setForm(f=>({...f,dob:e.target.value}))}/></div>
                <div><label style={lbl}>Phone</label><input style={inp} value={form.phone||""} onChange={e=>setForm(f=>({...f,phone:e.target.value}))}/></div>
                <div style={{gridColumn:"1/-1"}}><label style={lbl}>Parent/Guardian</label><input style={inp} value={form.parent||""} onChange={e=>setForm(f=>({...f,parent:e.target.value}))}/></div>
                <div style={{gridColumn:"1/-1"}}><label style={lbl}>Address</label><input style={inp} value={form.address||""} onChange={e=>setForm(f=>({...f,address:e.target.value}))}/></div>
              </div>
            </div>
          </div>
          <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
            <Btn onClick={()=>setModal(null)} outline>Cancel</Btn>
            <Btn onClick={save} disabled={saving}>{saving?"Saving…":modal==="add"?"Add":"Save"}</Btn>
          </div>
        </Modal>
      )}

      {showPromote && <PromoteStudentsModal students={students} saveStudent={saveStudent} saveFee={ctx.saveFee} onClose={()=>setShowPromote(false)}/>}
    </div>
  );
}

// ─── Promote Students (end-of-year bulk promotion) ──────────────────────────────
function PromoteStudentsModal({ students, saveStudent, saveFee, onClose }) {
  const [fromForm, setFromForm] = useState("Form 1");
  const toFormOptions = FORMS.slice(FORMS.indexOf(fromForm)+1);
  const isGraduating = fromForm === "Form 5";
  const [toForm, setToForm] = useState(isGraduating ? "" : (FORMS[FORMS.indexOf(fromForm)+1]||""));
  const [resetFees, setResetFees] = useState(true);
  const [excluded, setExcluded] = useState({}); // studentId → true if excluded from promotion
  const [step, setStep] = useState("select"); // "select" | "confirm" | "running" | "done"
  const [progress, setProgress] = useState({done:0,total:0});
  const [results, setResults] = useState(null);

  useEffect(() => {
    const nextIdx = FORMS.indexOf(fromForm)+1;
    setToForm(fromForm==="Form 5" ? "" : (FORMS[nextIdx]||""));
    setExcluded({});
  }, [fromForm]);

  const eligible = students.filter(s => s.active && s.form===fromForm && s.reg_status==="registered");
  const included = eligible.filter(s => !excluded[s.id]);

  function toggleExclude(id) {
    setExcluded(prev => ({ ...prev, [id]: !prev[id] }));
  }

  async function runPromotion() {
    setStep("running");
    setProgress({done:0, total:included.length});
    const success = [];
    const failed = [];

    for (const s of included) {
      try {
        if (isGraduating) {
          // Graduating students: mark inactive rather than deleting — preserves their history
          await saveStudent({ ...s, active:false, reg_status:s.reg_status, graduated:true });
        } else {
          await saveStudent({ ...s, form: toForm });
          if (resetFees) {
            await saveFee(s.id, 0); // new academic year, fresh balance
          }
        }
        success.push(s);
      } catch(e) {
        failed.push({ ...s, error: e.message });
      }
      setProgress(p => ({ ...p, done: p.done+1 }));
    }

    setResults({ success, failed });
    setStep("done");
  }

  return (
    <Modal title={isGraduating ? "🎓 Graduate Form 5 Students" : `🎓 Promote ${fromForm} → ${toForm}`} onClose={onClose}>
      {step==="select" && (
        <div>
          <Fr label="Promote FROM">
            <select style={inp} value={fromForm} onChange={e=>setFromForm(e.target.value)}>
              {FORMS.map(f=><option key={f}>{f}</option>)}
            </select>
          </Fr>
          {!isGraduating ? (
            <Fr label="Promote TO">
              <select style={inp} value={toForm} onChange={e=>setToForm(e.target.value)}>
                {toFormOptions.map(f=><option key={f}>{f}</option>)}
              </select>
            </Fr>
          ) : (
            <div style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:8,padding:"9px 12px",marginBottom:12,fontSize:12,color:"#92400e"}}>
              Form 5 is the final form. These students will be marked as <strong>graduated</strong> and moved out of active rolls. Their records (marks, report cards, fee history) are preserved for reference.
            </div>
          )}
          {!isGraduating && (
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
              <input type="checkbox" id="resetFees" checked={resetFees} onChange={e=>setResetFees(e.target.checked)}/>
              <label htmlFor="resetFees" style={{fontSize:12,color:C.navy}}>Reset fees to 0 paid for the new academic year</label>
            </div>
          )}

          <div style={{fontSize:12,fontWeight:700,color:C.navy,marginBottom:6}}>
            {eligible.length} student(s) in {fromForm} · {included.length} will be {isGraduating?"graduated":"promoted"}
          </div>
          <div style={{maxHeight:220,overflowY:"auto",border:`1px solid ${C.grayLight}`,borderRadius:8,marginBottom:12}}>
            {!eligible.length
              ? <div style={{padding:16,textAlign:"center",color:C.gray,fontSize:12}}>No registered students in {fromForm}.</div>
              : eligible.map(s => (
                  <div key={s.id} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",borderBottom:`1px solid ${C.grayBg}`}}>
                    <input type="checkbox" checked={!excluded[s.id]} onChange={()=>toggleExclude(s.id)}/>
                    <span style={{fontSize:12,color:excluded[s.id]?C.gray:C.navy,textDecoration:excluded[s.id]?"line-through":"none",flex:1}}>{s.name}</span>
                    <span style={{fontSize:10,fontFamily:"monospace",color:C.gold}}>{s.id}</span>
                  </div>
                ))
            }
          </div>

          <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
            <Btn onClick={onClose} outline>Cancel</Btn>
            <Btn onClick={()=>setStep("confirm")} disabled={!included.length || (!isGraduating && !toForm)} color={isGraduating?C.red:C.gold}>
              {isGraduating ? `Graduate ${included.length} Student(s) →` : `Promote ${included.length} Student(s) →`}
            </Btn>
          </div>
        </div>
      )}

      {step==="confirm" && (
        <div>
          <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:10,padding:"14px",marginBottom:14,textAlign:"center"}}>
            <div style={{fontSize:28,marginBottom:6}}>⚠️</div>
            <div style={{fontWeight:800,color:C.red,fontSize:14,marginBottom:4}}>This action cannot be undone in bulk</div>
            <div style={{fontSize:12,color:C.gray}}>
              {isGraduating
                ? `${included.length} student(s) will be marked as graduated and removed from active rolls.`
                : `${included.length} student(s) will move from ${fromForm} to ${toForm}${resetFees?" and their fee balance will reset to 0":""}.`}
            </div>
          </div>
          <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
            <Btn onClick={()=>setStep("select")} outline>← Back</Btn>
            <Btn onClick={runPromotion} color={isGraduating?C.red:C.gold}>Yes, Confirm →</Btn>
          </div>
        </div>
      )}

      {step==="running" && (
        <div style={{textAlign:"center",padding:"20px 0"}}>
          <div style={{background:C.grayBg,borderRadius:6,height:10,overflow:"hidden",marginBottom:10}}>
            <div style={{width:`${(progress.done/Math.max(progress.total,1))*100}%`,height:"100%",background:C.gold,transition:"width .3s"}}/>
          </div>
          <div style={{fontSize:13,color:C.gray}}>Processing {progress.done} / {progress.total}…</div>
        </div>
      )}

      {step==="done" && results && (
        <div>
          <div style={{display:"flex",gap:10,marginBottom:12,flexWrap:"wrap"}}>
            <Pill color={C.green}>✓ {results.success.length} {isGraduating?"graduated":"promoted"}</Pill>
            {results.failed.length>0 && <Pill color={C.red}>✕ {results.failed.length} failed</Pill>}
          </div>
          {results.failed.length>0 && (
            <div style={{marginBottom:12,maxHeight:150,overflowY:"auto"}}>
              {results.failed.map(f=>(
                <div key={f.id} style={{fontSize:11,color:C.gray,padding:"4px 0",borderBottom:`1px solid ${C.grayBg}`}}>{f.name}: {f.error}</div>
              ))}
            </div>
          )}
          <Btn onClick={onClose}>Done</Btn>
        </div>
      )}
    </Modal>
  );
}

// ─── Teachers ──────────────────────────────────────────────────────────────────
function TeachersPage({ ctx }) {
  const { teachers, saveTeacher } = ctx;
  const [modal,  setModal]  = useState(null);
  const [form,   setForm]   = useState({ name:"", email:"", password:"", subjects:[], forms:[], active:true });
  const [saving, setSaving] = useState(false);
  const [err,    setErr]    = useState("");

  const toggle = (k,v) => setForm(f => ({ ...f, [k]: f[k].includes(v) ? f[k].filter(x=>x!==v) : [...f[k],v] }));

  function randomPassword() {
    // Easy-to-read 8-character password: avoids ambiguous characters (0/O, 1/l/I)
    const chars = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789";
    let p = "";
    for (let i=0;i<8;i++) p += chars[Math.floor(Math.random()*chars.length)];
    return p;
  }

  // Temporary diagnostic — calls the function with plain fetch, bypassing the
  // Supabase client wrapper entirely, so the literal HTTP status and response
  // text are visible with zero interpretation in between.
  async function rawDiagnostic() {
    setErr("Running diagnostic…");
    try {
      const { data: { session } } = await supabase.auth.getSession();
      if (!session?.access_token) { setErr("DIAGNOSTIC: no active session — log out and back in first."); return; }

      const url = `https://xapbkapxpdvdvelcbpyo.supabase.co/functions/v1/manage-teacher`;
      const res = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${session.access_token}`,
          "apikey": session.access_token, // some setups need this too; harmless if not required
        },
        body: JSON.stringify({ action:"create", name:"Diagnostic Test", email:`diagtest${Date.now()}@example.com`, password:"testpass123", subjects:[], forms:[] }),
      });
      const rawText = await res.text();
      setErr(`DIAGNOSTIC RESULT — HTTP ${res.status}: ${rawText}`);
    } catch(e) {
      setErr("DIAGNOSTIC network error: " + e.message);
    }
  }

  async function createTeacher() {
    if (!form.name?.trim()||!form.email?.trim()||!form.password) return;
    setErr(""); setSaving(true);
    try {
      const { data, error } = await supabase.functions.invoke("manage-teacher", {
        body: { action:"create", name:form.name.trim(), email:form.email.trim(), password:form.password, subjects:form.subjects, forms:form.forms },
      });
      if (error) {
        // The Supabase client's error.message is a generic wrapper on non-2xx
        // responses ("Edge Function returned a non-2xx status code"). The
        // actual JSON body our function sent is on error.context, which is
        // the raw fetch Response — read it directly to get the real message.
        let realMessage = error.message;
        try {
          const body = await error.context.json();
          if (body?.error) realMessage = body.error;
        } catch { /* context wasn't readable JSON — fall back to error.message */ }
        throw new Error(realMessage);
      }
      if (data?.error) throw new Error(data.error);
      setModal({ justCreated:true, name:form.name, email:form.email, password:form.password });
    } catch(e) { setErr("Could not create teacher: " + e.message); }
    setSaving(false);
  }

  async function editTeacher() {
    if (!form.name?.trim()) return;
    setSaving(true);
    try { await saveTeacher({ id:modal.id, name:form.name, email:form.email, subjects:form.subjects, forms:form.forms, active:form.active, joined:form.joined }); setModal(null); }
    catch(e) { setErr("Error: "+e.message); }
    setSaving(false);
  }

  async function deleteTeacherFully(t) {
    if (!window.confirm(`Permanently delete ${t.name}'s account? They will no longer be able to log in. This cannot be undone.`)) return;
    try {
      const { data, error } = await supabase.functions.invoke("manage-teacher", {
        body: { action:"delete", teacherId:t.id, userId:t.user_id },
      });
      if (error) {
        let realMessage = error.message;
        try {
          const body = await error.context.json();
          if (body?.error) realMessage = body.error;
        } catch { /* fall back to error.message */ }
        throw new Error(realMessage);
      }
      if (data?.error) throw new Error(data.error);
      alert(`${t.name}'s account has been removed.`);
    } catch(e) { alert("Could not delete: " + e.message); }
  }

  return (
    <div>
      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:9,gap:8}}>
        <SmBtn onClick={rawDiagnostic} color={C.red}>🔧 Run Diagnostic</SmBtn>
        <Btn onClick={()=>{setForm({name:"",email:"",password:randomPassword(),subjects:[],forms:[],active:true,joined:todayStr()});setErr("");setModal("add");}}>+ Create Teacher Account</Btn>
      </div>
      {err && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"10px 12px",color:C.red,fontSize:11,marginBottom:11,wordBreak:"break-word",whiteSpace:"pre-wrap"}}>{err}</div>}
      {teachers.map(t => (
        <div key={t.id} style={{background:C.white,borderRadius:10,padding:13,marginBottom:9,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",borderTop:`4px solid ${t.active?C.navy:C.gray}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div>
              <div style={{fontFamily:"monospace",fontSize:10,color:C.gold}}>{t.id}</div>
              <div style={{fontWeight:800,color:C.navy,fontSize:14}}>{t.name}</div>
              <div style={{fontSize:12,color:C.gray}}>{t.email}</div>
            </div>
            <Pill color={t.active?C.green:C.red}>{t.active?"Active":"Inactive"}</Pill>
          </div>
          <div style={{marginTop:7,display:"flex",flexWrap:"wrap",gap:4}}>{(t.subjects||[]).map(s=><Pill key={s} color={C.navyMid}>{s}</Pill>)}</div>
          <div style={{marginTop:5,display:"flex",flexWrap:"wrap",gap:4}}>{(t.forms||[]).map(f=><Pill key={f} color={C.gold}>{f}</Pill>)}</div>
          <div style={{display:"flex",gap:7,marginTop:8,flexWrap:"wrap"}}>
            <SmBtn onClick={()=>{setForm({...t,password:""});setErr("");setModal(t);}} color={C.green}>Edit</SmBtn>
            <SmBtn onClick={()=>saveTeacher({...t,active:!t.active})} color={t.active?C.red:C.green}>{t.active?"Deactivate":"Activate"}</SmBtn>
            <SmBtn onClick={()=>deleteTeacherFully(t)} color={C.red}>Delete Account</SmBtn>
          </div>
        </div>
      ))}

      {modal && modal!=="add" && modal.justCreated && (
        <Modal title="✅ Teacher Account Created" onClose={()=>setModal(null)}>
          <div style={{background:"#f0fdf4",border:"1px solid #86efac",borderRadius:10,padding:14,marginBottom:14,textAlign:"center"}}>
            <div style={{fontSize:32,marginBottom:6}}>🎉</div>
            <div style={{fontWeight:800,color:C.green,fontSize:14}}>{modal.name} can now log in</div>
          </div>
          <div style={{background:C.grayBg,borderRadius:8,padding:"10px 12px",marginBottom:14}}>
            {[["Email",modal.email],["Password",modal.password]].map(([l,v])=>(
              <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:`1px solid ${C.grayLight}`}}>
                <span style={{fontSize:12,color:C.gray}}>{l}</span>
                <strong style={{fontSize:13,fontFamily:l==="Password"?"monospace":"inherit",color:C.navy}}>{v}</strong>
              </div>
            ))}
          </div>
          <div style={{fontSize:11,color:C.gray,marginBottom:12}}>⚠️ Write this down or share it now — the password won't be shown again. The teacher can change it themselves later from their Profile page.</div>
          <Btn onClick={()=>setModal(null)}>Done</Btn>
        </Modal>
      )}

      {modal && (modal==="add"||(!modal.justCreated)) && (
        <Modal title={modal==="add"?"Create Teacher Account":"Edit Teacher"} onClose={()=>setModal(null)}>
          <Fr label="Full Name"><input style={inp} value={form.name||""} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></Fr>
          <Fr label={modal==="add"?"Email (used for login)":"Email"}><input style={inp} type="email" value={form.email||""} disabled={modal!=="add"} onChange={e=>setForm(f=>({...f,email:e.target.value}))}/></Fr>
          {modal==="add" && (
            <Fr label="Temporary Password">
              <div style={{display:"flex",gap:6}}>
                <input style={{...inp,flex:1,fontFamily:"monospace"}} value={form.password||""} onChange={e=>setForm(f=>({...f,password:e.target.value}))}/>
                <SmBtn onClick={()=>setForm(f=>({...f,password:randomPassword()}))} color={C.navyMid}>🎲 New</SmBtn>
              </div>
            </Fr>
          )}
          <div style={{marginTop:8}}>
            <label style={lbl}>Subjects</label>
            <div style={{display:"flex",flexWrap:"wrap",gap:5,marginTop:5}}>
              {SUBJECTS_BY_FORM["Form 4"].map(s => (
                <button key={s} onClick={()=>toggle("subjects",s)} style={{padding:"4px 8px",borderRadius:12,border:`1px solid ${form.subjects?.includes(s)?C.navy:C.grayLight}`,background:form.subjects?.includes(s)?C.navy:C.white,color:form.subjects?.includes(s)?C.white:C.gray,fontSize:11,cursor:"pointer"}}>{s}</button>
              ))}
            </div>
          </div>
          <div style={{marginTop:8}}>
            <label style={lbl}>Assigned Forms</label>
            <div style={{display:"flex",flexWrap:"wrap",gap:5,marginTop:5}}>
              {FORMS.map(f => (
                <button key={f} onClick={()=>toggle("forms",f)} style={{padding:"4px 8px",borderRadius:12,border:`1px solid ${form.forms?.includes(f)?C.gold:C.grayLight}`,background:form.forms?.includes(f)?C.gold:C.white,color:form.forms?.includes(f)?C.white:C.gray,fontSize:11,cursor:"pointer"}}>{f}</button>
              ))}
            </div>
          </div>
          {err && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"8px 12px",color:C.red,fontSize:12,marginTop:10}}>{err}</div>}
          <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:13}}>
            <Btn onClick={()=>setModal(null)} outline>Cancel</Btn>
            <Btn onClick={modal==="add"?createTeacher:editTeacher} disabled={saving}>{saving?"Saving…":modal==="add"?"Create Account":"Save"}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ─── Marks Entry ───────────────────────────────────────────────────────────────
function MarksPage({ ctx }) {
  const { students, marksMap, saveMark, auth, teachers, loadAll } = ctx;
  const isTeacher  = auth.role==="teacher";
  const teacher    = teachers.find(t=>t.email===auth.user.email);
  const availForms = isTeacher ? (teacher?.forms||[]) : FORMS;

  const [filter, setFilter] = useState({ form:availForms[0]||"Form 1", seq:"SEQ 1", year:"2026/2027", subject:"" });
  const availSubs  = isTeacher
    ? (teacher?.subjects||[]).filter(s=>SUBJECTS_BY_FORM[filter.form]?.includes(s))
    : (SUBJECTS_BY_FORM[filter.form]||[]);
  const subject    = filter.subject || availSubs[0] || "";

  const [coeff,  setCoeff]  = useState(getCoeff(subject));
  const [local,  setLocal]  = useState({});
  const [saving, setSaving] = useState(false);
  const [saved,  setSaved]  = useState(false);

  useEffect(()=>setCoeff(getCoeff(subject)),[subject]);

  const formStudents = students.filter(s=>s.active&&s.form===filter.form&&s.reg_status==="registered");

  useEffect(() => {
    const init = {};
    formStudents.forEach(s => {
      const key = `${s.id}-${subject}-${filter.seq}-${filter.year}`;
      init[s.id] = marksMap[key]?.score ?? "";
    });
    setLocal(init); setSaved(false);
  }, [filter.form, subject, filter.seq, filter.year, students.length]);

  async function saveAll() {
    setSaving(true);
    try {
      // Save all marks one by one
      for (const s of formStudents) {
        const v = local[s.id];
        if (v===""||v===undefined) continue;
        const { error } = await supabase.from("marks").upsert({
          student_id: s.id, subject, seq: filter.seq,
          acad_year: filter.year, score: Number(v),
          coeff: Number(coeff)||1,
          teacher_id: teacher?.id||null,
        }, { onConflict:"student_id,subject,seq,acad_year" });
        if (error) throw error;
      }
      await ctx.loadAll(); // refresh all data
      setSaved(true);
    } catch(e) { alert("Save error: "+e.message); }
    setSaving(false);
  }

  return (
    <div>
      <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><label style={lbl}>Form</label><select style={inp} value={filter.form} onChange={e=>setFilter(f=>({...f,form:e.target.value,subject:""}))}>{availForms.map(f=><option key={f}>{f}</option>)}</select></div>
          <div><label style={lbl}>Sequence</label><select style={inp} value={filter.seq} onChange={e=>setFilter(f=>({...f,seq:e.target.value}))}>{SEQ_LABELS.map(s=><option key={s}>{s}</option>)}</select></div>
          <div><label style={lbl}>Academic Year</label><select style={inp} value={filter.year} onChange={e=>setFilter(f=>({...f,year:e.target.value}))}>{ACAD_YEARS.map(y=><option key={y}>{y}</option>)}</select></div>
          <div><label style={lbl}>Subject</label><select style={inp} value={subject} onChange={e=>setFilter(f=>({...f,subject:e.target.value}))}>{availSubs.map(s=><option key={s}>{s}</option>)}</select></div>
        </div>
        <div style={{marginTop:10,padding:"8px 11px",background:"#fffbeb",borderRadius:7,border:"1px solid #fde68a",display:"flex",alignItems:"center",gap:11,flexWrap:"wrap"}}>
          <span style={{fontSize:12,fontWeight:700,color:"#92400e"}}>📐 Coefficient:</span>
          <div style={{display:"flex",gap:5}}>
            {[1,2,3,4,5,6].map(n=>(
              <button key={n} onClick={()=>setCoeff(n)} style={{width:32,height:32,borderRadius:6,border:`2px solid ${coeff===n?C.navy:C.grayLight}`,background:coeff===n?C.navy:C.white,color:coeff===n?C.white:C.navy,fontWeight:800,fontSize:12,cursor:"pointer"}}>{n}</button>
            ))}
          </div>
          <span style={{fontSize:11,color:C.gray}}>Default for {subject}: <strong>{getCoeff(subject)}</strong></span>
        </div>
      </div>

      <div style={{background:C.white,borderRadius:10,overflow:"hidden",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{padding:"10px 12px",borderBottom:`1px solid ${C.grayLight}`,display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          <div style={{fontSize:12}}><strong style={{color:C.navy}}>{filter.form}</strong><span style={{color:C.gray,marginLeft:5}}>· {subject} · {filter.seq} · Coeff {coeff}</span></div>
          <div style={{display:"flex",gap:7,alignItems:"center"}}>
            {saved && <span style={{fontSize:12,color:C.green,fontWeight:700}}>✓ Saved to Supabase</span>}
            <Btn onClick={saveAll} disabled={saving}>{saving?"Saving…":"Save Marks"}</Btn>
          </div>
        </div>
        <div style={{padding:"5px 11px",background:"#fffbeb",borderBottom:`1px solid ${C.grayLight}`,fontSize:11,color:"#92400e"}}>
          Score out of <strong>20</strong> · Weighted = score × {coeff} = max {20*coeff} pts
        </div>
        {!formStudents.length ? <Empty text="No registered students in this form."/> : formStudents.map((s,i) => {
          const v = local[s.id] ?? "";
          const { grade } = scoreToGrade(v);
          const gc2 = gradeCol(grade);
          return (
            <div key={s.id} style={{padding:"8px 12px",borderBottom:`1px solid ${C.grayBg}`,display:"flex",alignItems:"center",gap:8,background:i%2===0?C.white:C.grayBg,flexWrap:"wrap"}}>
              <div style={{flex:1,minWidth:100,display:"flex",alignItems:"center",gap:8}}>
                <PhotoBox photo={s.photo_url} size={[30,36]}/>
                <div>
                  <div style={{fontWeight:700,color:C.navy,fontSize:13}}>{s.name}</div>
                  <div style={{fontSize:10,fontFamily:"monospace",color:C.gray}}>{s.id}</div>
                </div>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
                <div style={{textAlign:"center"}}>
                  <div style={{fontSize:9,color:C.gray,marginBottom:2}}>Score /20</div>
                  <input type="number" min={0} max={20} style={{width:52,padding:"5px",border:`1px solid ${C.border}`,borderRadius:7,fontSize:14,fontWeight:700,textAlign:"center",color:C.navy}} value={v} onChange={e=>{setLocal(p=>({...p,[s.id]:e.target.value}));setSaved(false);}}/>
                </div>
                <div style={{textAlign:"center",minWidth:36}}>
                  <div style={{fontSize:9,color:C.gray,marginBottom:2}}>×{coeff}</div>
                  <span style={{fontSize:12,fontWeight:700,color:C.navyMid}}>{v!==""?Number(v)*coeff:"—"}</span>
                </div>
                <div style={{textAlign:"center",minWidth:34}}>
                  <div style={{fontSize:9,color:C.gray,marginBottom:3}}>Grade</div>
                  <span style={{background:gc2+"18",color:gc2,borderRadius:6,padding:"3px 6px",fontWeight:800,fontSize:12}}>{grade}</span>
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── Report Cards ──────────────────────────────────────────────────────────────
// ─── Attendance ────────────────────────────────────────────────────────────────
const ATT_STATUSES = ["present","absent","late"];
const ATT_LABELS = { present:"Present", absent:"Absent", late:"Late" };
const ATT_COLORS = { present:C.green, absent:C.red, late:C.gold };
const ATT_ICONS  = { present:"✓", absent:"✕", late:"⏰" };

function AttendancePage({ ctx }) {
  const { students, attendanceMap, auth, teachers, saveAttendanceBulk } = ctx;
  const isTeacher  = auth.role==="teacher";
  const teacher    = teachers.find(t=>t.email===auth.user.email);
  const availForms = isTeacher ? (teacher?.forms||[]) : FORMS;

  const [tab,   setTab]   = useState("mark"); // "mark" | "summary"
  const [form,  setForm]  = useState(availForms[0]||"Form 1");
  const [date,  setDate]  = useState(todayStr());
  const [local, setLocal] = useState({}); // studentId → status
  const [saving,setSaving]= useState(false);
  const [saved, setSaved] = useState(false);

  const formStudents = students.filter(s=>s.active&&s.form===form&&s.reg_status==="registered");

  // Load existing attendance for this form+date whenever either changes
  useEffect(() => {
    const init = {};
    formStudents.forEach(s => {
      const key = `${s.id}-${date}`;
      init[s.id] = attendanceMap[key]?.status || "present"; // default assumption: present
    });
    setLocal(init);
    setSaved(false);
  }, [form, date, students.length]);

  function cycleStatus(studentId) {
    setLocal(prev => {
      const cur = prev[studentId] || "present";
      const idx = ATT_STATUSES.indexOf(cur);
      const next = ATT_STATUSES[(idx+1) % ATT_STATUSES.length];
      return { ...prev, [studentId]: next };
    });
    setSaved(false);
  }

  function markAllPresent() {
    const all = {};
    formStudents.forEach(s => { all[s.id] = "present"; });
    setLocal(all);
    setSaved(false);
  }

  async function saveAll() {
    setSaving(true);
    try {
      const records = formStudents.map(s => ({ studentId: s.id, status: local[s.id]||"present" }));
      await saveAttendanceBulk(date, records, auth.user.name);
      setSaved(true);
    } catch(e) { alert("Save error: " + e.message); }
    setSaving(false);
  }

  const presentCount = formStudents.filter(s=>(local[s.id]||"present")==="present").length;
  const absentCount  = formStudents.filter(s=>local[s.id]==="absent").length;
  const lateCount    = formStudents.filter(s=>local[s.id]==="late").length;

  return (
    <div>
      <div style={{display:"flex",background:C.white,borderRadius:10,padding:3,marginBottom:13,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        {[["mark","📝 Mark Today"],["summary","📊 Term Summary"]].map(([k,l]) => (
          <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"9px 4px",borderRadius:7,border:"none",cursor:"pointer",fontWeight:700,fontSize:12,background:tab===k?C.navy:"transparent",color:tab===k?C.white:C.gray}}>{l}</button>
        ))}
      </div>

      {tab==="mark" && (
        <div>
          <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              <div><label style={lbl}>Form</label>
                <select style={inp} value={form} onChange={e=>setForm(e.target.value)}>{availForms.map(f=><option key={f}>{f}</option>)}</select>
              </div>
              <div><label style={lbl}>Date</label>
                <input style={inp} type="date" value={date} max={todayStr()} onChange={e=>setDate(e.target.value)}/>
              </div>
            </div>
          </div>

          <div style={{display:"flex",gap:8,marginBottom:11,flexWrap:"wrap"}}>
            <Pill color={C.green}>{presentCount} Present</Pill>
            {absentCount>0 && <Pill color={C.red}>{absentCount} Absent</Pill>}
            {lateCount>0 && <Pill color={C.gold}>{lateCount} Late</Pill>}
            <SmBtn onClick={markAllPresent} color={C.navyMid} style={{marginLeft:"auto"}}>Mark All Present</SmBtn>
          </div>

          <div style={{background:C.white,borderRadius:10,overflow:"hidden",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{padding:"8px 12px",background:"#fffbeb",borderBottom:`1px solid ${C.grayLight}`,fontSize:11,color:"#92400e"}}>
              Tap a student's status to cycle: Present → Absent → Late
            </div>
            {!formStudents.length
              ? <Empty text="No registered students in this form."/>
              : formStudents.map((s,i) => {
                  const status = local[s.id] || "present";
                  return (
                    <div key={s.id} style={{padding:"9px 12px",borderBottom:`1px solid ${C.grayBg}`,display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,background:i%2===0?C.white:C.grayBg}}>
                      <div style={{display:"flex",alignItems:"center",gap:8,minWidth:0}}>
                        <PhotoBox photo={s.photo_url} size={[30,36]}/>
                        <div style={{minWidth:0}}>
                          <div style={{fontWeight:700,color:C.navy,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.name}</div>
                          <div style={{fontSize:10,fontFamily:"monospace",color:C.gray}}>{s.id}</div>
                        </div>
                      </div>
                      <button onClick={()=>cycleStatus(s.id)} style={{flexShrink:0,padding:"7px 14px",borderRadius:8,border:`2px solid ${ATT_COLORS[status]}`,background:ATT_COLORS[status]+"18",color:ATT_COLORS[status],fontWeight:800,fontSize:12,cursor:"pointer",minWidth:88,textAlign:"center"}}>
                        {ATT_ICONS[status]} {ATT_LABELS[status]}
                      </button>
                    </div>
                  );
                })
            }
          </div>

          {formStudents.length>0 && (
            <button onClick={saveAll} disabled={saving} style={{width:"100%",marginTop:13,padding:"13px",background:saved?C.green:C.navy,color:C.white,border:"none",borderRadius:10,fontWeight:800,fontSize:14,cursor:"pointer",opacity:saving?0.6:1}}>
              {saving?"Saving…":saved?"✓ Attendance Saved":`Save Attendance — ${form} · ${fmtDate(date)}`}
            </button>
          )}
        </div>
      )}

      {tab==="summary" && <AttendanceSummary form={form} setForm={setForm} availForms={availForms} students={formStudents} attendanceMap={attendanceMap}/>}
    </div>
  );
}

function AttendanceSummary({ form, setForm, availForms, students, attendanceMap }) {
  // Count how many present/absent/late days each student has recorded, all-time (per current academic year context — simple total for now)
  const counts = students.map(s => {
    let present=0, absent=0, late=0;
    Object.keys(attendanceMap).forEach(key => {
      if (!key.startsWith(s.id+"-")) return;
      const st = attendanceMap[key].status;
      if (st==="present") present++;
      else if (st==="absent") absent++;
      else if (st==="late") late++;
    });
    return { student:s, present, absent, late, total: present+absent+late };
  });

  return (
    <div>
      <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <label style={lbl}>Form</label>
        <select style={inp} value={form} onChange={e=>setForm(e.target.value)}>{availForms.map(f=><option key={f}>{f}</option>)}</select>
      </div>
      <div style={{background:C.white,borderRadius:10,overflow:"hidden",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        {!counts.length ? <Empty text="No registered students in this form."/> : counts.map((c,i) => (
          <div key={c.student.id} style={{padding:"10px 12px",borderBottom:`1px solid ${C.grayBg}`,background:i%2===0?C.white:C.grayBg}}>
            <div style={{fontWeight:700,color:C.navy,fontSize:13,marginBottom:6}}>{c.student.name}</div>
            <div style={{display:"flex",gap:14,fontSize:12}}>
              <span style={{color:C.green}}>✓ {c.present} present</span>
              <span style={{color:C.red}}>✕ {c.absent} absent</span>
              <span style={{color:C.gold}}>⏰ {c.late} late</span>
              <span style={{color:C.gray,marginLeft:"auto"}}>{c.total} days recorded</span>
            </div>
          </div>
        ))}
      </div>
      <div style={{fontSize:11,color:C.gray,marginTop:8,textAlign:"center"}}>
        These totals automatically appear on each student's report card under "Absences".
      </div>
    </div>
  );
}

function ReportsPage({ ctx }) {
  const SBC_SUBJECTS_LIST = [
    "English language","French/ Français","Mathematics","Health science",
    "Home management","Citizenship","Food and Nutrition","Chemistry",
    "History","Economics","Geography","Biology","Human Biology","Physics",
    "Literature in English","Computer Studies","Religious Studies",
    "Commerce","Accounts","Hygiene","Sport/Physical education","Manual Labour"
  ];
  const { students, marksMap, teachers, attendanceMap } = ctx;
  const [sel,     setSel]     = useState({ studentId:"", year:"2026/2027", mode:"term", term:"First Term" });
  const [conduct, setConduct] = useState({ present:"", absent:"", late:"", conduct:"Very Good", classTeacherRemark:"", principalRemark:"" });
  const [showC,   setShowC]   = useState(false);

  const student  = students.find(s=>s.id===sel.studentId);
  const subjects = student ? (SUBJECTS_BY_FORM[student.form]||[]) : [];
  const termSeqs = TERM_SEQS[sel.term];

  // Auto-compute real attendance totals for this student from the Attendance page's data
  const attTotals = (()=>{
    if (!student) return { present:0, absent:0, late:0 };
    let present=0, absent=0, late=0;
    Object.keys(attendanceMap||{}).forEach(key => {
      if (!key.startsWith(student.id+"-")) return;
      const st = attendanceMap[key].status;
      if (st==="present") present++;
      else if (st==="absent") absent++;
      else if (st==="late") late++;
    });
    return { present, absent, late };
  })();

  // Pre-fill conduct fields with real attendance whenever the selected student changes,
  // but only if the admin hasn't already typed something for this student
  useEffect(() => {
    if (!student) return;
    setConduct(c => ({
      ...c,
      absent: c.absent===""? String(attTotals.absent) : c.absent,
      present: c.present===""? String(attTotals.present) : c.present,
      late: c.late===""? String(attTotals.late) : c.late,
    }));
  }, [student?.id]);

  const teacherFor = sub => {
    const t = teachers.find(t=>t.active&&(t.subjects||[]).includes(sub)&&(t.forms||[]).includes(student?.form));
    return t?.name || "—";
  };
  const getScore = (sub,seq) => {
    const k = `${sel.studentId}-${sub}-${seq}-${sel.year}`;
    return marksMap[k]?.score ?? null;
  };
  const getC = sub => {
    const k = `${sel.studentId}-${sub}-SEQ 1-${sel.year}`;
    return marksMap[k]?.coeff || getCoeff(sub);
  };

  const subjectRows = subjects.map(sub => {
    const s1=getScore(sub,"SEQ 1"),s2=getScore(sub,"SEQ 2");
    const s3=getScore(sub,"SEQ 3"),s4=getScore(sub,"SEQ 4");
    const s5=getScore(sub,"SEQ 5"),s6=getScore(sub,"SEQ 6");
    const coeff   = getC(sub);
    const teacher = teacherFor(sub);
    const avg2    = (arr) => { const v=arr.filter(x=>x!==null); return v.length ? v.reduce((a,b)=>a+Number(b),0)/v.length : null; };
    const t1Avg=avg2([s1,s2]), t2Avg=avg2([s3,s4]), t3Avg=avg2([s5,s6]);
    const annualAvg   = avg2([t1Avg,t2Avg,t3Avg]);
    const termScores  = termSeqs.map(sq=>[s1,s2,s3,s4,s5,s6][parseInt(sq.replace("SEQ ",""))-1]);
    const termAvg     = avg2(termScores);
    return { sub,coeff,teacher,s1,s2,s3,s4,s5,s6,t1Avg,t2Avg,t3Avg,annualAvg,termScores,termAvg };
  });

  const modeRows = sel.mode==="term"
    ? subjectRows.filter(r=>r.termAvg!==null)
    : subjectRows.filter(r=>r.annualAvg!==null);

  const wAvg = rows => {
    let tw=0,tc=0;
    rows.forEach(r=>{ const a=sel.mode==="term"?r.termAvg:r.annualAvg; if(a!==null){tw+=a*r.coeff;tc+=r.coeff;} });
    return tc ? tw/tc : null;
  };
  const overallAvg  = wAvg(modeRows);
  const totalCoeff  = modeRows.reduce((a,r)=>{ const av=sel.mode==="term"?r.termAvg:r.annualAvg; return a+(av!==null?r.coeff:0); },0);
  const overallGrd  = overallAvg!==null ? scoreToGrade(overallAvg) : { grade:"—",remark:"—" };
  const passCount   = modeRows.filter(r=>{ const a=sel.mode==="term"?r.termAvg:r.annualAvg; return a!==null&&a>=10; }).length;
  const D           = v => v!==null ? Number(v).toFixed(1) : "—";

  function buildPrintHTML() {
    if (!student||!modeRows.length) return "";

    const SBC_SUBJECTS = [
      "English language","French/ Français","Mathematics","Health science",
      "Home management","Citizenship","Food and Nutrition","Chemistry",
      "History","Economics","Geography","Biology","Human Biology","Physics",
      "Literature in English","Computer Studies","Religious Studies",
      "Commerce","Accounts","Hygiene","Sport/Physical education","Manual Labour"
    ];

    // ── Compute class ranking ─────────────────────────────────────────────────
    const peers = students.filter(s=>s.active&&s.form===student.form&&s.reg_status==="registered");
    const peerAvgs = peers.map(ps=>{
      let tw=0,tc=0;
      SBC_SUBJECTS.forEach(sub=>{
        const seqs = sel.mode==="term"
          ? TERM_SEQS[sel.term].map(sq=>{ const k=ps.id+"-"+sub+"-"+sq+"-"+sel.year; return marksMap[k]?.score??null; })
          : SEQ_LABELS.map(sq=>{ const k=ps.id+"-"+sub+"-"+sq+"-"+sel.year; return marksMap[k]?.score??null; });
        const vals=seqs.filter(x=>x!==null);
        const c=getCoeff(sub);
        if(vals.length){tw+=vals.reduce((a,b)=>a+Number(b),0)/vals.length*c;tc+=c;}
      });
      return{id:ps.id,avg:tc?tw/tc:0};
    }).sort((a,b)=>b.avg-a.avg);
    const myRank = peerAvgs.findIndex(x=>x.id===student.id)+1;
    const rankStr = myRank>0?myRank+"/"+peers.length:"—";

    // ── Class average ─────────────────────────────────────────────────────────
    const classAvg = peerAvgs.length
      ? (peerAvgs.reduce((a,x)=>a+x.avg,0)/peerAvgs.length).toFixed(2)
      : "—";

    const photoHtml = student.photo_url
      ?`<img src="${student.photo_url}" style="width:60px;height:72px;object-fit:cover;border:1px solid #1a56a0;display:block">`
      :`<div style="width:60px;height:72px;border:1px solid #999;display:flex;align-items:center;justify-content:center;font-size:8px;color:#999;text-align:center">PHOTO</div>`;

    const termLabel = sel.mode==="term" ? sel.term.replace(" Term","").toUpperCase()+" TERM" : "ANNUAL";
    const seqNums = sel.mode==="term"
      ? termSeqs.map(sq=>sq.replace("SEQ ",""))
      : ["1","2","3","4","5","6"];
    const seqW = sel.mode==="term" ? "9%" : "5.2%";
    const seqHeaders = seqNums.map(n=>`<th style="background:#1a56a0;color:#fff;border:1px solid #fff;padding:2px 1px;font-size:7px;width:${seqW}">SQ${n}</th>`).join("");

    const rows = SBC_SUBJECTS.map((sub,i)=>{
      const r = modeRows.find(x=>
        x.sub===sub||
        (sub==="French/ Français"&&x.sub==="French")||
        (sub==="Computer Studies"&&x.sub==="Computer Science")||
        (sub==="Sport/Physical education"&&x.sub==="Physical Education")||
        (sub==="Literature in English"&&x.sub==="Literature")
      );
      if(!r){
        const emptyCols = sel.mode==="term" ? termSeqs.map(()=>`<td style="border:1px solid #1a56a0"></td>`).join("") : SEQ_LABELS.map(()=>`<td style="border:1px solid #1a56a0"></td>`).join("");
        return`<tr style="background:${i%2===0?"#fff":"#F9FAFB"}">
          <td style="padding:1px 2px;font-size:6.2px;border:1px solid #1a56a0;line-height:1.05;word-break:break-word">${sub}</td>
          ${emptyCols}
          <td style="border:1px solid #1a56a0"></td>
          <td style="border:1px solid #1a56a0"></td>
          <td style="border:1px solid #1a56a0"></td>
          <td style="border:1px solid #1a56a0"></td>
        </tr>`;
      }
      const avg = sel.mode==="term"?r.termAvg:r.annualAvg;
      const {remark} = scoreToGrade(avg);
      const pass = avg!==null&&avg>=10;
      const pond = avg!==null?(avg*r.coeff).toFixed(1):"";
      const scoreCols = sel.mode==="term"
        ? r.termScores.map(s=>`<td style="text-align:center;font-size:7.5px;border:1px solid #1a56a0;padding:2px">${s!=null?s:""}</td>`).join("")
        : SEQ_LABELS.map((_,j)=>{
            const sc=[r.s1,r.s2,r.s3,r.s4,r.s5,r.s6][j];
            return`<td style="text-align:center;font-size:7px;border:1px solid #1a56a0;padding:2px">${sc!=null?sc:""}</td>`;
          }).join("");
      const teacherName = r.teacher && r.teacher!=="—" ? r.teacher : "";
      const subjectCell = teacherName
        ? `${sub}<div style="font-size:6px;color:#1a56a0;font-style:italic;line-height:1.1">${teacherName}</div>`
        : sub;
      return`<tr style="background:${i%2===0?"#fff":"#F9FAFB"}">
        <td style="padding:1px 2px;font-size:6.2px;border:1px solid #1a56a0;line-height:1.05;word-break:break-word">${subjectCell}</td>
        ${scoreCols}
        <td style="text-align:center;font-size:7.5px;font-weight:700;border:1px solid #1a56a0;padding:2px">${avg!=null?avg.toFixed(1):""}</td>
        <td style="text-align:center;font-size:7.5px;border:1px solid #1a56a0;padding:2px">${r.coeff}</td>
        <td style="text-align:center;font-size:7.5px;font-weight:700;border:1px solid #1a56a0;padding:2px">${pond}</td>
        <td style="text-align:center;font-size:6.5px;color:${pass?"#15803D":"#B91C1C"};border:1px solid #1a56a0;padding:2px">${remark.slice(0,3)}</td>
      </tr>`;
    }).join("");


    // Single portrait card sized to exactly 190mm x 277mm (200mm page - 10mm margin, minus header space)
    return`<!DOCTYPE html><html><head><meta charset="utf-8">
<title>Report Card — ${student.name}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  @page{size:200mm 297mm portrait;margin:5mm}
  @media print{
    html,body{width:200mm!important;min-height:297mm!important}
  }
  html,body{font-family:Arial,Helvetica,sans-serif;color:#000;background:#fff;
    -webkit-print-color-adjust:exact;print-color-adjust:exact}
  table{width:100%;border-collapse:collapse}
  .page{width:190mm;max-width:190mm;margin:0 auto}
</style></head><body><div class="page">

<!-- HEADER -->
<table style="margin-bottom:3px">
  <tr>
    <td style="width:33%;font-size:7.5px;line-height:1.5;vertical-align:top">
      <strong>REPUBLIC OF CAMEROON</strong><br>Peace-Work-Fatherland<br>Ministry of Secondary Education
    </td>
    <td style="width:34%;text-align:center;vertical-align:middle;font-size:20px">🇨🇲</td>
    <td style="width:33%;font-size:7.5px;line-height:1.5;vertical-align:top;text-align:right">
      <strong>REPUBLIQUE DU CAMEROUN</strong><br>Paix – Travail – Patrie<br>Ministère De l'Enseignement
    </td>
  </tr>
</table>

<!-- SCHOOL NAME -->
<div style="background:#1a56a0;color:#fff;text-align:center;padding:5px 4px;margin-bottom:3px">
  <div style="font-size:15px;font-weight:900;letter-spacing:.5px">SAKER BAPTIST COLLAGE(SBC)-BAWE</div>
  <div style="font-size:8px;margin-top:1px;opacity:.85">Motto: Quality, Discipline, and Excellence education</div>
</div>

<!-- TITLE -->
<div style="text-align:center;margin-bottom:3px">
  <span style="font-size:11px;font-weight:800;text-decoration:underline">ACADEMIC REPORT CARD</span>
  <span style="font-size:10px;font-weight:700"> / BULLETIN DE NOTES</span>
</div>

<!-- TERM + STUDENT INFO + PHOTO -->
<table style="border:1px solid #000;margin-bottom:3px">
  <tr>
    <td colspan="4" style="border-bottom:1px solid #000;padding:2px 6px;font-size:8px">
      <strong>${termLabel}</strong> &nbsp;|&nbsp; Class: <strong>${student.form}</strong> &nbsp;|&nbsp;
      No. on Roll: <strong>${peers.length}</strong> &nbsp;|&nbsp; Academic Year: <strong>${sel.year}</strong>
    </td>
    <td rowspan="4" style="width:66px;padding:3px;border-left:1px solid #000;vertical-align:top;text-align:center">
      ${photoHtml}
    </td>
  </tr>
  <tr>
    <td colspan="2" style="padding:2px 6px;font-size:8px;border-bottom:1px solid #ccc;border-right:1px solid #ccc">
      Name / Nom: <strong>${student.name}</strong>
    </td>
    <td colspan="2" style="padding:2px 6px;font-size:8px;border-bottom:1px solid #ccc">
      Matricule / ID: <strong>${student.id}</strong>
    </td>
  </tr>
  <tr>
    <td style="padding:2px 6px;font-size:8px;border-bottom:1px solid #ccc;border-right:1px solid #ccc">
      Sex / Sexe: <strong>${student.gender}</strong>
    </td>
    <td style="padding:2px 6px;font-size:8px;border-bottom:1px solid #ccc;border-right:1px solid #ccc">
      DOB / Naissance: <strong>${fmtDate(student.dob)}</strong>
    </td>
    <td colspan="2" style="padding:2px 6px;font-size:8px;border-bottom:1px solid #ccc">
      Parent/Guardian: <strong>${student.parent||"—"}</strong>
    </td>
  </tr>
  <tr>
    <td colspan="4" style="padding:2px 6px;font-size:8px">
      Address / Domicile: <strong>${student.address||"—"}</strong>
    </td>
  </tr>
</table>

<!-- MARKS TABLE -->
<table style="border:1px solid #1a56a0;margin-bottom:3px">
  <thead>
    <tr style="background:#1a56a0;color:#fff">
      <th style="text-align:left;padding:2px 3px;font-size:7px;border:1px solid #fff;width:18%">SUBJECT</th>
      ${seqHeaders}
      <th style="padding:2px 1px;font-size:6.5px;border:1px solid #fff;width:8%">Avg</th>
      <th style="padding:2px 1px;font-size:6.5px;border:1px solid #fff;width:7%">Coeff</th>
      <th style="padding:2px 1px;font-size:6.5px;border:1px solid #fff;width:8%">Score</th>
      <th style="padding:2px 1px;font-size:6.5px;border:1px solid #fff;width:7%">Rmk</th>
    </tr>
  </thead>
  <tbody>${rows}</tbody>

</table>

<!-- SUMMARY -->
<table style="border:1px solid #000;margin-bottom:3px">
  <tr>
    <td style="padding:3px 6px;font-size:8.5px;border-right:1px solid #000;width:25%">
      Terminal Avg: <strong>${overallAvg?overallAvg.toFixed(2):"—"}/20</strong>
    </td>
    <td style="padding:3px 6px;font-size:8.5px;border-right:1px solid #000;width:25%">
      Class Avg: <strong>${classAvg}/20</strong>
    </td>
    <td style="padding:3px 6px;font-size:8.5px;border-right:1px solid #000;width:25%">
      Position / Rang: <strong>${rankStr}</strong>
    </td>
    <td style="padding:3px 6px;font-size:8.5px;width:25%">
      Promoted/Repeat: __________
    </td>
  </tr>
  <tr>
    <td style="padding:3px 6px;font-size:8.5px;border-right:1px solid #000;border-top:1px solid #ccc">
      Annual Avg: <strong>${sel.mode==="annual"&&overallAvg?overallAvg.toFixed(2):"—"}/20</strong>
    </td>
    <td style="padding:3px 6px;font-size:8.5px;border-right:1px solid #000;border-top:1px solid #ccc">
      Absences: <strong>${conduct.absent||"—"}</strong>
    </td>
    <td style="padding:3px 6px;font-size:8.5px;border-right:1px solid #000;border-top:1px solid #ccc">
      No. of Warnings: __________
    </td>
    <td style="padding:3px 6px;font-size:8.5px;border-top:1px solid #ccc">
      Fees Owing: __________
    </td>
  </tr>
  <tr>
    <td colspan="2" style="padding:3px 6px;font-size:8.5px;border-right:1px solid #000;border-top:1px solid #ccc">
      Class Teacher Remark: <strong>${conduct.classTeacherRemark||"_________________________"}</strong>
    </td>
    <td colspan="2" style="padding:3px 6px;font-size:8.5px;border-top:1px solid #ccc">
      Next Term Begins: __________
    </td>
  </tr>
</table>

<!-- SIGNATURES -->
<table style="margin-top:6px">
  <tr>
    <td style="width:33%;text-align:center;padding:0 8px">
      <div style="height:22px;border-bottom:1px solid #000;margin-bottom:2px"></div>
      <div style="font-size:8px">Class Teacher / Prof. Principal</div>
    </td>
    <td style="width:34%;text-align:center;padding:0 8px">
      <div style="height:22px;border-bottom:1px solid #000;margin-bottom:2px"></div>
      <div style="font-size:8px">Principal's Signature / Stamp</div>
    </td>
    <td style="width:33%;text-align:center;padding:0 8px">
      <div style="height:22px;border-bottom:1px solid #000;margin-bottom:2px"></div>
      <div style="font-size:8px">Parent / Guardian</div>
    </td>
  </tr>
</table>

<div style="text-align:center;font-size:7px;color:#888;margin-top:5px;border-top:1px dashed #ccc;padding-top:3px">
  Saker Baptist Collage (SBC)-Bawe · ${sel.year} · Printed: ${new Date().toLocaleDateString("en-GB")}
</div>
</div></body></html>`;
  }


  function printReport() {
    if (!student||!modeRows.length) return;
    domPrint("sbc-report", buildPrintHTML(), "200mm 297mm portrait", "5mm");
  }

  // Guard: show error if something crashes
  try {

  return (
    <div>
      {/* Controls */}
      <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div style={{gridColumn:"1/-1"}}>
            <label style={lbl}>Select Student</label>
            <select style={inp} value={sel.studentId} onChange={e=>setSel(s=>({...s,studentId:e.target.value}))}>
              <option value="">— Choose a student —</option>
              {FORMS.map(f=>(
                <optgroup key={f} label={f}>
                  {students.filter(s=>s.form===f&&s.active).map(s=><option key={s.id} value={s.id}>{s.name} ({s.id})</option>)}
                </optgroup>
              ))}
            </select>
          </div>
          <div><label style={lbl}>Academic Year</label><select style={inp} value={sel.year} onChange={e=>setSel(s=>({...s,year:e.target.value}))}>{ACAD_YEARS.map(y=><option key={y}>{y}</option>)}</select></div>
          <div><label style={lbl}>Report Type</label><select style={inp} value={sel.mode} onChange={e=>setSel(s=>({...s,mode:e.target.value}))}><option value="term">Term Report</option><option value="annual">Annual Report</option></select></div>
          {sel.mode==="term" && (
            <div style={{gridColumn:"1/-1"}}>
              <label style={lbl}>Term</label>
              <div style={{display:"flex",gap:6}}>
                {TERMS.map(t=>(
                  <button key={t} onClick={()=>setSel(s=>({...s,term:t}))} style={{flex:1,padding:"8px 4px",borderRadius:8,border:`2px solid ${sel.term===t?C.navy:C.grayLight}`,background:sel.term===t?C.navy:C.white,color:sel.term===t?C.white:C.gray,fontWeight:700,fontSize:11,cursor:"pointer"}}>{t}</button>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>

      {!sel.studentId && <Empty text="Select a student to preview and print their report card."/>}

      {student && (
        <div>
          {/* Photo + conduct */}
          <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:9}}>
              <span style={{fontWeight:800,color:C.navy,fontSize:13}}>📸 Photo & Conduct</span>
              <button onClick={()=>setShowC(v=>!v)} style={{fontSize:11,padding:"4px 10px",background:C.grayBg,border:`1px solid ${C.grayLight}`,borderRadius:6,cursor:"pointer",color:C.navy,fontWeight:600}}>{showC?"▲ Hide":"▼ Fill Conduct"}</button>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:showC?10:0}}>
              <PhotoBox photo={student.photo_url} size={[64,78]}/>
              <div style={{flex:1}}>
                <div style={{fontWeight:800,color:C.navy,fontSize:14}}>{student.name}</div>
                <div style={{fontSize:11,color:C.gold,fontFamily:"monospace"}}>{student.id}</div>
                <div style={{fontSize:11,color:C.gray}}>{student.form} · {student.gender}</div>
              </div>
            </div>
            {showC && (
              <div style={{borderTop:`1px solid ${C.grayLight}`,paddingTop:11}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                  <span style={{fontSize:10.5,color:C.gray}}>📅 Pre-filled from Attendance records — edit if needed</span>
                  <SmBtn onClick={()=>setConduct(c=>({...c,present:String(attTotals.present),absent:String(attTotals.absent),late:String(attTotals.late)}))} color={C.navyMid}>Reset to Attendance</SmBtn>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9}}>
                  {[["Days Present","present"],["Days Absent","absent"],["Times Late","late"]].map(([l,k])=>(
                    <div key={k}><label style={lbl}>{l}</label><input style={inp} value={conduct[k]} onChange={e=>setConduct(c=>({...c,[k]:e.target.value}))}/></div>
                  ))}
                  <div><label style={lbl}>Conduct</label>
                    <select style={inp} value={conduct.conduct} onChange={e=>setConduct(c=>({...c,conduct:e.target.value}))}>
                      {["Excellent","Very Good","Good","Satisfactory","Needs Improvement"].map(o=><option key={o}>{o}</option>)}
                    </select>
                  </div>
                  <div style={{gridColumn:"1/-1"}}><label style={lbl}>Class Teacher's Remark</label><input style={inp} value={conduct.classTeacherRemark} onChange={e=>setConduct(c=>({...c,classTeacherRemark:e.target.value}))}/></div>
                  <div style={{gridColumn:"1/-1"}}><label style={lbl}>Principal's Remark</label><input style={inp} value={conduct.principalRemark} onChange={e=>setConduct(c=>({...c,principalRemark:e.target.value}))}/></div>
                </div>
              </div>
            )}
          </div>

          {/* Preview card — SBC compact design */}
          <div style={{background:C.white,borderRadius:8,overflow:"hidden",
            boxShadow:"0 2px 10px rgba(0,0,0,0.1)",marginBottom:13,border:"2px solid #1a56a0"}}>

            {/* Govt header */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",
              padding:"5px 10px",borderBottom:"1px solid #ccc",fontSize:7.5}}>
              <div style={{lineHeight:1.5}}><strong>REPUBLIC OF CAMEROON</strong><br/>Peace-Work-Fatherland<br/>Ministry of Secondary Education</div>
              <div style={{fontSize:20}}>🇨🇲</div>
              <div style={{lineHeight:1.5,textAlign:"right"}}><strong>REPUBLIQUE DU CAMEROUN</strong><br/>Paix – Travail – Patrie<br/>Ministère De l'Enseignement</div>
            </div>

            {/* School banner */}
            <div style={{background:"#1a56a0",color:"#fff",textAlign:"center",padding:"6px 4px"}}>
              <div style={{fontSize:13,fontWeight:900,letterSpacing:.5}}>SAKER BAPTIST COLLAGE(SBC)-BAWE</div>
              <div style={{fontSize:7.5,marginTop:1,opacity:.85}}>Motto: Quality, Discipline, and Excellence education</div>
            </div>

            {/* Title */}
            <div style={{textAlign:"center",padding:"4px 0",borderBottom:"1px solid #1a56a0",fontSize:10,fontWeight:800,textDecoration:"underline"}}>
              ACADEMIC REPORT CARD / BULLETIN DE NOTES
            </div>

            {/* Term + student info + photo */}
            <div style={{display:"flex",gap:0,borderBottom:"1px solid #000"}}>
              <div style={{flex:1,fontSize:8}}>
                <div style={{padding:"2px 6px",borderBottom:"1px solid #ccc",borderRight:"1px solid #000"}}>
                  <strong>{sel.mode==="term"?sel.term.replace(" Term","").toUpperCase()+" TERM":"ANNUAL"}</strong> &nbsp;|&nbsp;
                  Class: <strong>{student.form}</strong> &nbsp;|&nbsp;
                  Roll: <strong>{students.filter(s=>s.active&&s.form===student.form).length}</strong> &nbsp;|&nbsp;
                  Year: <strong>{sel.year}</strong>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",borderBottom:"1px solid #ccc"}}>
                  <div style={{padding:"2px 6px",borderRight:"1px solid #ccc"}}>Name: <strong>{student.name}</strong></div>
                  <div style={{padding:"2px 6px"}}>Matricule: <strong>{student.id}</strong></div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",borderBottom:"1px solid #ccc"}}>
                  <div style={{padding:"2px 6px",borderRight:"1px solid #ccc"}}>Sex: <strong>{student.gender}</strong></div>
                  <div style={{padding:"2px 6px"}}>DOB: <strong>{fmtDate(student.dob)}</strong></div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr"}}>
                  <div style={{padding:"2px 6px",borderRight:"1px solid #ccc"}}>Parent: <strong>{student.parent||"—"}</strong></div>
                  <div style={{padding:"2px 6px"}}>Phone: <strong>{student.phone||"—"}</strong></div>
                </div>
              </div>
              <div style={{width:70,borderLeft:"1px solid #000",padding:3,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
                {student.photo_url
                  ?<img src={student.photo_url} alt="" style={{width:62,height:75,objectFit:"cover",border:"1px solid #1a56a0"}}/>
                  :<div style={{width:62,height:75,border:"1px solid #999",display:"flex",alignItems:"center",justifyContent:"center",fontSize:8,color:"#999",textAlign:"center"}}>PHOTO</div>
                }
              </div>
            </div>

            {/* Marks table */}
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",minWidth:480}}>
                <thead>
                  <tr style={{background:"#1a56a0",color:"#fff"}}>
                    <th style={{padding:"1px 2px",textAlign:"left",fontSize:6,border:"1px solid #1a56a0",width:"18%"}}>SUBJECT</th>
                    {sel.mode==="term"
                      ? termSeqs.map((sq)=><th key={sq} style={{padding:"1px",textAlign:"center",fontSize:6,border:"1px solid #fff",width:"9%"}}>SQ{sq.replace("SEQ ","")}</th>)
                      : ["1","2","3","4","5","6"].map(s=><th key={s} style={{padding:"1px",textAlign:"center",fontSize:5.5,border:"1px solid #fff",width:"5.2%"}}>SQ{s}</th>)
                    }
                    <th style={{padding:"1px",textAlign:"center",fontSize:6,border:"1px solid #fff",width:"8%"}}>Avg</th>
                    <th style={{padding:"1px",textAlign:"center",fontSize:6,border:"1px solid #fff",width:"7%"}}>Coeff</th>
                    <th style={{padding:"1px",textAlign:"center",fontSize:6,border:"1px solid #fff",width:"8%"}}>Score</th>
                    <th style={{padding:"1px",textAlign:"center",fontSize:6,border:"1px solid #fff",width:"7%"}}>Rmk</th>
                  </tr>
                </thead>
                <tbody>
                  {[
                    "English language","French/ Français","Mathematics","Health science",
                    "Home management","Citizenship","Food and Nutrition","Chemistry",
                    "History","Economics","Geography","Biology","Human Biology","Physics",
                    "Literature in English","Computer Studies","Religious Studies",
                    "Commerce","Accounts","Hygiene","Sport/Physical education","Manual Labour"
                  ].map((sub,i)=>{
                    const r=modeRows.find(x=>
                      x.sub===sub||
                      (sub==="French/ Français"&&x.sub==="French")||
                      (sub==="Computer Studies"&&x.sub==="Computer Science")||
                      (sub==="Sport/Physical education"&&x.sub==="Physical Education")||
                      (sub==="Literature in English"&&x.sub==="Literature")
                    );
                    const avg=r?(sel.mode==="term"?r.termAvg:r.annualAvg):null;
                    const score=avg!=null&&r?(avg*r.coeff).toFixed(1):"";
                    const {remark}=r&&avg!=null?scoreToGrade(avg):{remark:""};
                    const pass=avg!=null&&avg>=10;
                    const teacherName = r?.teacher && r.teacher!=="—" ? r.teacher : "";
                    return(
                      <tr key={sub} style={{background:i%2===0?"#fff":"#F8FAFC"}}>
                        <td style={{padding:"1px 2px",fontSize:5.8,border:"1px solid #1a56a0",lineHeight:1.05,wordBreak:"break-word"}}>
                          {sub}
                          {teacherName && <div style={{fontSize:5,color:"#1a56a0",fontStyle:"italic",lineHeight:1.05}}>{teacherName}</div>}
                        </td>
                        {sel.mode==="term"
                          ? r?.termScores.map((s,j)=><td key={j} style={{textAlign:"center",fontSize:6.5,border:"1px solid #1a56a0",padding:"1px"}}>{s!=null?s:""}</td>)
                          : [r?.s1,r?.s2,r?.s3,r?.s4,r?.s5,r?.s6].map((s,j)=><td key={j} style={{textAlign:"center",fontSize:6,border:"1px solid #1a56a0",padding:"1px"}}>{s!=null?s:""}</td>)
                        }
                        <td style={{textAlign:"center",fontSize:6.5,fontWeight:700,border:"1px solid #1a56a0",padding:"1px"}}>{avg!=null?avg.toFixed(1):""}</td>
                        <td style={{textAlign:"center",fontSize:6.5,border:"1px solid #1a56a0",padding:"1px"}}>{r?r.coeff:""}</td>
                        <td style={{textAlign:"center",fontSize:6.5,fontWeight:700,border:"1px solid #1a56a0",padding:"1px"}}>{score}</td>
                        <td style={{textAlign:"center",fontSize:5.5,color:avg!=null?(pass?C.green:C.red):"#000",border:"1px solid #1a56a0",padding:"1px"}}>{remark.slice(0,3)}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* Summary */}
            {(()=>{
              // Auto compute class ranking and class average for preview
              const peers2=students.filter(s=>s.active&&s.form===student.form&&s.reg_status==="registered");
              const peerAvgs2=peers2.map(ps=>{
                let tw=0,tc=0;
                modeRows.forEach(r=>{
                  const seqs=sel.mode==="term"
                    ? TERM_SEQS[sel.term].map(sq=>{const k=ps.id+"-"+r.sub+"-"+sq+"-"+sel.year;return marksMap[k]?.score??null;})
                    : SEQ_LABELS.map(sq=>{const k=ps.id+"-"+r.sub+"-"+sq+"-"+sel.year;return marksMap[k]?.score??null;});
                  const vals=seqs.filter(x=>x!==null);
                  if(vals.length){tw+=vals.reduce((a,b)=>a+Number(b),0)/vals.length*r.coeff;tc+=r.coeff;}
                });
                return{id:ps.id,avg:tc?tw/tc:0};
              }).sort((a,b)=>b.avg-a.avg);
              const myRank2=peerAvgs2.findIndex(x=>x.id===student.id)+1;
              const rankStr2=myRank2>0?myRank2+"/"+peers2.length:"—";
              const classAvg2=peerAvgs2.length?(peerAvgs2.reduce((a,x)=>a+x.avg,0)/peerAvgs2.length).toFixed(2):"—";
              return(
                <div style={{borderTop:"2px solid #1a56a0",padding:"5px 10px",fontSize:8.5}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:6,marginBottom:4}}>
                    <div>Terminal Avg: <strong>{overallAvg?overallAvg.toFixed(2):"—"}/20</strong></div>
                    <div>Class Avg: <strong>{classAvg2}/20</strong></div>
                    <div>Position: <strong>{rankStr2}</strong></div>
                    <div>Promoted/Repeat: ______</div>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:6,marginBottom:4}}>
                    <div>Annual Avg: <strong>{sel.mode==="annual"&&overallAvg?overallAvg.toFixed(2):"—"}/20</strong></div>
                    <div>Absences: <strong>{conduct.absent||"—"}</strong></div>
                    <div>Warnings: ______</div>
                    <div>Fees Owing: ______</div>
                  </div>
                  <div style={{borderTop:"1px solid #ccc",paddingTop:4,marginBottom:6}}>
                    Class Teacher Remark: <strong>{conduct.classTeacherRemark||"_________________________"}</strong>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginTop:8}}>
                    {["Class Teacher / Prof. Principal","Principal's Signature / Stamp","Parent / Guardian"].map((r,i)=>(
                      <div key={r} style={{textAlign:"center"}}>
                        <div style={{height:20,borderBottom:"1px solid #000",marginBottom:2}}></div>
                        <div style={{fontSize:7.5}}>{r}</div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}
          </div>

          {modeRows.length>0
            ? <button onClick={printReport} style={{width:"100%",padding:"13px",background:`linear-gradient(90deg,${C.navy},${C.navyMid})`,color:C.white,border:"none",borderRadius:10,fontWeight:800,fontSize:15,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
                🖨 &nbsp;Print A4 Report Card — {student.name}
              </button>
            : <div style={{background:"#fef3cd",borderRadius:10,padding:"12px 16px",textAlign:"center",fontSize:13,color:"#92400e",border:"1px solid #fde68a"}}>⚠️ No marks found for this period. Enter marks first.</div>
          }
        </div>
      )}
    </div>
  );
  } catch(err) {
    return <div style={{padding:20,color:"red",fontSize:13}}>
      <strong>Report Card Error:</strong> {err.message}<br/>
      <small>{err.stack}</small>
    </div>;
  }
}

// ─── Fees ──────────────────────────────────────────────────────────────────────
function FeesPage({ ctx }) {
  const { students, feesMap, saveFee, auth } = ctx;
  const [filter,    setFilter]    = useState({ form:"", search:"", status:"" });
  const [threshold, setThreshold] = useState(50);
  const [modal,     setModal]     = useState(null);
  const [amount,    setAmount]    = useState("");
  const [saving,    setSaving]    = useState(false);
  const [pendingPayments, setPendingPayments] = useState([]);
  const [loadingPending,  setLoadingPending]  = useState(true);
  const [reviewing,       setReviewing]       = useState(null); // submission being viewed full-size

  async function loadPendingPayments() {
    setLoadingPending(true);
    try {
      const { data, error } = await supabase
        .from("payment_submissions")
        .select("*, students(name, form)")
        .eq("status","pending")
        .order("submitted_at",{ascending:false});
      if (error) throw error;
      setPendingPayments(data||[]);
    } catch(e) { console.error("Failed to load pending payments:", e); }
    setLoadingPending(false);
  }

  useEffect(() => { loadPendingPayments(); }, []);

  async function approvePayment(sub) {
    if (!window.confirm(`Approve ${sub.amount.toLocaleString()} FCFA for ${sub.students?.name||sub.student_id}? This will add it to their fee balance.`)) return;
    try {
      const currentPaid = feesMap[sub.student_id]?.paid||0;
      await saveFee(sub.student_id, currentPaid + sub.amount);
      const reviewedAt = new Date().toISOString();
      await supabase.from("payment_submissions").update({
        status:"approved", reviewed_at:reviewedAt, reviewed_by:auth.user.name,
      }).eq("id", sub.id);
      await loadPendingPayments();
      setReviewing(null);
      if (window.confirm("Payment approved. Print a slip for the parent now?")) {
        printPaymentSlip(
          { ...sub, reviewed_at:reviewedAt, reviewed_by:auth.user.name },
          { id: sub.student_id, name: sub.students?.name, form: sub.students?.form }
        );
      }
    } catch(e) { alert("Error approving: "+e.message); }
  }

  async function rejectPayment(sub, reason) {
    try {
      await supabase.from("payment_submissions").update({
        status:"rejected", reviewed_at:new Date().toISOString(), reviewed_by:auth.user.name, reject_reason:reason||null,
      }).eq("id", sub.id);
      await loadPendingPayments();
      setReviewing(null);
    } catch(e) { alert("Error rejecting: "+e.message); }
  }

  const allActive       = students.filter(s=>s.active);
  const totalExpected   = allActive.length * TOTAL_FEE;
  const totalCollected  = allActive.reduce((a,s)=>a+(feesMap[s.id]?.paid||0),0);

  const filtered = allActive.filter(s => {
    const paid = feesMap[s.id]?.paid||0;
    const pct  = Math.round(paid/TOTAL_FEE*100);
    const st   = pct>=100?"Paid":pct>0?"Partial":"Unpaid";
    return (!filter.form   || s.form===filter.form)
      &&   (!filter.search || s.name?.toLowerCase().includes(filter.search.toLowerCase()))
      &&   (!filter.status || st===filter.status);
  });

  function feeExportRow(s) {
    const paid = feesMap[s.id]?.paid||0;
    const bal  = TOTAL_FEE-paid;
    const pct  = Math.round(paid/TOTAL_FEE*100);
    const st   = pct>=100?"Paid":pct>0?"Partial":"Unpaid";
    return {
      "Matricule": s.id,
      "Name":      s.name,
      "Form":      s.form,
      "Parent/Guardian": s.parent||"",
      "Phone":     s.phone||"",
      "Total Fee (FCFA)": TOTAL_FEE,
      "Paid (FCFA)":      paid,
      "Balance (FCFA)":   bal,
      "% Paid":           pct,
      "Status":           st,
    };
  }

  function exportFeesExcel() {
    if (filter.form) {
      // A single form is already selected — one sheet is enough
      const scope = allActive.filter(s => s.form===filter.form).sort((a,b)=>a.name.localeCompare(b.name));
      if (!scope.length) { alert("No students to export."); return; }
      exportToExcel(scope.map(feeExportRow), `SBC-Fees-${filter.form.replace(" ","")}-2026-2027.xlsx`, filter.form);
    } else {
      // No form filter — split into one sheet per class, all in one workbook
      if (!allActive.length) { alert("No students to export."); return; }
      const sheets = FORMS.map(f => ({
        name: f,
        rows: allActive.filter(s=>s.form===f).sort((a,b)=>a.name.localeCompare(b.name)).map(feeExportRow),
      }));
      exportToExcelMultiSheet(sheets, `SBC-Fees-AllForms-2026-2027.xlsx`);
    }
  }

  function printDebtors() {
    const debtors = allActive.filter(s => {
      const pct = Math.round((feesMap[s.id]?.paid||0)/TOTAL_FEE*100);
      return pct<threshold && (!filter.form || s.form===filter.form);
    }).sort((a,b)=>a.form.localeCompare(b.form)||a.name.localeCompare(b.name));

    if (!debtors.length) { alert("No students below this threshold."); return; }

    const byForm = {};
    debtors.forEach(s=>{ if(!byForm[s.form]) byForm[s.form]=[]; byForm[s.form].push(s); });

    const sections = Object.entries(byForm).map(([form,list]) => {
      const rows = list.map((s,i) => {
        const paid=feesMap[s.id]?.paid||0, bal=TOTAL_FEE-paid, pct=Math.round(paid/TOTAL_FEE*100);
        return `<tr style="background:${i%2===0?"#fff":"#F9FAFB"}">
          <td style="padding:5px 8px;font-size:11px">${i+1}</td>
          <td style="padding:5px 8px;font-size:11px;font-weight:700">${s.name}</td>
          <td style="padding:5px 8px;font-size:10px;font-family:monospace;color:#C9962A">${s.id}</td>
          <td style="padding:5px 8px;font-size:11px;text-align:right">${paid.toLocaleString()} F</td>
          <td style="padding:5px 8px;font-size:11px;text-align:right;font-weight:700;color:#B91C1C">${bal.toLocaleString()} F</td>
          <td style="padding:5px 8px;font-size:11px;text-align:center;color:${pct>=50?"#C9962A":"#B91C1C"};font-weight:700">${pct}%</td>
          <td style="padding:5px 8px;font-size:11px">${s.parent||""}</td>
          <td style="padding:5px 8px;font-size:11px">${s.phone||""}</td>
        </tr>`;
      }).join("");
      const fp  = list.reduce((a,s)=>a+(feesMap[s.id]?.paid||0),0);
      const fb  = list.length*TOTAL_FEE-fp;
      return `<div style="margin-bottom:14px">
        <div style="background:#0D2340;color:#fff;padding:5px 10px;font-size:11px;font-weight:800;border-radius:4px 4px 0 0">${form} — ${list.length} student(s) below ${threshold}%</div>
        <table style="width:100%;border-collapse:collapse;border:1px solid #E5E7EB">
          <thead><tr style="background:#163558;color:#fff">
            <th style="padding:5px 8px;font-size:9px;text-align:left">#</th>
            <th style="padding:5px 8px;font-size:9px;text-align:left">Student Name</th>
            <th style="padding:5px 8px;font-size:9px;text-align:left">ID</th>
            <th style="padding:5px 8px;font-size:9px;text-align:right">Paid</th>
            <th style="padding:5px 8px;font-size:9px;text-align:right">Balance</th>
            <th style="padding:5px 8px;font-size:9px;text-align:center">%</th>
            <th style="padding:5px 8px;font-size:9px;text-align:left">Parent</th>
            <th style="padding:5px 8px;font-size:9px;text-align:left">Phone</th>
          </tr></thead>
          <tbody>${rows}</tbody>
          <tfoot><tr style="background:#EEF2FF;border-top:2px solid #0D2340">
            <td colspan="3" style="padding:5px 8px;font-size:11px;font-weight:800">Subtotal — ${form}</td>
            <td style="padding:5px 8px;text-align:right;font-size:11px;font-weight:700">${fp.toLocaleString()} F</td>
            <td style="padding:5px 8px;text-align:right;font-size:11px;font-weight:800;color:#B91C1C">${fb.toLocaleString()} F</td>
            <td colspan="3"></td>
          </tr></tfoot>
        </table>
      </div>`;
    }).join("");

    const totBal = debtors.length*TOTAL_FEE - debtors.reduce((a,s)=>a+(feesMap[s.id]?.paid||0),0);
    const totPd  = debtors.reduce((a,s)=>a+(feesMap[s.id]?.paid||0),0);
    const html   = `<div style="font-family:Segoe UI,Arial,sans-serif;padding:12px;max-width:780px;margin:0 auto;color:#0D2340">
      <div style="text-align:center;margin-bottom:12px">
        <div style="font-size:22px">🎓</div>
        <h2 style="font-size:15px;font-weight:900;margin:3px 0">SAKER BAPTIST COLLEGE</h2>
        <p style="font-size:9px;color:#6B7280;margin:0">NGEPTANG · NONI · NW REGION</p>
        <div style="margin-top:7px;background:#B91C1C;display:inline-block;padding:3px 16px;border-radius:10px;color:#fff;font-size:11px;font-weight:800">SCHOOL FEES DEFAULTER LIST</div>
        <p style="font-size:9.5px;color:#374151;margin-top:5px">Below <strong>${threshold}%</strong> · Total fee: <strong>${TOTAL_FEE.toLocaleString()} FCFA</strong> · Academic Year 2026/2027 · ${new Date().toLocaleDateString("en-GB")}</p>
      </div>
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:12px;text-align:center">
        <div style="background:#EEF2FF;border-radius:6px;padding:8px;border-left:3px solid #0D2340"><div style="font-size:17px;font-weight:900">${debtors.length}</div><div style="font-size:9px;color:#6B7280">Defaulters</div></div>
        <div style="background:#FEF2F2;border-radius:6px;padding:8px;border-left:3px solid #B91C1C"><div style="font-size:17px;font-weight:900;color:#B91C1C">${totBal.toLocaleString()} F</div><div style="font-size:9px;color:#6B7280">Outstanding</div></div>
        <div style="background:#FFFBEB;border-radius:6px;padding:8px;border-left:3px solid #C9962A"><div style="font-size:17px;font-weight:900;color:#C9962A">${totPd.toLocaleString()} F</div><div style="font-size:9px;color:#6B7280">Partially Paid</div></div>
      </div>
      ${sections}
      <p style="font-size:8px;color:#9CA3AF;text-align:center;margin-top:10px">Confidential — administrative use only · Saker Baptist College</p>
    </div>`;
    domPrint("sbc-fees", html, "A4 portrait", "7mm");
  }

  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
        <div style={{background:C.white,borderRadius:10,padding:"11px 12px",borderLeft:`4px solid ${C.navy}`,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",gridColumn:"1/-1"}}>
          <div style={{fontSize:11,color:C.gray}}>Total Expected — {allActive.length} students × {TOTAL_FEE.toLocaleString()} FCFA</div>
          <div style={{fontSize:20,fontWeight:900,color:C.navy,marginTop:2}}>{totalExpected.toLocaleString()} FCFA</div>
        </div>
        {[["Collected",totalCollected,C.green],["Outstanding",totalExpected-totalCollected,C.red]].map(([l,v,col])=>(
          <div key={l} style={{background:C.white,borderRadius:10,padding:"11px 12px",borderLeft:`4px solid ${col}`,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{fontSize:11,color:C.gray}}>{l}</div>
            <div style={{fontSize:18,fontWeight:900,color:col,marginTop:2}}>{v.toLocaleString()} F</div>
          </div>
        ))}
      </div>

      {pendingPayments.length>0 && (
        <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",border:`2px solid ${C.gold}`}}>
          <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:9,display:"flex",alignItems:"center",gap:6}}>
            💳 Pending Payment Confirmations <Pill color={C.gold}>{pendingPayments.length}</Pill>
          </div>
          {pendingPayments.map(sub => (
            <div key={sub.id} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",borderBottom:`1px solid ${C.grayBg}`}}>
              {sub.screenshot && <img src={sub.screenshot} alt="" onClick={()=>setReviewing(sub)} style={{width:44,height:44,objectFit:"cover",borderRadius:6,cursor:"pointer",border:`1px solid ${C.grayLight}`,flexShrink:0}}/>}
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontWeight:700,fontSize:12,color:C.navy}}>{sub.students?.name||sub.student_id} <span style={{color:C.gray,fontWeight:400}}>({sub.students?.form})</span></div>
                <div style={{fontSize:11,color:C.gray}}>{sub.amount.toLocaleString()} FCFA · {sub.momo_number||"no number given"} · {fmtDate(sub.submitted_at?.slice(0,10))}</div>
              </div>
              <SmBtn onClick={()=>setReviewing(sub)} color={C.navyMid}>Review</SmBtn>
            </div>
          ))}
        </div>
      )}

      {/* Defaulter print panel */}
      <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:8}}>🖨 Print Defaulter List</div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:8}}>
          {[25,50,75,100].map(v=>(
            <button key={v} onClick={()=>setThreshold(v)} style={{padding:"6px 12px",borderRadius:7,border:`2px solid ${threshold===v?C.red:C.grayLight}`,background:threshold===v?"#FEF2F2":C.white,color:threshold===v?C.red:C.gray,fontWeight:700,fontSize:12,cursor:"pointer"}}>{"<"}{v}%</button>
          ))}
          <span style={{fontSize:12,color:C.gray,alignSelf:"center"}}>
            → <strong style={{color:C.red}}>{allActive.filter(s=>Math.round((feesMap[s.id]?.paid||0)/TOTAL_FEE*100)<threshold).length}</strong> student(s)
          </span>
        </div>
        <select style={{...inp,maxWidth:160,marginBottom:8}} value={filter.form} onChange={e=>setFilter(f=>({...f,form:e.target.value}))}>
          <option value="">All Forms</option>{FORMS.map(f=><option key={f}>{f}</option>)}
        </select>
        <button onClick={printDebtors} style={{width:"100%",padding:"10px",background:C.red,color:C.white,border:"none",borderRadius:8,fontWeight:800,fontSize:13,cursor:"pointer"}}>
          🖨 Print Defaulters — Below {threshold}%{filter.form?` (${filter.form})`:""}
        </button>
        <button onClick={exportFeesExcel} style={{width:"100%",padding:"10px",marginTop:8,background:C.green,color:C.white,border:"none",borderRadius:8,fontWeight:800,fontSize:13,cursor:"pointer"}}>
          📊 Export Full Fees Report — Excel {filter.form?`(${filter.form})`:"(All Forms)"}
        </button>
      </div>

      {/* Filters */}
      <div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>
        <input style={{...inp,flex:1,minWidth:120}} placeholder="Search…" value={filter.search} onChange={e=>setFilter(f=>({...f,search:e.target.value}))}/>
        <select style={{...inp,width:106}} value={filter.form} onChange={e=>setFilter(f=>({...f,form:e.target.value}))}><option value="">All Forms</option>{FORMS.map(f=><option key={f}>{f}</option>)}</select>
        <select style={{...inp,width:94}} value={filter.status} onChange={e=>setFilter(f=>({...f,status:e.target.value}))}><option value="">All</option><option>Paid</option><option>Partial</option><option>Unpaid</option></select>
      </div>

      {filtered.map(s => {
        const paid=feesMap[s.id]?.paid||0, bal=TOTAL_FEE-paid;
        const pct=Math.round(paid/TOTAL_FEE*100);
        const st=pct>=100?"Paid":pct>0?"Partial":"Unpaid";
        const sc=st==="Paid"?C.green:st==="Partial"?C.gold:C.red;
        return (
          <div key={s.id} style={{background:C.white,borderRadius:10,padding:"11px 12px",marginBottom:7,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div><div style={{fontWeight:700,color:C.navy}}>{s.name}</div><div style={{fontSize:11,color:C.gray}}>{s.id} · {s.form}</div></div>
              <div style={{textAlign:"right"}}><Pill color={sc}>{st}</Pill><div style={{fontSize:10,color:C.gray,marginTop:2}}>{pct}%</div></div>
            </div>
            <div style={{marginTop:7,background:C.grayBg,borderRadius:5,height:7,overflow:"hidden"}}>
              <div style={{width:`${pct}%`,height:"100%",background:sc,borderRadius:5,transition:"width .4s"}}/>
            </div>
            <div style={{display:"flex",justifyContent:"space-between",marginTop:3,fontSize:12}}>
              <span style={{color:C.gray}}>Paid: <strong>{paid.toLocaleString()}</strong> / {TOTAL_FEE.toLocaleString()} F</span>
              <span style={{color:bal>0?C.red:C.green}}>Bal: <strong>{bal.toLocaleString()} F</strong></span>
            </div>
            <SmBtn onClick={()=>{setModal(s);setAmount("");}} color={C.navy} style={{marginTop:7}}>Record Payment</SmBtn>
          </div>
        );
      })}
      {!filtered.length && <Empty text="No students match this filter."/>}

      {modal && (
        <Modal title={`Record Payment — ${modal.name}`} onClose={()=>setModal(null)}>
          <div style={{background:C.grayBg,borderRadius:8,padding:"9px 12px",marginBottom:11}}>
            {[
              ["Total Fee",   `${TOTAL_FEE.toLocaleString()} FCFA`,         C.navy],
              ["Paid",        `${(feesMap[modal.id]?.paid||0).toLocaleString()} FCFA`,  C.green],
              ["Balance",     `${(TOTAL_FEE-(feesMap[modal.id]?.paid||0)).toLocaleString()} FCFA`, C.red],
            ].map(([l,v,col])=>(
              <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:`1px solid ${C.grayLight}`}}>
                <span style={{fontSize:12,color:C.gray}}>{l}</span>
                <strong style={{fontSize:13,color:col}}>{v}</strong>
              </div>
            ))}
          </div>
          <Fr label="Payment Amount (FCFA)">
            <input style={inp} type="number" value={amount} onChange={e=>setAmount(e.target.value)} placeholder={`Max ${(TOTAL_FEE-(feesMap[modal.id]?.paid||0)).toLocaleString()} F`}/>
          </Fr>
          <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:11}}>
            <Btn onClick={()=>setModal(null)} outline>Cancel</Btn>
            <Btn onClick={async()=>{
              const pay=Math.min(Number(amount)||0,TOTAL_FEE-(feesMap[modal.id]?.paid||0));
              if(pay<=0) return;
              setSaving(true);
              try { await saveFee(modal.id,(feesMap[modal.id]?.paid||0)+pay); setModal(null); }
              catch(e){ alert("Error: "+e.message); }
              setSaving(false);
            }} disabled={saving}>{saving?"Saving…":"Record Payment"}</Btn>
          </div>
        </Modal>
      )}

      {reviewing && <PaymentReviewModal submission={reviewing} onApprove={approvePayment} onReject={rejectPayment} onClose={()=>setReviewing(null)}/>}
    </div>
  );
}

// ─── Payment Screenshot Review Modal ─────────────────────────────────────────────
function PaymentReviewModal({ submission, onApprove, onReject, onClose }) {
  const [showReject, setShowReject] = useState(false);
  const [reason, setReason] = useState("");

  return (
    <Modal title={`Review Payment — ${submission.students?.name||submission.student_id}`} onClose={onClose}>
      {submission.screenshot && (
        <img src={submission.screenshot} alt="Payment screenshot" style={{width:"100%",borderRadius:8,marginBottom:12,border:`1px solid ${C.grayLight}`}}/>
      )}
      <div style={{background:C.grayBg,borderRadius:8,padding:"10px 12px",marginBottom:12}}>
        {[
          ["Student", `${submission.students?.name||submission.student_id} (${submission.students?.form||""})`],
          ["Amount", `${submission.amount.toLocaleString()} FCFA`],
          ["Phone Used", submission.momo_number||"Not provided"],
          ["Submitted", fmtDate(submission.submitted_at?.slice(0,10))],
        ].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:`1px solid ${C.grayLight}`}}>
            <span style={{fontSize:12,color:C.gray}}>{l}</span>
            <strong style={{fontSize:12,color:C.navy}}>{v}</strong>
          </div>
        ))}
      </div>

      {!showReject ? (
        <div style={{display:"flex",gap:8}}>
          <Btn onClick={()=>setShowReject(true)} outline color={C.red}>Reject</Btn>
          <Btn onClick={()=>onApprove(submission)} color={C.green}>✓ Approve & Record Payment</Btn>
        </div>
      ) : (
        <div>
          <Fr label="Reason for rejection (optional)">
            <input style={inp} placeholder="e.g. Screenshot unclear, amount mismatch…" value={reason} onChange={e=>setReason(e.target.value)}/>
          </Fr>
          <div style={{display:"flex",gap:8}}>
            <Btn onClick={()=>setShowReject(false)} outline>Cancel</Btn>
            <Btn onClick={()=>onReject(submission, reason)} color={C.red}>Confirm Reject</Btn>
          </div>
        </div>
      )}
    </Modal>
  );
}

// ─── Notices ───────────────────────────────────────────────────────────────────
function NoticesPage({ ctx }) {
  const { notices, saveNotice, deleteNotice, auth } = ctx;
  const [modal, setModal] = useState(null);
  const [form,  setForm]  = useState({ title:"", body:"" });

  async function save() {
    if (!form.title.trim()) return;
    await saveNotice({ ...form, id:modal==="add"?null:modal.id, author:auth.user.name, posted_date:form.posted_date||todayStr() });
    setModal(null);
  }

  return (
    <div>
      {auth.role==="admin" && (
        <div style={{display:"flex",justifyContent:"flex-end",marginBottom:10}}>
          <Btn onClick={()=>{setForm({title:"",body:""});setModal("add");}}>+ Post Notice</Btn>
        </div>
      )}
      {notices.map(n => (
        <div key={n.id} style={{background:C.white,borderRadius:10,padding:13,marginBottom:9,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",borderLeft:`4px solid ${C.gold}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
            <h3 style={{margin:0,fontSize:13,color:C.navy}}>{n.title}</h3>
            {auth.role==="admin" && (
              <div style={{display:"flex",gap:5,flexShrink:0}}>
                <SmBtn onClick={()=>{setForm({title:n.title,body:n.body,posted_date:n.posted_date});setModal(n);}} color={C.green}>Edit</SmBtn>
                <SmBtn onClick={()=>{if(confirm("Delete this notice?")) deleteNotice(n.id);}} color={C.red}>Del</SmBtn>
              </div>
            )}
          </div>
          <p style={{margin:"7px 0 0",fontSize:12,color:C.gray,lineHeight:1.6}}>{n.body}</p>
          <div style={{marginTop:5,fontSize:10,color:C.gold,fontWeight:600}}>{fmtDate(n.posted_date)} — {n.author}</div>
        </div>
      ))}
      {!notices.length && <Empty text="No notices posted."/>}

      {modal && (
        <Modal title={modal==="add"?"Post Notice":"Edit Notice"} onClose={()=>setModal(null)}>
          <Fr label="Title"><input style={inp} value={form.title} onChange={e=>setForm(f=>({...f,title:e.target.value}))}/></Fr>
          <Fr label="Message"><textarea style={{...inp,minHeight:85,resize:"vertical"}} value={form.body} onChange={e=>setForm(f=>({...f,body:e.target.value}))}/></Fr>
          <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:12}}>
            <Btn onClick={()=>setModal(null)} outline>Cancel</Btn>
            <Btn onClick={save}>{modal==="add"?"Post":"Save"}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ─── Calendar ──────────────────────────────────────────────────────────────────
const EVENT_CATEGORIES = {
  exam:     { label:"Exam",     color:C.red,    icon:"📝" },
  holiday:  { label:"Holiday",  color:C.green,  icon:"🏖️" },
  meeting:  { label:"Meeting",  color:C.navyMid,icon:"👥" },
  deadline: { label:"Deadline", color:C.gold,   icon:"⏰" },
  other:    { label:"Other",    color:C.gray,   icon:"📌" },
};
const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];

function CalendarPage({ ctx }) {
  const { calendarEvents, saveCalendarEvent, deleteCalendarEvent, auth } = ctx;
  const [viewDate, setViewDate] = useState(new Date());
  const [jumpToDate, setJumpToDate] = useState(null); // when set, highlights that date's events in the list below
  const [filterCat, setFilterCat] = useState("");
  const [modal, setModal] = useState(null);
  const blank = { title:"", description:"", event_date:todayStr(), end_date:"", category:"other", forms:[] };
  const [form, setForm] = useState(blank);
  const [saving, setSaving] = useState(false);

  const year = viewDate.getFullYear();
  const month = viewDate.getMonth();
  const firstOfMonth = new Date(year, month, 1);
  const startWeekday = firstOfMonth.getDay(); // 0=Sun
  const daysInMonth = new Date(year, month+1, 0).getDate();

  function dateStr(y,m,d) { return `${y}-${String(m+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`; }

  const eventsByDate = {};
  calendarEvents.forEach(ev => {
    if (filterCat && ev.category!==filterCat) return;
    const start = ev.event_date;
    const end = ev.end_date || ev.event_date;
    // Mark every day in a multi-day range, not just the start
    let cur = new Date(start+"T00:00:00");
    const endD = new Date(end+"T00:00:00");
    while (cur <= endD) {
      const key = cur.toISOString().slice(0,10);
      if (!eventsByDate[key]) eventsByDate[key]=[];
      eventsByDate[key].push(ev);
      cur.setDate(cur.getDate()+1);
    }
  });

  const upcoming = calendarEvents
    .filter(ev => {
      if (filterCat && ev.category!==filterCat) return false;
      const isUpcoming = (ev.end_date||ev.event_date) >= todayStr();
      const matchesJump = jumpToDate && ev.event_date<=jumpToDate && (ev.end_date||ev.event_date)>=jumpToDate;
      return isUpcoming || matchesJump;
    })
    .sort((a,b)=>a.event_date.localeCompare(b.event_date));

  // Scroll to and briefly highlight the tapped day's event(s) in the list
  const eventRefs = useRef({});
  useEffect(() => {
    if (!jumpToDate) return;
    const firstMatch = upcoming.find(ev => ev.event_date<=jumpToDate && (ev.end_date||ev.event_date)>=jumpToDate);
    if (firstMatch && eventRefs.current[firstMatch.id]) {
      eventRefs.current[firstMatch.id].scrollIntoView({ behavior:"smooth", block:"center" });
    }
    const t = setTimeout(()=>setJumpToDate(null), 2000); // clear the highlight after a moment
    return () => clearTimeout(t);
  }, [jumpToDate]);

  async function save() {
    if (!form.title?.trim()||!form.event_date) return;
    setSaving(true);
    try { await saveCalendarEvent(form); setModal(null); }
    catch(e) { alert("Error: "+e.message); }
    setSaving(false);
  }

  const cells = [];
  for (let i=0;i<startWeekday;i++) cells.push(null);
  for (let d=1;d<=daysInMonth;d++) cells.push(d);

  return (
    <div>
      {/* Month navigator */}
      <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <button onClick={()=>setViewDate(new Date(year,month-1,1))} style={{background:C.grayBg,border:"none",borderRadius:8,width:32,height:32,fontSize:16,cursor:"pointer",color:C.navy}}>‹</button>
          <div style={{fontWeight:800,color:C.navy,fontSize:14}}>{MONTH_NAMES[month]} {year}</div>
          <button onClick={()=>setViewDate(new Date(year,month+1,1))} style={{background:C.grayBg,border:"none",borderRadius:8,width:32,height:32,fontSize:16,cursor:"pointer",color:C.navy}}>›</button>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2,marginBottom:4}}>
          {["S","M","T","W","T","F","S"].map((d,i)=><div key={i} style={{textAlign:"center",fontSize:10,color:C.gray,fontWeight:700}}>{d}</div>)}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2}}>
          {cells.map((d,i) => {
            if (d===null) return <div key={i}/>;
            const key = dateStr(year,month,d);
            const dayEvents = eventsByDate[key]||[];
            const isToday = key===todayStr();
            const hasEvents = dayEvents.length>0;
            const dominantColor = hasEvents ? (EVENT_CATEGORIES[dayEvents[0].category]?.color||C.gray) : null;
            const uniqueCats = [...new Set(dayEvents.map(ev=>ev.category))];

            // Background priority: today always shows solid navy so it never
            // gets lost, but if today ALSO has an event, a colored ring makes
            // that visible too. A non-today day with events gets a tinted
            // background in that event's category color, so it's obvious
            // at a glance which dates have something on them.
            let bg = "transparent";
            let border = "2px solid transparent";
            if (isToday) {
              bg = C.navy;
              if (hasEvents) border = `2px solid ${dominantColor}`;
            } else if (hasEvents) {
              bg = dominantColor + "28"; // soft tint, not a solid block
              border = `1.5px solid ${dominantColor}60`;
            }

            return (
              <button
                key={i}
                onClick={()=>hasEvents && setJumpToDate(key)}
                style={{
                  aspectRatio:"1", display:"flex", flexDirection:"column",
                  alignItems:"center", justifyContent:"center",
                  borderRadius:7, background:bg, border, boxSizing:"border-box",
                  color:isToday?C.white:(hasEvents?dominantColor:C.navy),
                  fontSize:11, fontWeight:(isToday||hasEvents)?800:500,
                  position:"relative", cursor:hasEvents?"pointer":"default",
                  padding:0,
                }}
              >
                {d}
                {hasEvents && (
                  <div style={{display:"flex",gap:1.5,marginTop:1,position:"absolute",bottom:2}}>
                    {uniqueCats.slice(0,3).map((cat,j)=><div key={j} style={{width:4,height:4,borderRadius:"50%",background:isToday?C.white:EVENT_CATEGORIES[cat]?.color||C.gray}}/>)}
                  </div>
                )}
              </button>
            );
          })}
        </div>
      </div>

      {/* Category filter */}
      <div style={{display:"flex",gap:6,marginBottom:11,flexWrap:"wrap"}}>
        <button onClick={()=>setFilterCat("")} style={{padding:"5px 10px",borderRadius:14,border:`1.5px solid ${!filterCat?C.navy:C.grayLight}`,background:!filterCat?C.navy:C.white,color:!filterCat?C.white:C.gray,fontSize:11,fontWeight:700,cursor:"pointer"}}>All</button>
        {Object.entries(EVENT_CATEGORIES).map(([k,c])=>(
          <button key={k} onClick={()=>setFilterCat(k)} style={{padding:"5px 10px",borderRadius:14,border:`1.5px solid ${filterCat===k?c.color:C.grayLight}`,background:filterCat===k?c.color:C.white,color:filterCat===k?C.white:C.gray,fontSize:11,fontWeight:700,cursor:"pointer"}}>{c.icon} {c.label}</button>
        ))}
      </div>

      {auth.role==="admin" && (
        <div style={{display:"flex",justifyContent:"flex-end",marginBottom:11}}>
          <Btn onClick={()=>{setForm(blank);setModal("add");}}>+ Add Event</Btn>
        </div>
      )}

      {/* Upcoming list */}
      <div style={{fontSize:12,fontWeight:700,color:C.gray,marginBottom:7}}>Upcoming</div>
      <div style={{background:C.white,borderRadius:10,overflow:"hidden",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        {!upcoming.length
          ? <Empty text="No upcoming events."/>
          : upcoming.map((ev,i) => {
              const cat = EVENT_CATEGORIES[ev.category]||EVENT_CATEGORIES.other;
              const isRange = ev.end_date && ev.end_date!==ev.event_date;
              const isJumped = jumpToDate && ev.event_date<=jumpToDate && (ev.end_date||ev.event_date)>=jumpToDate;
              return (
                <div
                  key={ev.id}
                  ref={el => { if(el) eventRefs.current[ev.id]=el; }}
                  style={{
                    padding:"10px 12px", borderBottom:`1px solid ${C.grayBg}`,
                    background:isJumped?cat.color+"22":(i%2===0?C.white:C.grayBg),
                    borderLeft:`4px solid ${cat.color}`,
                    transition:"background 0.6s ease",
                  }}
                >
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontWeight:700,color:C.navy,fontSize:13}}>{cat.icon} {ev.title}</div>
                      <div style={{fontSize:11,color:C.gray,marginTop:2}}>
                        {fmtDate(ev.event_date)}{isRange?` – ${fmtDate(ev.end_date)}`:""}
                        {ev.forms?.length>0 && <span> · {ev.forms.join(", ")}</span>}
                      </div>
                      {ev.description && <div style={{fontSize:11,color:C.gray,marginTop:3}}>{ev.description}</div>}
                    </div>
                    {auth.role==="admin" && (
                      <div style={{display:"flex",gap:5,flexShrink:0}}>
                        <SmBtn onClick={()=>{setForm({...ev});setModal(ev);}} color={C.green}>Edit</SmBtn>
                        <SmBtn onClick={()=>{if(window.confirm(`Delete "${ev.title}"?`))deleteCalendarEvent(ev.id);}} color={C.red}>Del</SmBtn>
                      </div>
                    )}
                  </div>
                </div>
              );
            })
        }
      </div>

      {modal && (
        <Modal title={modal==="add"?"Add Calendar Event":"Edit Event"} onClose={()=>setModal(null)}>
          <Fr label="Title"><input style={inp} value={form.title} onChange={e=>setForm(f=>({...f,title:e.target.value}))}/></Fr>
          <Fr label="Category">
            <select style={inp} value={form.category} onChange={e=>setForm(f=>({...f,category:e.target.value}))}>
              {Object.entries(EVENT_CATEGORIES).map(([k,c])=><option key={k} value={k}>{c.icon} {c.label}</option>)}
            </select>
          </Fr>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9}}>
            <Fr label="Start Date"><input style={inp} type="date" value={form.event_date} onChange={e=>setForm(f=>({...f,event_date:e.target.value}))}/></Fr>
            <Fr label="End Date (optional)"><input style={inp} type="date" value={form.end_date||""} onChange={e=>setForm(f=>({...f,end_date:e.target.value}))}/></Fr>
          </div>
          <Fr label="Description (optional)"><textarea style={{...inp,minHeight:60,resize:"vertical"}} value={form.description||""} onChange={e=>setForm(f=>({...f,description:e.target.value}))}/></Fr>
          <div style={{marginTop:4,marginBottom:8}}>
            <label style={lbl}>Applies to (leave blank for whole school)</label>
            <div style={{display:"flex",flexWrap:"wrap",gap:5,marginTop:5}}>
              {FORMS.map(f => (
                <button key={f} onClick={()=>setForm(prev=>({...prev,forms:prev.forms?.includes(f)?prev.forms.filter(x=>x!==f):[...(prev.forms||[]),f]}))} style={{padding:"4px 8px",borderRadius:12,border:`1px solid ${form.forms?.includes(f)?C.gold:C.grayLight}`,background:form.forms?.includes(f)?C.gold:C.white,color:form.forms?.includes(f)?C.white:C.gray,fontSize:11,cursor:"pointer"}}>{f}</button>
              ))}
            </div>
          </div>
          <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:12}}>
            <Btn onClick={()=>setModal(null)} outline>Cancel</Btn>
            <Btn onClick={save} disabled={saving}>{saving?"Saving…":modal==="add"?"Add Event":"Save"}</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ─── Profile ───────────────────────────────────────────────────────────────────
function ProfilePage({ ctx }) {
  const { auth, teachers } = ctx;
  const u = auth.user;
  const t = teachers.find(x=>x.email===u.email);

  const [showPwForm, setShowPwForm] = useState(false);
  const [pass1, setPass1] = useState("");
  const [pass2, setPass2] = useState("");
  const [pwErr, setPwErr] = useState("");
  const [pwOk,  setPwOk]  = useState("");
  const [pwBusy,setPwBusy]= useState(false);

  async function changePassword() {
    setPwErr(""); setPwOk("");
    if (pass1.length < 6) { setPwErr("Password must be at least 6 characters."); return; }
    if (pass1 !== pass2)  { setPwErr("Passwords do not match."); return; }
    setPwBusy(true);
    try {
      const { error } = await supabase.auth.updateUser({ password: pass1 });
      if (error) throw error;
      setPwOk("Password updated successfully.");
      setPass1(""); setPass2("");
      setTimeout(()=>setShowPwForm(false), 1500);
    } catch(e) { setPwErr("Could not update password: " + e.message); }
    setPwBusy(false);
  }

  return (
    <div style={{maxWidth:480,margin:"0 auto"}}>
      <div style={{background:C.white,borderRadius:13,padding:22,boxShadow:"0 1px 4px rgba(0,0,0,0.06)",marginBottom:12}}>
        <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:18}}>
          <div style={{width:56,height:56,borderRadius:"50%",background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,flexShrink:0}}>
            {auth.role==="admin"?"🏫":"👩‍🏫"}
          </div>
          <div>
            <div style={{fontWeight:800,fontSize:16,color:C.navy}}>{u.name}</div>
            <Pill color={C.gold}>{auth.role==="admin"?"Administrator":"Teacher"}</Pill>
          </div>
        </div>
        {[
          ["Email", u.email],
          ...( t ? [
            ["Subjects", (t.subjects||[]).join(", ")||"—"],
            ["Forms",    (t.forms||[]).join(", ")||"—"],
            ["Joined",   fmtDate(t.joined)],
          ] : [] ),
        ].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${C.grayBg}`}}>
            <span style={{fontSize:12,color:C.gray}}>{l}</span>
            <span style={{fontSize:12,fontWeight:600,color:C.navy,maxWidth:"60%",textAlign:"right"}}>{v||"—"}</span>
          </div>
        ))}
      </div>

      <div style={{background:C.white,borderRadius:13,padding:22,boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:showPwForm?14:0}}>
          <div style={{fontWeight:800,fontSize:14,color:C.navy}}>🔑 Password</div>
          <SmBtn onClick={()=>{setShowPwForm(v=>!v);setPwErr("");setPwOk("");}} color={C.navyMid}>{showPwForm?"Cancel":"Change Password"}</SmBtn>
        </div>
        {showPwForm && (
          <div>
            <Fr label="New Password">
              <input style={inp} type="password" autoComplete="new-password" placeholder="At least 6 characters" value={pass1} onChange={e=>setPass1(e.target.value)} disabled={pwBusy}/>
            </Fr>
            <Fr label="Confirm New Password">
              <input style={inp} type="password" autoComplete="new-password" placeholder="Re-type password" value={pass2} onChange={e=>setPass2(e.target.value)} onKeyDown={e=>e.key==="Enter"&&changePassword()} disabled={pwBusy}/>
            </Fr>
            {pwErr && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"8px 12px",color:C.red,fontSize:12,marginBottom:10}}>{pwErr}</div>}
            {pwOk  && <div style={{background:"#f0fdf4",border:"1px solid #86efac",borderRadius:8,padding:"8px 12px",color:C.green,fontSize:12,marginBottom:10}}>{pwOk}</div>}
            <Btn onClick={changePassword} disabled={pwBusy}>{pwBusy?"Updating…":"Update Password"}</Btn>
          </div>
        )}
      </div>
    </div>
  );
}