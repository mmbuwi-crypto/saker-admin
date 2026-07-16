// ─────────────────────────────────────────────────────────────────────────────
//  SAKER BAPTIST COLLEGE — Main App (Supabase backend)
// ─────────────────────────────────────────────────────────────────────────────
import { useState, useEffect, useRef, useCallback } from "react";
import { supabase } from "./supabase.js";

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
      const MAX = 200;
      let w = img.width, h = img.height;
      const scale = Math.min(MAX/w, MAX/h, 1);
      w = Math.round(w*scale); h = Math.round(h*scale);
      const canvas = document.createElement("canvas");
      canvas.width = w; canvas.height = h;
      canvas.getContext("2d").drawImage(img, 0, 0, w, h);
      resolve(canvas.toDataURL("image/jpeg", 0.6));
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

// ─── Nav ──────────────────────────────────────────────────────────────────────
const navItems = [
  { key:"dashboard",    label:"Dashboard",    icon:"📊", roles:["admin","teacher"] },
  { key:"registration", label:"Registration", icon:"📋", roles:["admin"] },
  { key:"students",     label:"Students",     icon:"🎒", roles:["admin","teacher"] },
  { key:"teachers",     label:"Teachers",     icon:"👩‍🏫", roles:["admin"] },
  { key:"marks",        label:"Enter Marks",  icon:"✏️", roles:["admin","teacher"] },
  { key:"reports",      label:"Report Cards", icon:"📄", roles:["admin","teacher"] },
  { key:"fees",         label:"School Fees",  icon:"💰", roles:["admin"] },
  { key:"notices",      label:"Notices",      icon:"📢", roles:["admin","teacher"] },
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
  const [session,    setSession]    = useState(undefined); // undefined=loading
  const [userRole,   setUserRole]   = useState(null);
  const [userProfile,setUserProfile]= useState(null);
  const [page,       setPage]       = useState("dashboard");
  const [menuOpen,   setMenuOpen]   = useState(false);

  // Data state
  const [students,   setStudents]   = useState([]);
  const [teachers,   setTeachers]   = useState([]);
  const [marksMap,   setMarksMap]   = useState({}); // key → mark record
  const [feesMap,    setFeesMap]    = useState({}); // studentId → {paid,total}
  const [notices,    setNotices]    = useState([]);
  const [loading,    setLoading]    = useState(true);
  const [dbError,    setDbError]    = useState(null);

  // ── Auth ───────────────────────────────────────────────────────────────────
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, session) => setSession(session));
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
        { data: fe }, { data: nt }
      ] = await Promise.all([
        supabase.from("students").select("*").order("form").order("name"),
        supabase.from("teachers").select("*").order("name"),
        supabase.from("marks").select("*"),
        supabase.from("fees").select("*"),
        supabase.from("notices").select("*").order("posted_date", { ascending: false }),
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
      setNotices(nt || []);
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

  // ── DB write helpers (passed via ctx) ──────────────────────────────────────
  async function saveStudent(s) {
    // photo_url is already compressed base64 or null — save directly
    let photoUrl = s.photo_url || null;
    const row = {
      id: s.id, name: s.name, form: s.form, gender: s.gender,
      dob: s.dob||null, parent: s.parent, phone: s.phone, address: s.address,
      photo_url: photoUrl||null, active: s.active!==false,
      reg_status: s.reg_status||"pending", reg_date: s.reg_date||null,
      reg_fee: s.reg_fee||null, reg_receipt: s.reg_receipt||null,
      reg_paid_by: s.reg_paid_by||null, reg_cashier: s.reg_cashier||null,
      is_late_reg: s.is_late_reg||false,
    };
    const { error } = await supabase.from("students").upsert(row, { onConflict:"id" });
    if (error) throw error;
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

  // ── Current user info ───────────────────────────────────────────────────────
  const currentTeacher = teachers.find(t => t.email === session?.user?.email);
  const currentUser = {
    id:    session?.user?.id || "ADMIN",
    name:  currentTeacher?.name || userProfile?.name || session?.user?.email || "Admin",
    email: session?.user?.email || "",
  };

  // ── Loading / not-logged-in screens ────────────────────────────────────────
  if (session === undefined) return (
    <div style={{minHeight:"100vh",background:C.navy,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:14}}>
      <div style={{fontSize:48}}>🎓</div>
      <div style={{color:C.white,fontWeight:800,fontSize:18}}>Saker Baptist College</div>
      <div style={{color:C.goldLight,fontSize:13}}>Starting up…</div>
    </div>
  );
  if (!session) return <LoginScreen onLogin={doLogin} />;

  const ctx = {
    auth: { user: currentUser, role: userRole||"teacher" },
    students, teachers, marksMap, feesMap, notices,
    saveStudent, deleteStudent, saveTeacher, saveMark,
    saveFee, saveNotice, deleteNotice, loadAll,
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
          {page==="reports"      && <ReportsPage      ctx={ctx}/>}
          {page==="fees"         && ctx.auth.role==="admin" && <FeesPage        ctx={ctx}/>}
          {page==="notices"      && <NoticesPage      ctx={ctx}/>}
          {page==="profile"      && <ProfilePage      ctx={ctx}/>}
        </>}
      </div>
    </div>
  );
}

// ─── Login Screen ──────────────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [email, setEmail] = useState("");
  const [pass,  setPass]  = useState("");
  const [err,   setErr]   = useState("");
  const [busy,  setBusy]  = useState(false);

  async function handle() {
    setErr(""); setBusy(true);
    try { await onLogin(email, pass); }
    catch(e) { setErr(e.message.includes("Invalid login") ? "Incorrect email or password." : "Login error: "+e.message); }
    setBusy(false);
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
            <input style={inp} type="email" placeholder="your@email.com" value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handle()}/>
          </Fr>
          <Fr label="Password">
            <input style={inp} type="password" placeholder="••••••••" value={pass} onChange={e=>setPass(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handle()}/>
          </Fr>
          {err && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"8px 12px",color:C.red,fontSize:12,marginBottom:12}}>{err}</div>}
          <Btn onClick={handle} disabled={busy}>{busy?"Signing in…":"Sign In →"}</Btn>
          <p style={{textAlign:"center",fontSize:10,color:C.gray,marginTop:14,marginBottom:0}}>Contact your administrator if you forgot your password.</p>
        </div>
      </div>
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

  const fee = form.isLate ? REG_LATE : REG_NORMAL;

  async function handleRegister() {
    if (!form.name.trim()||!form.parent.trim()||!form.phone.trim()) return;
    setSaving(true);
    const fNum = form.form.replace("Form ","");
    const n    = students.filter(s=>s.form===form.form).length + 1;
    const id   = `SBC0${fNum.padStart(2,"0")}${String(n).padStart(3,"0")}`;
    const rec  = `RCP-${new Date().getFullYear()}-${String(Date.now()).slice(-4)}`;
    try {
      // Compress photo then save student
      let photoUrl = null;
      if (form.photo_url) {
        photoUrl = await compressPhoto(form.photo_url);
      }
      const studentData = {
        name:form.name, form:form.form, gender:form.gender,
        dob:form.dob||null, parent:form.parent, phone:form.phone,
        address:form.address||null, photo_url:photoUrl,
        id, active:true,
        reg_status:"registered", reg_date:todayStr(),
        reg_fee:fee, reg_receipt:rec,
        reg_paid_by:form.paidBy||form.parent,
        reg_cashier:auth.user.name,
        is_late_reg:form.isLate,
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
        {[["register","➕ Register"],["list","📋 List"],["receipt","🧾 Receipt"]].map(([k,l]) => (
          <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"9px 4px",borderRadius:7,border:"none",cursor:"pointer",fontWeight:700,fontSize:12,background:tab===k?C.navy:"transparent",color:tab===k?C.white:C.gray}}>{l}</button>
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
                <label style={{display:"block",marginTop:4,fontSize:10,color:C.navyMid,cursor:"pointer",fontWeight:700,background:C.grayBg,borderRadius:4,padding:"3px 5px",textAlign:"center"}}>
                  📷 {form.photo_url?"Change":"Upload"}
                  <input type="file" accept="image/*" style={{display:"none"}} onChange={e=>{
                    const f=e.target.files[0]; if(!f) return;
                    const r=new FileReader();
                    r.onload=ev=>setForm(f=>({...f,photo_url:ev.target.result}));
                    r.readAsDataURL(f);
                  }}/>
                </label>
                {form.photo_url && <button onClick={()=>setForm(f=>({...f,photo_url:null}))} style={{marginTop:2,fontSize:9,color:C.red,background:"none",border:"none",cursor:"pointer"}}>✕ Remove</button>}
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

// ─── Students ──────────────────────────────────────────────────────────────────
function StudentsPage({ ctx }) {
  const { students, saveStudent, deleteStudent, auth } = ctx;
  const [filter, setFilter] = useState({ form:"", search:"" });
  const [modal,  setModal]  = useState(null);
  const [viewS,  setViewS]  = useState(null);
  const [saving, setSaving] = useState(false);
  const blank = { name:"", form:"Form 1", gender:"Female", dob:"", parent:"", phone:"", address:"", photo_url:null };
  const [form, setForm] = useState(blank);

  const filtered = students.filter(s => s.active
    && (!filter.form   || s.form===filter.form)
    && (!filter.search || s.name?.toLowerCase().includes(filter.search.toLowerCase()) || s.id?.includes(filter.search))
  );

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

  return (
    <div>
      <div style={{display:"flex",gap:8,marginBottom:11,flexWrap:"wrap"}}>
        <input style={{...inp,flex:1,minWidth:130}} placeholder="Search name or ID…" value={filter.search} onChange={e=>setFilter(f=>({...f,search:e.target.value}))}/>
        <select style={{...inp,width:114}} value={filter.form} onChange={e=>setFilter(f=>({...f,form:e.target.value}))}>
          <option value="">All Forms</option>{FORMS.map(f=><option key={f}>{f}</option>)}
        </select>
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
        </Modal>
      )}

      {modal && (
        <Modal title={modal==="add"?"Add Student":"Edit Student"} onClose={()=>setModal(null)}>
          <div style={{display:"flex",gap:11,alignItems:"flex-start",marginBottom:10}}>
            <div style={{flexShrink:0}}>
              <PhotoBox photo={form.photo_url} size={[62,76]}/>
              <label style={{display:"block",marginTop:3,fontSize:9,color:C.navyMid,cursor:"pointer",fontWeight:700,background:C.grayBg,borderRadius:4,padding:"2px 5px",textAlign:"center"}}>
                📷 Photo
                <input type="file" accept="image/*" style={{display:"none"}} onChange={e=>{
                  const f=e.target.files[0]; if(!f) return;
                  const r=new FileReader();
                  r.onload=ev=>setForm(f=>({...f,photo_url:ev.target.result}));
                  r.readAsDataURL(f);
                }}/>
              </label>
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
    </div>
  );
}

// ─── Teachers ──────────────────────────────────────────────────────────────────
function TeachersPage({ ctx }) {
  const { teachers, saveTeacher } = ctx;
  const [modal,  setModal]  = useState(null);
  const [form,   setForm]   = useState({ name:"", email:"", subjects:[], forms:[], active:true });
  const [saving, setSaving] = useState(false);

  const toggle = (k,v) => setForm(f => ({ ...f, [k]: f[k].includes(v) ? f[k].filter(x=>x!==v) : [...f[k],v] }));

  async function save() {
    if (!form.name?.trim()||!form.email?.trim()) return;
    setSaving(true);
    const id = modal==="add" ? "TCH"+Date.now().toString().slice(-6) : modal.id;
    try { await saveTeacher({ ...form, id, joined: form.joined||todayStr() }); setModal(null); }
    catch(e) { alert("Error: "+e.message); }
    setSaving(false);
  }

  return (
    <div>
      <div style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:9,padding:"9px 12px",marginBottom:11,fontSize:12,color:"#92400e"}}>
        ℹ️ After creating a teacher here, go to <strong>Supabase → Authentication → Add User</strong> with the same email to activate their login. Then add their role in the <strong>users</strong> table.
      </div>
      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:9}}>
        <Btn onClick={()=>{setForm({name:"",email:"",subjects:[],forms:[],active:true,joined:todayStr()});setModal("add");}}>+ Create Teacher Portal</Btn>
      </div>
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
          <div style={{display:"flex",gap:7,marginTop:8}}>
            <SmBtn onClick={()=>{setForm({...t});setModal(t);}} color={C.green}>Edit</SmBtn>
            <SmBtn onClick={()=>saveTeacher({...t,active:!t.active})} color={t.active?C.red:C.green}>{t.active?"Deactivate":"Activate"}</SmBtn>
          </div>
        </div>
      ))}

      {modal && (
        <Modal title={modal==="add"?"Create Teacher Portal":"Edit Teacher"} onClose={()=>setModal(null)}>
          <Fr label="Full Name"><input style={inp} value={form.name||""} onChange={e=>setForm(f=>({...f,name:e.target.value}))}/></Fr>
          <Fr label="Email (used for login)"><input style={inp} type="email" value={form.email||""} onChange={e=>setForm(f=>({...f,email:e.target.value}))}/></Fr>
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
          <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:13}}>
            <Btn onClick={()=>setModal(null)} outline>Cancel</Btn>
            <Btn onClick={save} disabled={saving}>{saving?"Saving…":modal==="add"?"Create":"Save"}</Btn>
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
function ReportsPage({ ctx }) {
  const SBC_SUBJECTS_LIST = [
    "English language","French/ Français","Mathematics","Health science",
    "Home management","Citizenship","Food and Nutrition","Chemistry",
    "History","Economics","Geography","Biology","Human Biology","Physics",
    "Literature in English","Computer Studies","Religious Studies",
    "Commerce","Accounts","Hygiene","Sport/Physical education","Manual Labour"
  ];
  const { students, marksMap, teachers } = ctx;
  const [sel,     setSel]     = useState({ studentId:"", year:"2026/2027", mode:"term", term:"First Term" });
  const [conduct, setConduct] = useState({ present:"", absent:"", late:"", conduct:"Very Good", classTeacherRemark:"", principalRemark:"" });
  const [showC,   setShowC]   = useState(false);

  const student  = students.find(s=>s.id===sel.studentId);
  const subjects = student ? (SUBJECTS_BY_FORM[student.form]||[]) : [];
  const termSeqs = TERM_SEQS[sel.term];

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
      ?`<img src="${student.photo_url}" style="width:65px;height:78px;object-fit:cover;border:1px solid #1a56a0;display:block">`
      :`<div style="width:65px;height:78px;border:1px solid #999;display:flex;align-items:center;justify-content:center;font-size:9px;color:#999;text-align:center">PHOTO</div>`;

    const termLabel = sel.mode==="term" ? sel.term.replace(" Term","").toUpperCase()+" TERM" : "ANNUAL";
    const seqNums = sel.mode==="term"
      ? termSeqs.map(sq=>sq.replace("SEQ ",""))
      : ["1","2","3","4","5","6"];
    const seqHeaders = seqNums.map(n=>`<th style="background:#1a56a0;color:#fff;border:1px solid #fff;padding:2px 1px;font-size:7px;width:${sel.mode==="term"?"7%":"4.5%"}">SQ${n}</th>`).join("");

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
          <td style="padding:1px 3px;font-size:7.5px;border:1px solid #1a56a0">${sub}</td>
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
        ? r.termScores.map(s=>`<td style="text-align:center;font-size:8px;border:1px solid #1a56a0">${s!=null?s:""}</td>`).join("")
        : SEQ_LABELS.map((_,j)=>{
            const sc=[r.s1,r.s2,r.s3,r.s4,r.s5,r.s6][j];
            return`<td style="text-align:center;font-size:7.5px;border:1px solid #1a56a0">${sc!=null?sc:""}</td>`;
          }).join("");
      return`<tr style="background:${i%2===0?"#fff":"#F9FAFB"}">
        <td style="padding:1px 3px;font-size:7.5px;border:1px solid #1a56a0">${sub}</td>
        ${scoreCols}
        <td style="text-align:center;font-size:8px;font-weight:700;border:1px solid #1a56a0">${avg!=null?avg.toFixed(1):""}</td>
        <td style="text-align:center;font-size:8px;border:1px solid #1a56a0">${r.coeff}</td>
        <td style="text-align:center;font-size:8px;font-weight:700;border:1px solid #1a56a0">${pond}</td>
        <td style="text-align:center;font-size:7px;color:${pass?"#15803D":"#B91C1C"};border:1px solid #1a56a0">${remark}</td>
      <td style="padding:1px 3px;font-size:6.5px;color:#163558;border:1px solid #1a56a0">${r.teacher||""}</td>
      </tr>`;
    }).join("");

    return`<!DOCTYPE html><html><head><meta charset="utf-8">
<title>Report Card — ${student.name}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  @page{size:A4 portrait;margin:6mm}
  html,body{font-family:Arial,Helvetica,sans-serif;color:#000;background:#fff;
    -webkit-print-color-adjust:exact;print-color-adjust:exact}
  table{width:100%;border-collapse:collapse}
  .page{width:198mm;margin:0 auto}
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
      No. on Roll: <strong>${peers.length}</strong> &nbsp;|&nbsp; Effective: <strong>${students.filter(s=>s.active&&s.form===student.form&&s.reg_status==="registered").length}</strong> &nbsp;|&nbsp;
      Academic Year: <strong>${sel.year}</strong>
    </td>
    <td rowspan="4" style="width:72px;padding:3px;border-left:1px solid #000;vertical-align:top;text-align:center">
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
      <th style="text-align:left;padding:2px 4px;font-size:7.5px;border:1px solid #fff;width:20%">SUBJECT</th>
      ${seqHeaders}
      <th style="padding:2px 2px;font-size:7px;border:1px solid #fff;width:6%">Avg/20</th>
      <th style="padding:2px 2px;font-size:7px;border:1px solid #fff;width:4%">Cf</th>
      <th style="padding:2px 2px;font-size:7px;border:1px solid #fff;width:6%">Score</th>
      <th style="padding:2px 2px;font-size:7px;border:1px solid #fff;width:9%">Remarks</th>
      <th style="padding:2px 2px;font-size:7px;border:1px solid #fff;width:14%">Teacher</th>
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
    domPrint("sbc-report", buildPrintHTML(), "A4 portrait", "7mm");
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
                    <th style={{padding:"2px 4px",textAlign:"left",fontSize:7.5,border:"1px solid #1a56a0",width:"22%"}}>SUBJECT</th>
                    {sel.mode==="term"
                      ? termSeqs.map((sq)=><th key={sq} style={{padding:"2px 1px",textAlign:"center",fontSize:7,border:"1px solid #fff",width:"8%"}}>SQ{sq.replace("SEQ ","")}</th>)
                      : ["SQ1","SQ2","SQ3","SQ4","SQ5","SQ6"].map(s=><th key={s} style={{padding:"2px 1px",textAlign:"center",fontSize:6.5,border:"1px solid #fff",width:"5%"}}>{s}</th>)
                    }
                    <th style={{padding:"2px 1px",textAlign:"center",fontSize:7,border:"1px solid #fff",width:"7%"}}>Avg</th>
                    <th style={{padding:"2px 1px",textAlign:"center",fontSize:7,border:"1px solid #fff",width:"5%"}}>Cf</th>
                    <th style={{padding:"2px 1px",textAlign:"center",fontSize:7,border:"1px solid #fff",width:"7%"}}>Score</th>
                    <th style={{padding:"2px 1px",textAlign:"center",fontSize:7,border:"1px solid #fff",width:"10%"}}>Rmk</th>
                    <th style={{padding:"2px 1px",textAlign:"center",fontSize:7,border:"1px solid #fff",width:"16%"}}>Teacher</th>
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
                    return(
                      <tr key={sub} style={{background:i%2===0?"#fff":"#F8FAFC"}}>
                        <td style={{padding:"1px 3px",fontSize:7.5,border:"1px solid #1a56a0"}}>{sub}</td>
                        {sel.mode==="term"
                          ? r?.termScores.map((s,j)=><td key={j} style={{textAlign:"center",fontSize:7.5,border:"1px solid #1a56a0",padding:"1px"}}>{s!=null?s:""}</td>)
                          : [r?.s1,r?.s2,r?.s3,r?.s4,r?.s5,r?.s6].map((s,j)=><td key={j} style={{textAlign:"center",fontSize:7,border:"1px solid #1a56a0",padding:"1px"}}>{s!=null?s:""}</td>)
                        }
                        <td style={{textAlign:"center",fontSize:7.5,fontWeight:700,border:"1px solid #1a56a0",padding:"1px"}}>{avg!=null?avg.toFixed(1):""}</td>
                        <td style={{textAlign:"center",fontSize:7.5,border:"1px solid #1a56a0",padding:"1px"}}>{r?r.coeff:""}</td>
                        <td style={{textAlign:"center",fontSize:7.5,fontWeight:700,border:"1px solid #1a56a0",padding:"1px"}}>{score}</td>
                        <td style={{textAlign:"center",fontSize:7,color:avg!=null?(pass?C.green:C.red):"#000",border:"1px solid #1a56a0"}}>{remark}</td>
                        <td style={{padding:"1px 2px",fontSize:6.5,color:"#163558",border:"1px solid #1a56a0"}}>{r?.teacher||""}</td>
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
  const { students, feesMap, saveFee } = ctx;
  const [filter,    setFilter]    = useState({ form:"", search:"", status:"" });
  const [threshold, setThreshold] = useState(50);
  const [modal,     setModal]     = useState(null);
  const [amount,    setAmount]    = useState("");
  const [saving,    setSaving]    = useState(false);

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
    </div>
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

// ─── Profile ───────────────────────────────────────────────────────────────────
function ProfilePage({ ctx }) {
  const { auth, teachers } = ctx;
  const u = auth.user;
  const t = teachers.find(x=>x.email===u.email);
  return (
    <div style={{maxWidth:480,margin:"0 auto"}}>
      <div style={{background:C.white,borderRadius:13,padding:22,boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
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
    </div>
  );
}