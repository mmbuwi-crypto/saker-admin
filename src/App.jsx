// ─────────────────────────────────────────────────────────────────────────────
//  SAKER BAPTIST COMPREHENSIVE COLLEGE BAWE — Main App (Supabase backend)
// ─────────────────────────────────────────────────────────────────────────────
import { useState, useEffect, useRef, useCallback, useMemo } from "react";
import { supabase, SUPABASE_URL } from "./supabase.js";
import * as XLSX from "xlsx";
import { SCHOOL_LOGO } from "./schoolLogo.js";

// ─── Palette ──────────────────────────────────────────────────────────────────
const C = {
  navy:"#0D2340", navyMid:"#163558", gold:"#C9962A", goldLight:"#F0C050",
  white:"#FFFFFF", red:"#B91C1C", green:"#15803D",
  gray:"#6B7280", grayLight:"#E5E7EB", grayBg:"#F3F4F6", border:"#D1D5DB",
};

// ─── Constants ────────────────────────────────────────────────────────────────
const FORMS = ["Form 1","Form 2","Form 3","Form 4","Form 5"];
// One master list: every subject printed on the report card. Every form, the teacher setup,
// Form 5 student subjects and the marks pages all draw from it, so they always match.
const ALL_SUBJECTS = [
  "English language","French/ Français","Mathematics","Health science",
  "Home management","Citizenship","Food and Nutrition","Chemistry",
  "History","Economics","Geography","Biology","Human Biology","Physics",
  "Literature in English","Computer Studies","Religious Studies",
  "Commerce","Accounts","Hygiene","Sport/Physical education","Manual Labour"
];
const SUBJECTS_BY_FORM = {
  "Form 1":[...ALL_SUBJECTS], "Form 2":[...ALL_SUBJECTS], "Form 3":[...ALL_SUBJECTS],
  "Form 4":[...ALL_SUBJECTS], "Form 5":[...ALL_SUBJECTS],
};
const SEQ_LABELS  = ["SEQ 1","SEQ 2","SEQ 3","SEQ 4","SEQ 5","SEQ 6"];
const TERM_SEQS   = {"First Term":["SEQ 1","SEQ 2"],"Second Term":["SEQ 3","SEQ 4"],"Third Term":["SEQ 5","SEQ 6"]};
const TERMS       = ["First Term","Second Term","Third Term"];
// Form 5 (exam class) is not assessed by sequences. It has three assessments instead.
const F5_ASSESSMENTS = ["Pre Mock","Regional Mock","GCE Result"];
const F5_MOCKS       = ["Pre Mock","Regional Mock"];      // scored /20, these make the Form 5 average
const GCE_SEQ        = "GCE Result";                      // entered and shown as a grade, not a score
const GCE_GRADES     = ["A","B","C","D","E","U"];         // stored in the score column as 5,4,3,2,1,0
const gceToPoints    = g => { const i = GCE_GRADES.indexOf(g); return i<0 ? "" : 5-i; };
const gceFromPoints  = v => (v===null||v===undefined||v==="") ? "" : (GCE_GRADES[5-Number(v)] || "");
// Form 5 setup, kept in app_settings (no database change needed):
//   f5_grade_scale → {A:16,B:14,C:10,D:8,E:6}  minimum score /20 for each grade (below E = U), set by admin
//   f5_students    → {studentId:{subjects:[...], candNo:"..."}}  subjects each student writes + candidate number
const F5_DEFAULT_SCALE = { A:16, B:14, C:10, D:8, E:6 };
const parseJSON = (v, fb) => { try { let o = typeof v==="string" ? JSON.parse(v) : v; if (typeof o==="string") o = JSON.parse(o); return o && typeof o==="object" ? o : fb; } catch { return fb; } };
// Form 5 students shown on the admin tools: active ones, plus graduated ones who have marks in the chosen year
function f5PoolFor(students, marksMap, year) {
  const has = new Set();
  Object.values(marksMap||{}).forEach(m => { if (m.acad_year===year) has.add(m.student_id); });
  return students.filter(x => x.form==="Form 5" && x.reg_status==="registered" && (x.active || has.has(x.id))).sort((a,b)=>a.name.localeCompare(b.name));
}
const f5Scale    = settings => ({ ...F5_DEFAULT_SCALE, ...parseJSON(settings?.f5_grade_scale, {}) });
function f5Grade(settings, score) {
  if (score===null || score===undefined || score==="") return { grade:"—", pass:false, remark:"—" };
  const sc = f5Scale(settings), v = Number(score);
  const grade = v>=sc.A?"A" : v>=sc.B?"B" : v>=sc.C?"C" : v>=sc.D?"D" : v>=sc.E?"E" : "U";
  const pass = ["A","B","C"].includes(grade);
  return { grade, pass, remark: pass ? "Pass" : "Fail" };
}
const f5Students   = settings => parseJSON(settings?.f5_students, {});
// A Form 5 student with no list saved yet is treated as writing every Form 5 subject
const f5SubjectsOf = (settings, st) => { const all = SUBJECTS_BY_FORM["Form 5"]; const c = f5Students(settings)[st.id]; return c?.subjects?.length ? all.filter(x => c.subjects.includes(x)) : all; };
const f5Writes     = (settings, st, sub) => !isF5(st.form) || f5SubjectsOf(settings, st).includes(sub);
const f5CandNo     = (settings, st) => f5Students(settings)[st.id]?.candNo || "";
// Parents choose their child's subjects in the Parent Portal until a closing date set by the admin.
// Settings: f5_subject_deadline (YYYY-MM-DD), f5_subject_closed ("true" = closed by admin),
//           f5_parent_submitted → {studentId: date the parent saved}.  The database enforces the same rule.
const f5SubjectWindow = settings => {
  const dl = String(settings?.f5_subject_deadline||"").trim();
  if (/^(true|1|yes)$/i.test(String(settings?.f5_subject_closed||"").trim())) return { open:false, reason:"closed_by_admin", deadline:dl };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dl)) return { open:false, reason:"not_opened", deadline:"" };
  if (localToday() > dl) return { open:false, reason:"deadline_passed", deadline:dl };
  return { open:true, reason:"", deadline:dl };
};
const f5ParentSubmitted = settings => parseJSON(settings?.f5_parent_submitted, {});
// Form 5 mocks are written in papers. Each paper is stored as its own marks row ("Pre Mock P1" …);
// the combined score out of 20 is stored in the normal "Pre Mock" row so ranking, locks and progress keep working.
//   Every subject:    P1 MCQ /50 = 40%,  P2 Structural /100 = 60%
//   Computer Studies: P1 MCQ /50 = 30%,  P2 Structural /100 = 40%,  P3 Practical /100 = 30%   (change max here if P3 is out of something else)
const F5_PAPERS_STD  = [ {k:"P1",max:50,w:40}, {k:"P2",max:100,w:60} ];
const F5_PAPERS_COMP = [ {k:"P1",max:50,w:30}, {k:"P2",max:100,w:40}, {k:"P3",max:100,w:30} ];
const isComputerSub  = sub => /^computer/i.test(sub||"");
const f5PapersFor    = sub => isComputerSub(sub) ? F5_PAPERS_COMP : F5_PAPERS_STD;
const isPaperSeq     = q => /^(Pre Mock|Regional Mock) P[123]$/.test(q||"");
const f5PaperNote    = sub => f5PapersFor(sub).map(p => `Paper ${p.k.slice(1)} /${p.max} (${p.w}%)`).join(" + ");
function f5Combine(sub, vals) { // → score out of 20, or null until every paper is in
  let t = 0;
  for (const p of f5PapersFor(sub)) {
    const x = vals?.[p.k];
    if (x===""||x===null||x===undefined||!Number.isFinite(Number(x))) return null;
    t += Number(x)/p.max*p.w;
  }
  return Math.round(t/5*100)/100;
}
const gcePassed      = v => v!==null && v!==undefined && v!=="" && Number(v)>=3; // A, B, C
const F5_SHORT       = { "Pre Mock":"PRE MOCK", "Regional Mock":"REG. MOCK", "GCE Result":"GCE RESULT" };
const isF5           = form => form === "Form 5";
const seqsForForm    = form => isF5(form) ? F5_ASSESSMENTS : SEQ_LABELS;
const ALL_SEQ_LABELS = [...SEQ_LABELS, ...F5_ASSESSMENTS];
const seqOrder       = q => ALL_SEQ_LABELS.indexOf(q);
const ACAD_YEARS  = ["2026/2027","2027/2028","2028/2029"];
// The year the school is working in. Set from app_settings.current_year after load; changed by the New Academic Year tool.
let ACTIVE_YEAR = ACAD_YEARS[0];
const TOTAL_FEE   = 35000;
const FEE_CASHIER = "Rev. Kimbung Christopher Mbuwi"; // name printed as cashier on fee receipts
const REG_NORMAL  = 500;
const REG_LATE    = 1000;
let   LATE_CUTOFF = "2026-09-30"; // overridden by app_settings.late_cutoff (set by the New Academic Year tool)
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

// Match a logged-in user to their teacher profile (ignores capital letters and stray spaces in the email)
const normEmail = e => String(e || "").trim().toLowerCase();
// Subjects a teacher teaches in ONE class. Uses the per-class assignments; teachers not converted yet
// fall back to the old rule (every subject in every ticked class).
function teacherSubjectsIn(t, form) {
  if (!t) return [];
  if (t.assignments && Object.keys(t.assignments).length) return t.assignments[form] || [];
  return (t.forms||[]).includes(form) ? (t.subjects||[]) : [];
}
function teacherAssignments(t) {
  if (t?.assignments && Object.keys(t.assignments).length) return { ...t.assignments };
  return Object.fromEntries((t?.forms||[]).map(f => [f, [...(t.subjects||[])]]));
}
const findTeacher = (teachers, email) => teachers.find(t => normEmail(t.email) === normEmail(email));

// ─── Helpers ──────────────────────────────────────────────────────────────────
const todayStr = () => new Date().toISOString().slice(0,10);
const fmtDate  = d => { if(!d) return "—"; const [y,m,day]=d.split("-"); return `${day}/${m}/${y}`; };
const isLate   = d => d && d > LATE_CUTOFF;
// Next free matricule for a form = highest number already used with that form's prefix + 1.
// (Counting the students currently in the form re-used IDs once students were promoted out of it.)
function nextMatricule(form, ids, taken=0) {
  const prefix = `SBCC0${form.replace("Form ","").padStart(2,"0")}`;
  let max = 0;
  (ids||[]).forEach(id => { if (typeof id==="string" && id.startsWith(prefix)) { const n = parseInt(id.slice(prefix.length),10); if (Number.isFinite(n) && n>max) max = n; } });
  return `${prefix}${String(max+1+taken).padStart(3,"0")}`;
}
const localToday = () => new Date().toLocaleDateString("en-CA", { timeZone:"Africa/Douala" }); // YYYY-MM-DD in Cameroon time

// Is this class's marks entry closed for teachers? (deadline passed and the admin hasn't unlocked the class)
function lockStatus(form, seq, year, deadlines, unlocks) {
  const deadline = (deadlines||[]).find(d=>d.acad_year===year && d.seq===seq)?.deadline || null;
  const unlocked = (unlocks||[]).some(u=>u.form===form && u.acad_year===year && u.seq===seq);
  const passed   = !!deadline && localToday() > deadline;
  return { deadline, unlocked, passed, locked: passed && !unlocked };
}

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
  return new Promise((resolve, reject) => {
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
    img.onerror = () => reject(new Error("That file couldn't be loaded as an image."));
    img.src = base64;
  });
}
async function uploadPhoto(studentId, base64) {
  return await compressPhoto(base64);
}

// ─── Offline Queue ────────────────────────────────────────────────────────────
// Registrations and fee payments made with no internet are saved here, in the
// browser's localStorage, and pushed to Supabase automatically once the
// connection returns. This survives closing the app entirely — the queue is
// still there next time it opens, even after a phone restart.
const OFFLINE_QUEUE_KEY = "sccb_offline_queue_v1";

function generateTempId(form) {
  // Guaranteed-unique across different phones with no coordination: combines
  // the exact millisecond timestamp with a random suffix. Two phones would
  // need to register a student in the same form in the exact same
  // millisecond AND roll the same random suffix to collide — effectively
  // impossible. Clearly marked TEMP- so it's obviously not a real matricule
  // anywhere it's displayed, until sync assigns the true SBCC0... ID.
  const fNum = form.replace("Form ","");
  const rand = Math.random().toString(36).slice(2,7).toUpperCase();
  return `TEMP-${fNum}-${Date.now()}-${rand}`;
}

function loadOfflineQueue() {
  try {
    const raw = localStorage.getItem(OFFLINE_QUEUE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch(e) {
    console.error("Failed to read offline queue:", e);
    return [];
  }
}

function saveOfflineQueue(queue) {
  try {
    localStorage.setItem(OFFLINE_QUEUE_KEY, JSON.stringify(queue));
  } catch(e) {
    console.error("Failed to save offline queue:", e);
  }
}

function addToOfflineQueue(item) {
  const queue = loadOfflineQueue();
  queue.push({ ...item, uid: CURRENT_UID, queuedAt: new Date().toISOString(), id: generateTempId("Form 0").replace("TEMP-0","QITEM") });
  saveOfflineQueue(queue);
  return queue;
}

// ── Marks written offline: queued on the phone, uploaded when the connection returns ──
const markKey = r => `${r.student_id}-${r.subject}-${r.seq}-${r.acad_year}`;
async function pushMarkRows(rows, oldScores, who) {
  const { error } = await supabase.from("marks").upsert(rows, { onConflict:"student_id,subject,seq,acad_year" });
  if (error) throw error;
}
// A dead connection can leave a request hanging for minutes. After 12 s treat it as offline and queue the marks
// (safe: the upload is an upsert, so a late duplicate just writes the same values again).
const withTimeout = (p, ms) => Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error("network timeout: failed to fetch")), ms))]);
async function writeMarkRows(rows, online, who, oldScores) {
  if (!rows.length) return { queued:false };
  const queue = () => { addToOfflineQueue({ type:"marks", data:{ rows, who, oldScores } }); return { queued:true }; };
  if (!online) return queue();
  try { await withTimeout(pushMarkRows(rows, oldScores, who), 12000); return { queued:false }; }
  catch(e) { if (/fetch|network|failed to send|load failed/i.test(e?.message||"")) return queue(); throw e; }
}
// Last data loaded online, kept on the phone so the app can open and take marks without internet (photos left out to save space)
// Kept in IndexedDB, not localStorage: a whole school's marks is far bigger than localStorage's ~5 MB limit,
// which made the save fail silently and offline sign-in then said "not set up". One copy per signed-in user,
// so two teachers sharing a phone never see each other's class lists.
const SNAP_KEY = "sbcc_snapshot_v1";   // old localStorage copy, still read as a fallback
const SNAP_DB = "sbcc_offline_db", SNAP_STORE = "snapshots";
function idbOpen() {
  return new Promise((res, rej) => {
    if (!window.indexedDB) return rej(new Error("IndexedDB not available"));
    const r = indexedDB.open(SNAP_DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(SNAP_STORE);
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function idbGet(key) {
  const db = await idbOpen();
  return new Promise((res, rej) => { const q = db.transaction(SNAP_STORE).objectStore(SNAP_STORE).get(key); q.onsuccess = () => res(q.result ?? null); q.onerror = () => rej(q.error); });
}
async function idbSet(key, val) {
  const db = await idbOpen();
  return new Promise((res, rej) => { const tx = db.transaction(SNAP_STORE, "readwrite"); tx.objectStore(SNAP_STORE).put(val, key); tx.oncomplete = () => res(true); tx.onerror = () => rej(tx.error); tx.onabort = () => rej(tx.error); });
}
async function writeSnapshot(uid, o) {
  if (!uid) return false;
  try { await idbSet("snap_"+uid, { ...o, savedAt:Date.now() }); return true; }
  catch(e) { console.warn("Snapshot not saved:", e); return false; }
}
async function readSnapshot(uid) {
  try {
    const v = uid ? await idbGet("snap_"+uid) : null;
    if (v) return v;
    const r = localStorage.getItem(SNAP_KEY); return r ? JSON.parse(r) : null; // older phones: copy saved before this update
  } catch { return null; }
}

// ── Offline sign-in for teachers ─────────────────────────────────────────────
// After a teacher signs in ONLINE on this phone (with "Allow offline sign-in" ticked), a salted PBKDF2 check of the
// password is kept on the phone, never the password itself. With no internet the teacher signs in against that check,
// opens the saved class lists and enters marks. Marks go to the offline queue and upload once a real sign-in happens.
// Teachers only. Valid 30 days after the last online sign-in. 5 wrong tries lock it for 10 minutes.
const OFFLINE_LOGIN_KEY = "sbcc_offline_login_v1";
const OFFLINE_LOGIN_DAYS = 30;
const OFFLINE_LOGIN_DEFAULT_ON = true;   // the "Allow offline sign-in" box starts ticked (set false to start unticked)
let OFFLINE_PW_MEM = null;               // password of the current offline session, kept in memory only, so the app can sign in for real when the internet returns
let CURRENT_UID = null;                  // who is signed in now (stamped on queued items)
const offlineErr = msg => { const e = new Error(msg); e.friendly = true; return e; };
const hexOf = u8 => Array.from(u8).map(b => b.toString(16).padStart(2,"0")).join("");
const emailKey = e => String(e||"").trim().toLowerCase();
async function pbkdf2Hex(password, saltHex, iter) {
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveBits"]);
  const salt = Uint8Array.from(saltHex.match(/../g).map(h => parseInt(h,16)));
  return hexOf(new Uint8Array(await crypto.subtle.deriveBits({ name:"PBKDF2", hash:"SHA-256", salt, iterations:iter }, key, 256)));
}
function readOfflineLogins() { try { return JSON.parse(localStorage.getItem(OFFLINE_LOGIN_KEY)||"{}") || {}; } catch { return {}; } }
function writeOfflineLogins(o) { try { localStorage.setItem(OFFLINE_LOGIN_KEY, JSON.stringify(o)); } catch(e) { console.warn("Offline sign-in not saved:", e); } }
function removeOfflineLogin(email) { const all = readOfflineLogins(); delete all[emailKey(email)]; writeOfflineLogins(all); }
async function saveOfflineLogin(email, password, user) {
  if (!(window.crypto && window.crypto.subtle)) return;
  const salt = hexOf(crypto.getRandomValues(new Uint8Array(16))), iter = 150000;
  const hash = await pbkdf2Hex(password, salt, iter);
  const all = readOfflineLogins();
  all[emailKey(email)] = { id:user.id, email:emailKey(email), salt, iter, hash, savedAt:Date.now(), fails:0, lockUntil:0 };
  writeOfflineLogins(all);
}
async function checkOfflineLogin(email, password) {
  if (!(window.crypto && window.crypto.subtle)) throw offlineErr("Offline sign-in is not available in this browser. Connect to the internet to sign in.");
  const all = readOfflineLogins(), rec = all[emailKey(email)];
  let cachedRole = null;
  if (rec) { try { cachedRole = JSON.parse(localStorage.getItem("sbcc_profile_"+rec.id)||"null")?.role; } catch {} }
  if (!rec || cachedRole!=="teacher" || !(await readSnapshot(rec.id))) throw offlineErr("Offline sign-in is not set up for this account on this phone. Connect to the internet and sign in once, with Allow offline sign-in ticked.");
  if (Date.now()-rec.savedAt > OFFLINE_LOGIN_DAYS*86400000) throw offlineErr(`Offline sign-in has expired (${OFFLINE_LOGIN_DAYS} days). Connect to the internet and sign in again.`);
  if (rec.lockUntil > Date.now()) throw offlineErr(`Too many wrong tries. Try again in ${Math.ceil((rec.lockUntil-Date.now())/60000)} minute(s).`);
  const hash = await pbkdf2Hex(password, rec.salt, rec.iter);
  let diff = hash.length ^ rec.hash.length;
  for (let i=0; i<hash.length; i++) diff |= hash.charCodeAt(i) ^ (rec.hash.charCodeAt(i)||0);
  if (diff!==0) {
    rec.fails = (rec.fails||0)+1;
    if (rec.fails>=5) { rec.fails = 0; rec.lockUntil = Date.now()+10*60000; }
    writeOfflineLogins(all);
    throw new Error("Invalid login credentials");
  }
  rec.fails = 0; rec.lockUntil = 0; writeOfflineLogins(all);
  return rec;
}
// Marks typed on this phone but not uploaded yet: shown on screen straight away
function overlayQueuedMarks(map) {
  loadOfflineQueue().filter(q => q.type==="marks").forEach(q => (q.data?.rows||[]).forEach(r => { map[markKey(r)] = { ...(map[markKey(r)]||{}), ...r }; }));
  return map;
}

function removeFromOfflineQueue(queueItemId) {
  const queue = loadOfflineQueue().filter(q => q.id !== queueItemId);
  saveOfflineQueue(queue);
  return queue;
}

// Real connectivity check — navigator.onLine alone is unreliable (it can
// report "online" while connected to a WiFi network with no actual working
// internet). This does a tiny, cheap, no-cache fetch to confirm a real
// response comes back, with a short timeout so it never hangs the UI.
// mode "no-cors": we only care that the server answered, not what it said, so a missing CORS header can never
// make a working connection look offline.
async function checkRealConnectivity(ms = 4000) {
  if (!navigator.onLine) return false;
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), ms);
  try {
    await fetch(`${SUPABASE_URL}/auth/v1/health`, { method:"GET", mode:"no-cors", signal: controller.signal, cache:"no-store" });
    return true;
  } catch {
    return false;
  } finally {
    clearTimeout(timeout);
  }
}

// Print helper - works on mobile and desktop
function domPrint(id, html, size="A4 portrait", margin="8mm") {
  const isFull = /^\s*<!DOCTYPE/i.test(html); // a finished page (e.g. the report card) is used exactly as given
  const fullHtml = isFull ? html : `<!DOCTYPE html><html><head><meta charset="utf-8">
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
  let inner = html, extraCss = "";
  if (isFull) {
    const d = new DOMParser().parseFromString(html, "text/html");
    inner = d.body.innerHTML;
    extraCss = [...d.querySelectorAll("style")].map(s=>s.textContent).join("\n");
  }
  const st = document.createElement("style");
  st.id = id+"-s";
  st.textContent = `@media print{body>*:not(#${id}){display:none!important}#${id}{display:block!important;position:fixed;inset:0;background:#fff;z-index:99999}@page{size:${size};margin:${margin}}}#${id}{display:none}`;
  if (extraCss) st.textContent += "\n@media print{" + extraCss + "}";
  document.head.appendChild(st);
  const div = document.createElement("div");
  div.id = id; div.innerHTML = inner;
  document.body.appendChild(div);
  setTimeout(()=>{
    window.print();
    setTimeout(()=>{ document.getElementById(id)?.remove(); document.getElementById(id+"-s")?.remove(); }, 2500);
  }, 300);
}

// Export an array of row-objects to a downloadable .xlsx file.
// rows: array of plain objects (keys become column headers)
// filename: e.g. "SBCC-Students-2026-2027.xlsx"
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
        <img src="${SCHOOL_LOGO}" alt="SBCC" style="width:54px;height:54px;object-fit:contain">
        <h2 style="font-size:14px;font-weight:900;margin:4px 0">SAKER BAPTIST COMPREHENSIVE COLLEGE BAWE</h2>
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
      <div style="text-align:center;font-size:9px;color:#9ca3af;border-top:1px dashed #e5e7eb;padding:8px 14px">This slip confirms a payment toward school fees for Academic Year ${ACTIVE_YEAR}.<br>Keep for your records.</div>
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
  { key:"markctl",      label:"Marks Control",icon:"🔒", roles:["admin"] },
  { key:"remarks",      label:"Class Remarks",icon:"🗒️", roles:["admin","teacher"] },
  { key:"attendance",   label:"Attendance",   icon:"🗓️", roles:["admin","teacher"] },
  { key:"reports",      label:"Report Cards", icon:"📄", roles:["admin"] },
  { key:"performance",  label:"Performance",  icon:"📈", roles:["admin","teacher"] },
  { key:"subjres",      label:"Subject Results", icon:"🎯", roles:["admin","teacher"] },
  { key:"council",      label:"Class Council",icon:"🏛️", roles:["admin"] },
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
  const [roleLoading, setRoleLoading] = useState(false);
  const [roleError, setRoleError] = useState("");
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
  const [councilMeetings,  setCouncilMeetings]  = useState([]); // class council meetings
  const [councilDecisions, setCouncilDecisions] = useState([]); // decisions taken at those meetings
  const [councilReady,     setCouncilReady]     = useState(true); // false until the council tables exist in Supabase
  const [notices,    setNotices]    = useState([]);
  const [settings,   setSettings]   = useState({});   // app_settings key → value (e.g. next_term_begins)
  const [remarksMap, setRemarksMap] = useState({}); // "studentId|year|term" → {conduct, remark}
  const [markDeadlines, setMarkDeadlines] = useState([]); // {acad_year, seq, deadline}
  const [markUnlocks,   setMarkUnlocks]   = useState([]); // {form, acad_year, seq}
  const [loading,    setLoading]    = useState(true);
  const [dbError,    setDbError]    = useState(null);
  const [offlineQueue, setOfflineQueue] = useState(() => loadOfflineQueue());
  const [isOnline,     setIsOnline]     = useState(navigator.onLine);
  const [syncing,      setSyncing]      = useState(false);
  const [syncResult,   setSyncResult]   = useState(null); // {success, failed} shown briefly after a sync run
  const [showQueueDetails, setShowQueueDetails] = useState(false);
  const sessionRef = useRef(undefined);   // always the latest session (event handlers read it)
  sessionRef.current = session;
  CURRENT_UID = session?.user?.id || null;
  const offlineWantRef = useRef(null);    // {email, password} to remember for offline sign-in once the profile shows a teacher
  const upgradingRef = useRef(false);

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
    let cancelled = false;
    if (!session?.user?.id) {
      setUserRole(null);
      setUserProfile(null);
      setRoleError("");
      setRoleLoading(false);
      return () => { cancelled = true; };
    }

    // The database profile is the source of truth for application permissions.
    // Never infer a role from Auth metadata or silently default to "teacher".
    setUserRole(null);
    setUserProfile(null);
    setRoleError("");
    setRoleLoading(true);

    (async () => {
      try {
        if (session.offline) {
          // Signed in without internet: use the profile saved on this phone (teachers only)
          let c = null; try { c = JSON.parse(localStorage.getItem("sbcc_profile_"+session.user.id)||"null"); } catch {}
          if (c && c.role==="teacher") { if (!cancelled) { setUserProfile(c); setUserRole("teacher"); } return; }
          throw new Error("Offline sign-in needs a teacher profile saved on this phone. Connect to the internet and sign in once.");
        }
        const { data, error } = await supabase
          .from("profiles")
          .select("*")
          .eq("id", session.user.id)
          .maybeSingle();

        if (error) throw error;
        if (!data) throw new Error("No profile was found for this account. Ask an administrator to create or repair your profile.");
        if (!["admin", "teacher"].includes(data.role)) {
          throw new Error("This account has an unsupported or missing profile role. Access has been blocked.");
        }
        try { localStorage.setItem("sbcc_profile_"+session.user.id, JSON.stringify(data)); } catch {}
        {
          const want = offlineWantRef.current;
          if (want && emailKey(want.email)===emailKey(session.user.email)) {
            offlineWantRef.current = null;
            if (data.role==="teacher") saveOfflineLogin(want.email, want.password, session.user).catch(() => {});
            else removeOfflineLogin(want.email);
          }
        }
        if (!cancelled) {
          setUserProfile(data);
          setUserRole(data.role);
        }
      } catch (error) {
        console.error("Unable to load authenticated profile:", error);
        let cached = null;
        if (/fetch|network|load failed/i.test(error?.message||"")) { try { cached = JSON.parse(localStorage.getItem("sbcc_profile_"+session.user.id)||"null"); } catch {} }
        if (cached && ["admin","teacher"].includes(cached.role)) {
          if (!cancelled) { setUserProfile(cached); setUserRole(cached.role); }
        } else if (!cancelled) {
          setUserRole(null);
          setUserProfile(null);
          setRoleError(error?.message || "Could not verify your account role. Please try again.");
        }
      } finally {
        if (!cancelled) setRoleLoading(false);
      }
    })();

    return () => { cancelled = true; };
  }, [session?.user?.id]);

  // Keep the selected page valid if the profile role changes during a session.
  useEffect(() => {
    if (!userRole) return;
    const selected = navItems.find(n => n.key === page);
    if (!selected || !selected.roles.includes(userRole) || (selected.key==="fees" && !(userRole==="admin" && userProfile?.can_manage_fees===true))) setPage("dashboard");
  }, [userRole, userProfile, page]);

  // ── Load all data ──────────────────────────────────────────────────────────
  useEffect(() => {
    if (!session) return;
    loadAll();
  }, [session]);

  function applySnapshot(snap) {
    // Open with the last data seen online. Marks entered now go into the offline queue.
    if (ACAD_YEARS.includes(snap.settings?.current_year)) ACTIVE_YEAR = snap.settings.current_year;
    setStudents(snap.students||[]); setTeachers(snap.teachers||[]);
    const mm = {}; (snap.marks||[]).forEach(m => { mm[`${m.student_id}-${m.subject}-${m.seq}-${m.acad_year}`] = m; });
    setMarksMap(overlayQueuedMarks(mm));
    const fm = {}; (snap.fees||[]).forEach(f => { fm[f.student_id] = { paid:f.paid||0, total:f.total||TOTAL_FEE }; }); setFeesMap(fm);
    setSettings(snap.settings||{}); setMarkDeadlines(snap.deadlines||[]); setMarkUnlocks(snap.unlocks||[]);
    setIsOnline(false); setDbError(null);
  }
  async function loadAll() {
    if (sessionRef.current?.offline) {
      // Offline sign-in: no network calls, just the data saved on this phone
      const snap = await readSnapshot(sessionRef.current?.user?.id);
      if (snap) applySnapshot(snap); else setDbError("No class lists are saved on this phone yet. Connect to the internet and sign in once.");
      setLoading(false);
      return;
    }
    setLoading(true);
    try {
      const [
        { data: st, error: stErr }, { data: tc }, { data: mk },
        { data: fe }, { data: nt }, { data: at }, { data: cal },
        { data: cm, error: cmErr }, { data: cd, error: cdErr }
      ] = await Promise.all([
        supabase.from("students").select("*").order("form").order("name"),
        supabase.from("teachers").select("*").order("name"),
        supabase.from("marks").select("*"),
        supabase.from("fees").select("*"),
        supabase.from("notices").select("*").order("posted_date", { ascending: false }),
        supabase.from("attendance").select("*"),
        supabase.from("calendar_events").select("*").order("event_date"),
        supabase.from("council_meetings").select("*").order("meeting_date", { ascending: false }),
        supabase.from("council_decisions").select("*"),
      ]);
      if (stErr) throw stErr;
      setStudents(st || []);
      setTeachers(tc || []);
      // Build marks map: "studentId-subject-seq-year" → record
      const mm = {};
      (mk||[]).forEach(m => { mm[`${m.student_id}-${m.subject}-${m.seq}-${m.acad_year}`] = m; });
      setMarksMap(overlayQueuedMarks(mm)); // marks still waiting to upload stay visible
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
      setCouncilMeetings(cm || []);
      setCouncilDecisions(cd || []);
      setCouncilReady(!cmErr && !cdErr);
      // Optional tables from migration.sql — the app still works if they don't exist yet
      const [{ data: ss }, { data: dl }, { data: ul }, { data: rm }] = await Promise.all([
        supabase.from("app_settings").select("*"),
        supabase.from("mark_deadlines").select("*"),
        supabase.from("mark_unlocks").select("*"),
        supabase.from("class_remarks").select("*"),
      ]);
      setRemarksMap(Object.fromEntries((rm||[]).map(r => [`${r.student_id}|${r.acad_year}|${r.term}`, r])));
      const settingsObj = Object.fromEntries((ss||[]).map(r => [r.key, r.value]));
      if (ACAD_YEARS.includes(settingsObj.current_year)) ACTIVE_YEAR = settingsObj.current_year;
      if (/^\d{4}-\d{2}-\d{2}$/.test(settingsObj.late_cutoff||"")) LATE_CUTOFF = settingsObj.late_cutoff;
      setSettings(settingsObj);
      setMarkDeadlines(dl || []);
      setMarkUnlocks(ul || []);
      setDbError(null);
      writeSnapshot(sessionRef.current?.user?.id, { students:(st||[]).map(({ photo_url, ...r }) => r), teachers:tc||[], marks:mk||[], fees:fe||[], settings:settingsObj, deadlines:dl||[], unlocks:ul||[] });
    } catch(e) {
      const snap = await readSnapshot(sessionRef.current?.user?.id);
      if (snap && /fetch|network|load failed/i.test(e?.message||"")) {
        applySnapshot(snap);
      } else setDbError(e.message);
    }
    setLoading(false);
  }

  // ── Install on the phone + offline shell (needs public/manifest.webmanifest and public/sw.js) ──
  useEffect(() => {
    if (!document.querySelector('link[rel="manifest"]')) { const l = document.createElement("link"); l.rel = "manifest"; l.href = "/manifest.webmanifest"; document.head.appendChild(l); }
    if ("serviceWorker" in navigator && window.location.protocol==="https:") navigator.serviceWorker.register("/sw.js").catch(() => {});
  }, []);

  // ── Real-time subscriptions ─────────────────────────────────────────────────
  useEffect(() => {
    if (!session || session.offline) return;
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
  const isNetErr = m => /fetch|network|load failed|timeout/i.test(m||"");
  async function doLogin(email, password, remember = true) {
    // WiFi/data with no real internet makes a normal sign-in hang, so check first (long timeout for slow networks).
    // With no offline sign-in saved for this email there is nothing to fall back on, so always try online then.
    const reachable = navigator.onLine && await checkRealConnectivity(8000);
    if (navigator.onLine && (reachable || !readOfflineLogins()[emailKey(email)])) {
      try {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (!error) { offlineWantRef.current = remember ? { email, password } : null; if (!remember) removeOfflineLogin(email); return; }
        if (!isNetErr(error.message)) throw error; // wrong password etc.
      } catch(e) {
        if (!isNetErr(e?.message)) throw e;
      }
    }
    // No internet: sign in against the check saved on this phone (teachers only)
    const rec = await checkOfflineLogin(email, password);
    OFFLINE_PW_MEM = password;
    setSession({ user:{ id:rec.id, email:rec.email }, offline:true });
  }
  async function doLogout() {
    OFFLINE_PW_MEM = null;
    try { await supabase.auth.signOut(); } catch {}
    setSession(null); // also ends an offline sign-in, and covers a sign-out attempted with no internet
    setMenuOpen(false); setPage("dashboard");
  }
  // When the internet returns during an offline sign-in, sign in for real so queued marks can upload
  async function upgradeOfflineSession() {
    const s = sessionRef.current;
    if (!s?.offline || upgradingRef.current || !OFFLINE_PW_MEM) return;
    upgradingRef.current = true;
    try {
      const { error } = await supabase.auth.signInWithPassword({ email:s.user.email, password:OFFLINE_PW_MEM });
      if (!error) offlineWantRef.current = { email:s.user.email, password:OFFLINE_PW_MEM }; // renews the 30 days
      else if (/invalid login/i.test(error.message||"")) {
        removeOfflineLogin(s.user.email); OFFLINE_PW_MEM = null; setSession(null);
        alert("Your password has changed, so please sign in again. Marks you entered offline are kept and will upload after you sign in.");
      }
      // a network error just means we stay offline and try again shortly
    } catch {}
    upgradingRef.current = false;
  }
  useEffect(() => {
    if (!session?.offline) return;
    const t = setInterval(async () => { if (await checkRealConnectivity()) upgradeOfflineSession(); }, 15000);
    return () => clearInterval(t);
  }, [session?.offline]);
  // After a real sign-in, upload anything this person queued on this phone
  useEffect(() => {
    if (!session?.user?.id || session.offline) return;
    if (loadOfflineQueue().length) syncOfflineQueue();
  }, [session?.user?.id, session?.offline]);
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
      ...(t.assignments !== undefined ? { assignments: t.assignments } : {}),
      ...(t.class_teacher_of !== undefined ? { class_teacher_of: t.class_teacher_of } : {}),
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

  async function saveRemarks(rows) {
    const now = new Date().toISOString();
    const { error } = await supabase.from("class_remarks").upsert(
      rows.map(r => ({ ...r, updated_by:currentUser.name, updated_at:now })),
      { onConflict:"student_id,acad_year,term" });
    if (error) throw error;
    await loadAll();
  }

  async function saveSetting(key, value) {
    const { error } = await supabase.from("app_settings").upsert({ key, value }, { onConflict:"key" });
    if (error) throw error;
    await loadAll();
  }

  async function saveDeadline(year, seq, date) {
    const q = supabase.from("mark_deadlines");
    const { error } = date
      ? await q.upsert({ acad_year:year, seq, deadline:date }, { onConflict:"acad_year,seq" })
      : await q.delete().eq("acad_year", year).eq("seq", seq);
    if (error) throw error;
    await loadAll();
  }

  async function setClassUnlocked(form, year, seq, unlocked) {
    const q = supabase.from("mark_unlocks");
    const { error } = unlocked
      ? await q.upsert({ form, acad_year:year, seq, unlocked_by:currentUser.name }, { onConflict:"form,acad_year,seq" })
      : await q.delete().eq("form", form).eq("acad_year", year).eq("seq", seq);
    if (error) throw error;
    await loadAll();
  }

  async function saveFee(studentId, paid) {
    const { error } = await supabase.from("fees").upsert(
      { student_id: studentId, paid, total: TOTAL_FEE },
      { onConflict:"student_id" }
    );
    if (error) throw error;
    await loadAll();
  }

  // ── Offline sync engine ─────────────────────────────────────────────────────
  // Processes the queue in the order items were created. For a queued
  // registration, generates the REAL matricule now (only safe to do with a
  // genuine connection, since it needs an accurate count of existing students
  // in that form) and swaps it in everywhere — including on any queued fee
  // payment for that same offline-registered student, matched by temp ID.
  async function syncOfflineQueue() {
    const cur = sessionRef.current;
    if (!cur || cur.offline) return; // uploading needs a real signed-in session
    const queue = loadOfflineQueue().filter(it => !it.uid || it.uid===cur.user.id);
    if (!queue.length || syncing) return;

    // navigator.onLine can be wrong — it only reflects whether the device has
    // a network interface active, not whether that network can actually reach
    // the internet (e.g. connected to WiFi with no working connection). Do a
    // real check before attempting to push anything, so a false "online"
    // signal doesn't cause a burst of failed writes.
    const reallyOnline = await checkRealConnectivity();
    if (!reallyOnline) {
      setIsOnline(false);
      return;
    }

    setSyncing(true);
    const success = [];
    const failed = [];
    let tempIdMap = {}; // tempId -> real matricule, built as registrations sync

    for (const item of queue) {
      try {
        if (item.type === "register") {
          const fNum = item.data.form.replace("Form ","");
          const { data: usedIds } = await supabase
            .from("students").select("id").like("id", `SBCC0${fNum.padStart(2,"0")}%`);
          const realId = nextMatricule(item.data.form, [...(usedIds||[]).map(r=>r.id), ...Object.values(tempIdMap)]);
          tempIdMap[item.tempId] = realId;

          const row = {
            id: realId, name: item.data.name, form: item.data.form, gender: item.data.gender,
            dob: item.data.dob||null, parent: item.data.parent, phone: item.data.phone,
            address: item.data.address||null, photo_url: item.data.photo_url||null,
            active: true, reg_status:"registered", reg_date: item.data.reg_date,
            reg_fee: item.data.reg_fee, reg_receipt: item.data.reg_receipt,
            reg_paid_by: item.data.reg_paid_by, reg_cashier: item.data.reg_cashier,
            is_late_reg: item.data.is_late_reg||false, graduated:false,
            parent_pin: item.data.parent_pin,
          };
          const { error } = await supabase.from("students").upsert(row, { onConflict:"id" });
          if (error) throw error;
          await supabase.from("fees").upsert({ student_id: realId, paid:0, total: TOTAL_FEE }, { onConflict:"student_id", ignoreDuplicates:true });
          success.push({ ...item, realId });

        } else if (item.type === "payment") {
          // If this payment was for a student registered offline in the same
          // batch, resolve their temp ID to the real one just assigned above.
          const targetId = tempIdMap[item.data.studentId] || item.data.studentId;
          const { data: existingFee } = await supabase.from("fees").select("paid").eq("student_id", targetId).single();
          const newPaid = (existingFee?.paid||0) + item.data.amount;
          const { error } = await supabase.from("fees").upsert(
            { student_id: targetId, paid: newPaid, total: TOTAL_FEE },
            { onConflict:"student_id" }
          );
          if (error) throw error;
          success.push(item);

        } else if (item.type === "marks") {
          await pushMarkRows(item.data.rows, item.data.oldScores||{}, item.data.who);
          success.push(item);
        }
        removeFromOfflineQueue(item.id);
      } catch(e) {
        console.error("Sync failed for item:", item, e);
        failed.push({ ...item, error: e.message });
      }
    }

    setOfflineQueue(loadOfflineQueue());
    setSyncResult({ success, failed });
    if (success.length) await loadAll();
    setSyncing(false);
  }

  // Watch for connectivity changes and auto-sync when we come back online
  useEffect(() => {
    function handleOnline() { setIsOnline(true); if (sessionRef.current?.offline) { upgradeOfflineSession(); return; } loadAll(); syncOfflineQueue(); }
    function handleOffline() { setIsOnline(false); }
    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    // On mount, navigator.onLine can be wrong (e.g. connected to WiFi with no
    // real internet) — verify properly, correct isOnline if needed, and sync
    // any leftover queue from a previous session that ended before syncing.
    (async () => {
      const reallyOnline = await checkRealConnectivity();
      setIsOnline(reallyOnline);
      if (reallyOnline && loadOfflineQueue().length) syncOfflineQueue();
    })();
    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, []);

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

  // ── Class Council ──────────────────────────────────────────────────────────
  // Saves one meeting and keeps its decisions in sync (update kept ones, add new ones, remove deleted ones)
  async function saveCouncilMeeting(m, decisions) {
    const row = {
      form: m.form, acad_year: m.acad_year, term: m.term, meeting_date: m.meeting_date,
      attendees: m.attendees || [], discussion: m.discussion || null, snapshot: m.snapshot || null,
    };
    let meetingId = m.id;
    if (meetingId) {
      const { error } = await supabase.from("council_meetings").update(row).eq("id", meetingId);
      if (error) throw error;
    } else {
      const { data, error } = await supabase.from("council_meetings").insert({ ...row, created_by: currentUser.name }).select().single();
      if (error) throw error;
      meetingId = data.id;
    }
    const keepIds = decisions.filter(d => d.id).map(d => d.id);
    const { data: existing } = await supabase.from("council_decisions").select("id").eq("meeting_id", meetingId);
    const removeIds = (existing || []).map(e => e.id).filter(id => !keepIds.includes(id));
    if (removeIds.length) {
      const { error } = await supabase.from("council_decisions").delete().in("id", removeIds);
      if (error) throw error;
    }
    for (const d of decisions) {
      const drow = {
        meeting_id: meetingId, decision: d.decision, responsible: d.responsible || null,
        deadline: d.deadline || null, status: d.status || "pending",
        completed_on: d.status === "completed" ? (d.completed_on || todayStr()) : null,
      };
      const { error } = d.id
        ? await supabase.from("council_decisions").update(drow).eq("id", d.id)
        : await supabase.from("council_decisions").insert(drow);
      if (error) throw error;
    }
    await loadAll();
  }

  async function deleteCouncilMeeting(id) {
    const { error } = await supabase.from("council_meetings").delete().eq("id", id); // its decisions are removed automatically
    if (error) throw error;
    await loadAll();
  }

  async function updateDecisionStatus(id, status) {
    const { error } = await supabase.from("council_decisions")
      .update({ status, completed_on: status === "completed" ? todayStr() : null }).eq("id", id);
    if (error) throw error;
    await loadAll();
  }

  // ── Current user info ───────────────────────────────────────────────────────
  const currentTeacher = findTeacher(teachers, session?.user?.email);
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
      <img src={SCHOOL_LOGO} alt="SBCC Bawe" style={{width:96,height:96,objectFit:"contain"}}/>
      <div style={{color:C.white,fontWeight:800,fontSize:18}}>Saker Baptist Comprehensive College Bawe</div>
      <div style={{color:C.goldLight,fontSize:13}}>Starting up…</div>
    </div>
  );
  // Password recovery gives a temporary session — intercept it before normal routing
  if (recoveryMode) return <SetNewPasswordScreen onSetPassword={doSetNewPassword} onCancel={doLogout}/>;
  if (!session) return <LoginScreen onLogin={doLogin} onRequestReset={doRequestPasswordReset} onOpenPortal={()=>setPortalMode(true)}/>;

  if (roleLoading) return (
    <div style={{minHeight:"100vh",background:C.navy,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:12,color:C.white}}>
      <img src={SCHOOL_LOGO} alt="SBCC Bawe" style={{width:80,height:80,objectFit:"contain"}}/><strong>Verifying account permissions…</strong>
    </div>
  );
  if (roleError || !userRole) return (
    <div style={{minHeight:"100vh",background:C.grayBg,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:C.white,borderRadius:12,padding:24,maxWidth:480,width:"100%",boxShadow:"0 8px 30px rgba(0,0,0,.08)"}}>
        <h2 style={{marginTop:0,color:C.red}}>Account access could not be verified</h2>
        <p style={{color:C.gray}}>{roleError || "Your profile role has not loaded. No dashboard access has been granted."}</p>
        <button onClick={()=>{setSession({...session});}} style={{padding:"9px 14px",background:C.navy,color:C.white,border:0,borderRadius:7,cursor:"pointer",marginRight:8}}>Retry</button>
        <button onClick={doLogout} style={{padding:"9px 14px",background:"#fff",color:C.navy,border:"1px solid #ccd",borderRadius:7,cursor:"pointer"}}>Sign out</button>
      </div>
    </div>
  );

  // Fees are limited to ONE administrator: the admin profile whose can_manage_fees flag is true (set in the database)
  const canFees = userRole==="admin" && userProfile?.can_manage_fees===true;
  const allowedNavItems = navItems.filter(n => n.roles.includes(userRole) && (n.key!=="fees" || canFees));

  const ctx = {
    auth: { user: currentUser, role: userRole, canFees, offline: !!session?.offline },
    students, teachers, marksMap, feesMap, notices, attendanceMap, calendarEvents,
    saveStudent, deleteStudent, saveTeacher, saveMark,
    remarksMap, saveRemarks,
    settings, markDeadlines, markUnlocks, saveSetting, saveDeadline, setClassUnlocked,
    saveFee, saveNotice, deleteNotice, loadAll, saveAttendanceBulk,
    saveCalendarEvent, deleteCalendarEvent,
    councilMeetings, councilDecisions, councilReady,
    saveCouncilMeeting, deleteCouncilMeeting, updateDecisionStatus,
    offlineQueue, isOnline, syncOfflineQueue, setOfflineQueue,
    applyQueuedMarks: () => setMarksMap(prev => overlayQueuedMarks({ ...prev })),
  };

  return (
    <div style={{fontFamily:"'Segoe UI',system-ui,sans-serif",background:C.grayBg,minHeight:"100vh"}}>
      {/* ── Top bar ─────────────────────────────────────────────────────────── */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:C.navy,padding:"11px 14px",position:"sticky",top:0,zIndex:200}}>
        <div style={{display:"flex",alignItems:"center",gap:9}}>
          <img src={SCHOOL_LOGO} alt="SBCC Bawe" style={{width:34,height:34,objectFit:"contain",flexShrink:0}}/>
          <div>
            <div style={{color:C.white,fontWeight:800,fontSize:13,lineHeight:1.2}}>Saker Baptist Comprehensive College Bawe</div>
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

      {/* ── Offline / pending sync banner — visible on every page ──────────── */}
      {(!isOnline || offlineQueue.length>0) && (
        <div style={{background:isOnline?"#fffbeb":"#fef2f2",borderBottom:`1px solid ${isOnline?"#fde68a":"#fca5a5"}`}}>
          <div onClick={()=>offlineQueue.length>0 && setShowQueueDetails(v=>!v)} style={{padding:"8px 14px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,flexWrap:"wrap",cursor:offlineQueue.length>0?"pointer":"default"}}>
            <div style={{fontSize:11.5,color:isOnline?"#92400e":"#991b1b",fontWeight:600}}>
              {!isOnline
                ? `📡 No internet — ${offlineQueue.length} item${offlineQueue.length!==1?"s":""} will save on this device and upload automatically once you're back online`
                : `⏳ ${offlineQueue.length} item${offlineQueue.length!==1?"s":""} waiting to sync ${offlineQueue.length>0?(showQueueDetails?"▲":"▼"):""}`
              }
            </div>
            {isOnline && offlineQueue.length>0 && (
              <button onClick={(e)=>{e.stopPropagation();syncOfflineQueue();}} disabled={syncing} style={{padding:"5px 12px",background:"#92400e",color:"#fff",border:"none",borderRadius:6,fontSize:11,fontWeight:700,cursor:"pointer",flexShrink:0}}>
                {syncing?"Syncing…":"Sync Now"}
              </button>
            )}
          </div>
          {showQueueDetails && offlineQueue.length>0 && (
            <div style={{padding:"0 14px 10px"}}>
              {offlineQueue.map(item => (
                <div key={item.id} style={{background:"rgba(255,255,255,0.6)",borderRadius:6,padding:"6px 10px",marginBottom:5,fontSize:11,color:"#78350f"}}>
                  {item.type==="register"
                    ? `📋 Registration: ${item.data.name} (${item.data.form})`
                    : item.type==="marks"
                    ? `✏️ Marks: ${item.data.rows.length} mark(s) — ${item.data.rows[0]?.subject||""} ${item.data.rows[0]?.seq||""}`
                    : `💰 Payment: ${item.data.amount.toLocaleString()} FCFA for ${item.data.studentId}`
                  }
                  <span style={{opacity:.7,marginLeft:6}}>· queued {fmtDate(item.queuedAt?.slice(0,10))}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {syncResult && (
        <div style={{background:syncResult.failed.length?"#fef2f2":"#f0fdf4",borderBottom:`1px solid ${syncResult.failed.length?"#fca5a5":"#86efac"}`,padding:"8px 14px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
          <div style={{fontSize:11.5,color:syncResult.failed.length?"#991b1b":"#166534",fontWeight:600}}>
            ✓ Synced {syncResult.success.length} item{syncResult.success.length!==1?"s":""}
            {syncResult.failed.length>0 && ` — ✕ ${syncResult.failed.length} failed`}
          </div>
          <button onClick={()=>setSyncResult(null)} style={{background:"none",border:"none",fontSize:16,cursor:"pointer",color:syncResult.failed.length?"#991b1b":"#166534"}}>×</button>
        </div>
      )}

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
              {navItems.filter(n=>n.roles.includes(ctx.auth.role) && (n.key!=="fees" || ctx.auth.canFees)).map(n => (
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
          {page==="markctl"      && ctx.auth.role==="admin" && <MarksControlPage ctx={ctx}/>}
          {page==="remarks"      && <ClassRemarksPage ctx={ctx}/>}
          {page==="attendance"   && <AttendancePage   ctx={ctx}/>}
          {page==="reports"      && userRole==="admin" && <ReportsPage ctx={ctx}/>}
          {page==="performance"  && <PerformancePage  ctx={ctx}/>}
          {page==="subjres"      && <SubjectResultsPage ctx={ctx}/>}
          {page==="council"      && ctx.auth.role==="admin" && <ClassCouncilPage ctx={ctx}/>}
          {page==="fees"         && ctx.auth.canFees && <FeesPage        ctx={ctx}/>}
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
  const [remember, setRemember] = useState(OFFLINE_LOGIN_DEFAULT_ON);
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
      await onLogin(email.trim(), pass, remember);
      setAttempts(0);
    } catch(e) {
      if (e.friendly) { setErr(e.message); setBusy(false); return; }
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
            <p style={{color:C.goldLight,fontSize:11,margin:"4px 0 0",letterSpacing:1.5}}>SAKER BAPTIST COMPREHENSIVE COLLEGE BAWE</p>
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
          <img src={SCHOOL_LOGO} alt="SBCC Bawe crest" style={{width:110,height:110,objectFit:"contain",display:"block",margin:"0 auto 10px",filter:"drop-shadow(0 3px 10px rgba(0,0,0,0.4))"}}/>
          <h1 style={{color:C.white,fontSize:20,fontWeight:900,margin:0}}>Saker Baptist Comprehensive College Bawe</h1>
          <p style={{color:C.goldLight,fontSize:11,margin:"4px 0 0",letterSpacing:1.5}}>NGEPTANG · NONI</p>
          <p style={{color:"rgba(255,255,255,0.45)",fontSize:10,margin:"5px 0 0"}}>School Administrative Portal · {ACTIVE_YEAR}</p>
        </div>
        <div style={{background:C.white,borderRadius:14,padding:22,boxShadow:"0 20px 60px rgba(0,0,0,0.35)"}}>
          <h2 style={{margin:"0 0 16px",fontSize:16,fontWeight:800,color:C.navy}}>Sign In</h2>
          <Fr label="Email Address">
            <input style={inp} type="email" autoComplete="username" placeholder="your@email.com" value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handle()} disabled={isLocked}/>
          </Fr>
          <Fr label="Password">
            <input style={inp} type="password" autoComplete="current-password" placeholder="••••••••" value={pass} onChange={e=>setPass(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handle()} disabled={isLocked}/>
          </Fr>
          {!navigator.onLine && <div style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:8,padding:"8px 12px",color:"#92400e",fontSize:12,marginBottom:12}}>📡 No internet. Teachers can sign in with the email and password they used on this phone before.</div>}
          <label style={{display:"flex",gap:8,alignItems:"flex-start",fontSize:11.5,color:C.gray,marginBottom:12,cursor:"pointer"}}>
            <input type="checkbox" checked={remember} onChange={e=>setRemember(e.target.checked)} style={{marginTop:2}}/>
            <span>Allow offline sign-in on this phone (teachers). Do not tick on a shared phone.</span>
          </label>
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
          <p style={{color:C.goldLight,fontSize:11,margin:"4px 0 0",letterSpacing:1.5}}>SAKER BAPTIST COMPREHENSIVE COLLEGE BAWE</p>
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
          <p style={{color:C.goldLight,fontSize:11,margin:"4px 0 0",letterSpacing:1.5}}>SAKER BAPTIST COMPREHENSIVE COLLEGE BAWE</p>
          <p style={{color:"rgba(255,255,255,0.45)",fontSize:10,margin:"5px 0 0"}}>View your child's marks, attendance & fees</p>
        </div>
        <div style={{background:C.white,borderRadius:14,padding:22,boxShadow:"0 20px 60px rgba(0,0,0,0.35)"}}>
          <Fr label="Student Matricule">
            <input style={inp} placeholder="e.g. SBCC01001" value={matricule} onChange={e=>setMatricule(e.target.value)} disabled={isLocked} onKeyDown={e=>e.key==="Enter"&&lookup()}/>
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

// ─── Parent Portal → Form 5 subject selection (open until the admin's closing date) ───
function ParentF5Subjects({ matricule, pin }) {
  const all = SUBJECTS_BY_FORM["Form 5"];
  const [st, setSt]     = useState(null);   // status from the database
  const [err, setErr]   = useState("");
  const [pick, setPick] = useState([]);
  const [busy, setBusy] = useState(false);
  const [saved, setSaved] = useState(false);
  const box = { background:C.white, borderRadius:10, padding:12, marginBottom:11, boxShadow:"0 1px 3px rgba(0,0,0,0.06)" };
  useEffect(() => {
    (async () => {
      try {
        const { data, error } = await supabase.rpc("portal_f5_subjects_status", { p_matricule: matricule, p_pin: pin });
        if (error) throw error;
        if (!data?.ok) throw new Error("not available");
        setSt(data);
        setPick(Array.isArray(data.subjects) ? all.filter(x => data.subjects.includes(x)) : []);
      } catch(e) { setErr("Subject selection is not available right now. Please contact the school office."); }
    })();
  }, []);
  async function save() {
    if (!pick.length) { alert("Tick at least one subject."); return; }
    if (!window.confirm(`Save these ${pick.length} subject(s)? You can change them until ${fmtDate(st.deadline)}.`)) return;
    setBusy(true);
    try {
      const { data, error } = await supabase.rpc("portal_f5_submit_subjects", { p_matricule: matricule, p_pin: pin, p_subjects: pick });
      if (error) throw error;
      if (!data?.ok) {
        if (["deadline_passed","closed_by_admin","not_opened"].includes(data?.error)) { setSt(x => ({ ...x, open:false, reason:data.error })); alert("Sorry, subject selection has just closed."); }
        else alert("Could not save. Please try again.");
      } else { setSt(x => ({ ...x, submitted:true, subjects:data.subjects })); setSaved(true); }
    } catch(e) { alert("Could not save: " + e.message); }
    setBusy(false);
  }
  if (err) return <div style={box}><div style={{fontSize:13,color:C.red}}>{err}</div></div>;
  if (!st) return <div style={box}><div style={{fontSize:13,color:C.gray}}>Loading…</div></div>;
  if (!st.applicable) return <div style={box}><div style={{fontSize:13,color:C.gray}}>Subject selection is for Form 5 students only.</div></div>;

  const chosen = all.filter(x => (st.subjects||[]).includes(x));
  if (!st.open) {
    const why = st.reason==="not_opened" ? "The school has not opened subject selection yet."
      : st.reason==="closed_by_admin" ? "Subject selection has been closed by the school."
      : `Subject selection closed on ${fmtDate(st.deadline)}.`;
    return (
      <div style={box}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:6}}>📚 Form 5 subjects</div>
        <div style={{background:"#FEF2F2",border:"1px solid #FCA5A5",borderRadius:8,padding:"8px 12px",color:C.red,fontSize:12,marginBottom:10}}>🔒 {why} Subjects can no longer be filled or changed here.</div>
        {st.submitted && chosen.length>0
          ? <><div style={{fontSize:12,color:C.gray,marginBottom:4}}>Subjects submitted ({chosen.length}):</div>{chosen.map(x => <div key={x} style={{fontSize:13,color:C.navy,padding:"2px 0"}}>✓ {x}</div>)}</>
          : st.reason!=="not_opened" && <div style={{fontSize:12,color:C.gray}}>No subjects were submitted from the portal. Please contact the school office.</div>}
      </div>
    );
  }
  return (
    <div style={box}>
      <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:4}}>📚 Choose the subjects your child will write</div>
      <div style={{background:"#FFFBEB",border:"1px solid #FCD34D",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#92400E",marginBottom:10}}>
        Open until <strong>{fmtDate(st.deadline)}</strong>. You can change your choice until then; after that it is locked for everyone.
      </div>
      {st.submitted && <div style={{fontSize:12,color:C.green,marginBottom:8}}>✓ Your choice is saved{saved ? "" : ` (${chosen.length} subjects)`}. You may still change it before the closing date.</div>}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
        <span style={{fontSize:12,color:C.gray}}>{pick.length} selected</span>
        <button onClick={()=>setPick([])} style={{background:"none",border:"none",color:C.gray,fontSize:12,cursor:"pointer"}}>Clear</button>
      </div>
      {all.map(sub => (
        <label key={sub} style={{display:"flex",alignItems:"center",gap:8,fontSize:13,color:C.navy,padding:"6px 2px",borderBottom:`1px solid ${C.grayBg}`}}>
          <input type="checkbox" checked={pick.includes(sub)} onChange={e=>setPick(b=>e.target.checked?[...b,sub]:b.filter(x=>x!==sub))}/>{sub}
        </label>
      ))}
      <div style={{marginTop:12}}><Btn onClick={save} disabled={busy}>{busy?"Saving…":"Save subjects"}</Btn></div>
    </div>
  );
}

function ParentPortalResults({ data, matricule, pin, onBack, onExit }) {
  const { student, marks, attendance, fee, recent_submissions } = data;
  const [tab, setTab] = useState("marks");
  const [refreshedData, setRefreshedData] = useState(data);
  const [extra, setExtra] = useState(null); // class teacher remarks + next term date (needs migration3.sql)
  useEffect(() => {
    supabase.rpc("parent_report_extra", { p_matricule: matricule, p_pin: pin })
      .then(({ data: d, error }) => { if (!error && d) setExtra(d); }, () => {});
  }, []);

  // Only the latest academic year's marks (when the data says which year)
  const allMarks = marks || [];
  const curYear  = allMarks.map(m => m.acad_year).filter(Boolean).sort().pop();
  const marksNow = (curYear ? allMarks.filter(m => m.acad_year===curYear) : allMarks).filter(m => !isPaperSeq(m.seq));
  const f5Student = isF5(student.form);
  const periods   = f5Student ? F5_ASSESSMENTS : TERMS;
  // The admin's Form 5 grade scale (public, non-sensitive); falls back to the default scale if the SQL function isn't installed
  const [f5Raw, setF5Raw] = useState(null);
  useEffect(() => {
    if (!isF5(student.form)) return;
    supabase.rpc("portal_f5_scale").then(({ data: d, error }) => { if (!error && d) setF5Raw(d); }, () => {});
  }, [student.form]);
  const portalSettings = { f5_grade_scale: f5Raw };
  const paperNow = (curYear ? allMarks.filter(m => m.acad_year===curYear) : allMarks).filter(m => isPaperSeq(m.seq));
  const termAverage = term => {
    const bySub = {};
    const periodSeqs = f5Student ? [term] : TERM_SEQS[term];
    marksNow.filter(m => periodSeqs.includes(m.seq) && m.score!==null && m.score!==undefined)
            .forEach(m => { (bySub[m.subject] = bySub[m.subject] || []).push(m); });
    let tw = 0, tc = 0;
    Object.entries(bySub).forEach(([sub, list]) => {
      const avg = list.reduce((a,m)=>a+Number(m.score),0) / list.length;
      const c = Number(list[0].coeff) || getCoeff(sub);
      tw += avg*c; tc += c;
    });
    return tc ? tw/tc : null;
  };
  const remarkRows = (extra?.remarks||[]).filter(r => !curYear || r.acad_year===curYear)
    .sort((a,b) => TERMS.indexOf(a.term) - TERMS.indexOf(b.term));

  // Group marks by subject, term
  const bySubject = {};
  marksNow.forEach(m => {
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
          {[["marks","📝 Marks"],["attendance","🗓️ Attendance"],["fees","💰 Fees"], ...(f5Student ? [["subjects","📚 Subjects"]] : [])].map(([k,l]) => (
            <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"9px 4px",borderRadius:7,border:"none",cursor:"pointer",fontWeight:700,fontSize:12,background:tab===k?C.navy:"transparent",color:tab===k?C.white:C.gray}}>{l}</button>
          ))}
        </div>

        {tab==="subjects" && f5Student && <ParentF5Subjects matricule={matricule} pin={pin}/>}

        {tab==="marks" && marksNow.length>0 && (
          <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:8}}>{f5Student ? "Mock averages (out of 20) and GCE passes" : "Term averages (out of 20)"}</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,textAlign:"center"}}>
              {periods.map(t => {
                if (t===GCE_SEQ) {
                  const gm = marksNow.filter(m => m.seq===GCE_SEQ && m.score!==null && m.score!==undefined);
                  const ok = gm.filter(m => gcePassed(m.score)).length;
                  return (
                    <div key={t} style={{background:C.grayBg,borderRadius:8,padding:"8px 4px"}}>
                      <div style={{fontSize:10,color:C.gray}}>GCE Result</div>
                      <div style={{fontSize:18,fontWeight:900,color:gm.length?(ok===gm.length?C.green:C.gold):C.gray}}>{gm.length?`${ok}/${gm.length}`:"—"}</div>
                      <div style={{fontSize:10,color:C.gray,fontWeight:700}}>{gm.length?"subjects passed":""}</div>
                    </div>
                  );
                }
                const a = termAverage(t); const g = a!==null ? scoreToGrade(a) : null;
                const g5 = f5Student && a!==null ? f5Grade(portalSettings, a) : null;
                const gc = g5 ? (g5.pass ? C.green : C.red) : g ? gradeCol(g.grade) : C.gray;
                return (
                <div key={t} style={{background:C.grayBg,borderRadius:8,padding:"8px 4px"}}>
                  <div style={{fontSize:10,color:C.gray}}>{f5Student ? t : t.replace(" Term","")}</div>
                  <div style={{fontSize:18,fontWeight:900,color:gc}}>{a!==null?a.toFixed(2):"—"}</div>
                  <div style={{fontSize:10,color:gc,fontWeight:700}}>{g5 ? g5.grade+" · "+g5.remark : g ? g.remark : ""}</div>
                </div>
              ); })}
            </div>
            {extra?.next_term_begins && <div style={{fontSize:12,color:C.navy,marginTop:10}}>📅 Next term begins: <strong>{fmtDate(extra.next_term_begins)}</strong></div>}
          </div>
        )}
        {tab==="marks" && remarkRows.length>0 && (
          <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:8}}>Class teacher's remarks</div>
            {remarkRows.map(r => (
              <div key={r.term} style={{padding:"7px 0",borderTop:`1px solid ${C.grayBg}`}}>
                <div style={{fontSize:11,color:C.gray}}>{r.term}{r.conduct ? <> · Conduct: <strong style={{color:C.navy}}>{r.conduct}</strong></> : null}</div>
                {r.remark && <div style={{fontSize:12.5,color:C.navy,marginTop:2}}>{r.remark}</div>}
              </div>
            ))}
          </div>
        )}
        {tab==="marks" && (
          <div style={{background:C.white,borderRadius:10,overflow:"hidden",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            {Object.keys(bySubject).length===0
              ? <Empty text="No marks recorded yet."/>
              : Object.entries(bySubject).map(([sub, entries], i) => (
                  <div key={sub} style={{padding:"10px 12px",borderBottom:`1px solid ${C.grayBg}`,background:i%2===0?C.white:C.grayBg}}>
                    <div style={{fontWeight:700,color:C.navy,fontSize:13,marginBottom:6}}>{sub}</div>
                    <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                      {entries.sort((a,b)=>seqOrder(a.seq)-seqOrder(b.seq)).map(m=>{
                        if (m.seq===GCE_SEQ) { const gc = gcePassed(m.score) ? C.green : C.red; return <span key={m.seq} style={{fontSize:11,background:gc+"18",color:gc,padding:"3px 8px",borderRadius:6,fontWeight:700}}>GCE: {gceFromPoints(m.score)||"—"}</span>; }
                        if (f5Student) {
                          const g5 = f5Grade(portalSettings, m.score), gc = g5.pass ? C.green : C.red;
                          const pp = paperNow.filter(x => x.subject===sub && x.seq.startsWith(m.seq+" ")).sort((a,b)=>a.seq.localeCompare(b.seq));
                          return (
                            <span key={m.seq} style={{fontSize:11,background:gc+"18",color:gc,padding:"3px 8px",borderRadius:6,fontWeight:700}}>
                              {m.seq}: {Number(m.score).toFixed(1)}/20 · {g5.grade}
                              {pp.length>0 && <span style={{display:"block",fontSize:9.5,fontWeight:600,color:C.gray}}>{pp.map(x=>"P"+x.seq.slice(-1)+" "+x.score).join(" · ")}</span>}
                            </span>
                          );
                        }
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
            <li>Confirm the payment of <strong>{Number(amount).toLocaleString()} FCFA</strong> to Saker Baptist Comprehensive College Bawe</li>
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
// ─── Missing-marks alert (teachers' dashboard) ────────────────────────────────
function MissingMarksAlert({ ctx, setPage }) {
  const { auth, teachers, students, marksMap, markDeadlines, markUnlocks } = ctx;
  if (auth.role !== "teacher") return null;
  const teacher = findTeacher(teachers, auth.user.email);
  if (!teacher) return null;
  const today = new Date(localToday());
  const items = [];
  markDeadlines.forEach(d => {
    FORMS.forEach(form => {
      if (!seqsForForm(form).includes(d.seq)) return; // this assessment does not apply to this class
      const subs = teacherSubjectsIn(teacher, form).filter(sb => SUBJECTS_BY_FORM[form]?.includes(sb));
      if (!subs.length || lockStatus(form, d.seq, d.acad_year, markDeadlines, markUnlocks).locked) return;
      const pool0 = students.filter(s => s.active && s.reg_status==="registered" && s.form===form);
      if (!pool0.length) return;
      subs.forEach(sub => {
        const pool = pool0.filter(s => f5Writes(ctx.settings, s, sub));
        if (!pool.length) return;
        const n = pool.filter(s => marksMap[`${s.id}-${sub}-${d.seq}-${d.acad_year}`]).length;
        if (n < pool.length) items.push({ seq:d.seq, deadline:d.deadline, form, sub, n, total:pool.length });
      });
    });
  });
  if (!items.length) return null;
  items.sort((a,b) => a.deadline.localeCompare(b.deadline) || a.form.localeCompare(b.form));
  const shown = items.slice(0, 8);
  const days = dl => Math.round((new Date(dl) - today) / 864e5);
  return (
    <div style={{background:"#fffbeb",border:"1px solid #fcd34d",borderRadius:11,padding:12,marginBottom:14}}>
      <div style={{fontWeight:800,color:"#92400e",fontSize:13,marginBottom:6}}>⚠️ Marks still to enter ({items.length})</div>
      {shown.map((x,i) => {
        const dd = days(x.deadline);
        return (
          <div key={i} style={{display:"flex",justifyContent:"space-between",gap:8,padding:"4px 0",borderTop:i?"1px solid #fde68a":"none",fontSize:12}}>
            <span style={{color:C.navy}}><strong>{x.form}</strong> · {x.sub} · {x.seq}</span>
            <span style={{color:dd<0?C.red:"#92400e",fontWeight:700,whiteSpace:"nowrap"}}>{x.n}/{x.total} · {dd<0?"overdue":dd===0?"due today":`${dd}d left`}</span>
          </div>
        );
      })}
      {items.length>shown.length && <div style={{fontSize:11,color:"#92400e",marginTop:4}}>+ {items.length-shown.length} more</div>}
      <button onClick={()=>setPage("marks")} style={{marginTop:9,padding:"8px 14px",background:"#92400e",color:C.white,border:"none",borderRadius:8,fontWeight:700,fontSize:12,cursor:"pointer"}}>Enter marks</button>
    </div>
  );
}

// ─── Full backup to Excel (admin) ─────────────────────────────────────────────
async function fetchAllRows(table, orderCols) {
  const out = [], size = 1000;
  for (let from = 0; ; from += size) {
    let q = supabase.from(table).select("*");
    orderCols.forEach(c => { q = q.order(c); });
    const { data, error } = await q.range(from, from+size-1);
    if (error) throw error;
    out.push(...data);
    if (data.length < size) break;
  }
  return out;
}

async function exportBackup(canFees) {
  const [students, teachers, marks, fees, att, rem, dl] = await Promise.all([
    fetchAllRows("students", ["id"]),
    fetchAllRows("teachers", ["id"]),
    fetchAllRows("marks", ["student_id","subject","seq","acad_year"]),
    canFees ? fetchAllRows("fees", ["student_id"]) : Promise.resolve([]),
    fetchAllRows("attendance", ["student_id","date"]),
    supabase.from("class_remarks").select("*").then(r => r.data || []),
    supabase.from("mark_deadlines").select("*").then(r => r.data || []),
  ]);
  const st = Object.fromEntries(students.map(s => [s.id, s]));
  const tn = Object.fromEntries(teachers.map(t => [t.id, t.name]));
  const nm = id => st[id]?.name || "";
  const fm = id => st[id]?.form || "";
  exportToExcelMultiSheet([
    { name:"Students", rows: students.map(s => ({ Matricule:s.id, Name:s.name, Form:s.form, Gender:s.gender||"", "Date of birth":s.dob||"", Parent:s.parent||"", Phone:s.phone||"", Address:s.address||"", Registration:s.reg_status||"", Active:s.active===false?"No":"Yes", Graduated:s.graduated?"Yes":"No" })) },
    { name:"Marks", rows: marks.filter(m => !isPaperSeq(m.seq)).map(m => ({ Matricule:m.student_id, Name:nm(m.student_id), Form:fm(m.student_id), Subject:m.subject, Sequence:m.seq, Year:m.acad_year, "Score /20":m.score, Coeff:m.coeff, Teacher:tn[m.teacher_id]||"" })) },
    ...(canFees ? [{ name:"Fees", rows: fees.map(f => ({ Matricule:f.student_id, Name:nm(f.student_id), Form:fm(f.student_id), Paid:f.paid||0, Total:f.total||TOTAL_FEE, Balance:(f.total||TOTAL_FEE)-(f.paid||0) })) }] : []),
    { name:"Attendance", rows: att.map(a => ({ Matricule:a.student_id, Name:nm(a.student_id), Form:fm(a.student_id), Date:a.date, Status:a.status })) },
    { name:"Teachers", rows: teachers.map(t => ({ Name:t.name, Email:t.email, Active:t.active===false?"No":"Yes", "Class teacher of":t.class_teacher_of||"", Classes: FORMS.filter(f=>teacherSubjectsIn(t,f).length).map(f=>f+": "+teacherSubjectsIn(t,f).join(", ")).join(" | ") })) },
    { name:"Class remarks", rows: rem.map(r => ({ Matricule:r.student_id, Name:nm(r.student_id), Form:fm(r.student_id), Year:r.acad_year, Term:r.term, Conduct:r.conduct||"", Remark:r.remark||"" })) },
    { name:"Deadlines", rows: dl.map(d => ({ Year:d.acad_year, Sequence:d.seq, Deadline:d.deadline })) },
  ], `SBCC-Backup-${todayStr()}.xlsx`);
}

function DashboardPage({ ctx, setPage }) {
  const { students, feesMap, notices, auth } = ctx;
  const [backingUp, setBackingUp] = useState(false);
  const active     = students.filter(s => s.active);
  const regCount   = active.filter(s => s.reg_status==="registered").length;
  const pendCount  = active.filter(s => s.reg_status==="pending").length;
  const collected  = active.reduce((a,s) => a+(feesMap[s.id]?.paid||0), 0);

  const cards = [
    { label:"Registered",      value:regCount,                   icon:"✅", color:C.green,    page:"students" },
    { label:"Pending Reg.",     value:pendCount,                  icon:"⏳", color:C.red,      page:"registration" },
    ...(auth.canFees ? [{ label:"Fees Collected", value:`${Math.round(collected/1000)}k F`, icon:"💰", color:C.gold, page:"fees" }] : []),
  ];

  return (
    <div>
      <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,borderRadius:12,padding:"18px",marginBottom:16,color:C.white,display:"flex",alignItems:"center",gap:14}}>
        <img src={SCHOOL_LOGO} alt="SBCC Bawe crest" style={{width:76,height:76,objectFit:"contain",flexShrink:0,filter:"drop-shadow(0 2px 6px rgba(0,0,0,0.35))"}}/>
        <div style={{minWidth:0}}>
          <h2 style={{margin:0,fontSize:17,fontWeight:800}}>Welcome, {auth.user.name.split(" ")[0]}! 👋</h2>
          <p style={{margin:"4px 0 0",color:"rgba(255,255,255,0.6)",fontSize:12}}>Academic Year {ACTIVE_YEAR} · Saker Baptist Comprehensive College Bawe</p>
        </div>
      </div>
      <MissingMarksAlert ctx={ctx} setPage={setPage}/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
        {cards.map(c => (
          <div key={c.label} onClick={()=>setPage(c.page)} style={{background:C.white,borderRadius:11,padding:13,borderLeft:`4px solid ${c.color}`,cursor:"pointer",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{fontSize:20,marginBottom:3}}>{c.icon}</div>
            <div style={{fontSize:22,fontWeight:900,color:c.color}}>{c.value}</div>
            <div style={{fontSize:11,color:C.gray,marginTop:1}}>{c.label}</div>
          </div>
        ))}
      </div>
      {auth.role==="admin" && (
        <div style={{marginBottom:11}}>
          <Card title="💾 Backup">
            <div style={{fontSize:11.5,color:C.gray,marginBottom:8}}>One Excel file with students, marks, {auth.canFees ? "fees, " : ""}attendance, teachers and class remarks. Photos are not included.</div>
            <Btn disabled={backingUp} onClick={async()=>{ setBackingUp(true); try { await exportBackup(auth.canFees); } catch(e) { alert("Backup failed: "+e.message); } setBackingUp(false); }}>{backingUp?"Preparing…":"Download backup (Excel)"}</Btn>
          </Card>
        </div>
      )}
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
  const { students, saveStudent, auth, isOnline, setOfflineQueue } = ctx;
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
    const rec  = `RCP-${new Date().getFullYear()}-${String(Date.now()).slice(-4)}`;
    const pin  = String(Math.floor(1000 + Math.random()*9000));

    try {
      // Compress photo first — needs to happen either way, online or offline
      let photoUrl = null;
      if (form.photo_url) {
        photoUrl = await compressPhoto(form.photo_url);
      }
      const baseData = {
        name:form.name, form:form.form, gender:form.gender,
        dob:form.dob||null, parent:form.parent, phone:form.phone,
        address:form.address||null, photo_url:photoUrl,
        reg_date:todayStr(), reg_fee:fee, reg_receipt:rec,
        reg_paid_by:form.paidBy||form.parent,
        reg_cashier:auth.user.name,
        is_late_reg:form.isLate,
        parent_pin: pin,
      };

      if (!isOnline) {
        // No internet — save with a temporary ID, queue for later sync
        const tempId = generateTempId(form.form);
        const queuedItem = { type:"register", data:{ ...baseData, studentId:tempId }, tempId };
        addToOfflineQueue(queuedItem);
        setOfflineQueue(loadOfflineQueue());
        setReceipt({ ...baseData, id:tempId, _offline:true });
        setTab("receipt");
        setForm(blank);
      } else {
        // Online — register normally, exactly as before
        const fNum = form.form.replace("Form ","");
        const id   = nextMatricule(form.form, students.map(s=>s.id));
        const studentData = { ...baseData, id, active:true, reg_status:"registered" };
        await saveStudent(studentData);
        setReceipt({ ...studentData });
        setTab("receipt");
        setForm(blank);
      }
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
      ["Cashier",receipt.reg_cashier],["Academic Year",ACTIVE_YEAR],
    ].map(([l,v]) => `<div style="display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #f0f0f0;font-size:11.5px"><span style="color:#6B7280">${l}</span><span style="font-weight:600">${v||"—"}</span></div>`).join("");

    const amtCol = receipt.reg_fee===REG_LATE ? "#C9962A" : "#15803D";
    const html = `<div style="font-family:Segoe UI,sans-serif;padding:14px;color:#0D2340">
      <div style="max-width:360px;margin:0 auto;border:2px solid #0D2340;border-radius:10px;overflow:hidden">
        <div style="background:#0D2340;color:#fff;padding:14px;text-align:center">
          <img src="${SCHOOL_LOGO}" alt="SBCC" style="width:54px;height:54px;object-fit:contain">
          <h2 style="font-size:14px;font-weight:900;margin:4px 0">SAKER BAPTIST COMPREHENSIVE COLLEGE BAWE</h2>
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
        <div style="text-align:center;font-size:9px;color:#9ca3af;border-top:1px dashed #e5e7eb;padding:8px 14px">Academic Year ${ACTIVE_YEAR} · Keep for your records.</div>
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
                <input ref={photoInputRef} type="file" accept="image/*" style={{display:"none"}} onChange={e=>{
                  const f=e.target.files[0]; if(!f) return;
                  const inputEl = e.target;
                  const r=new FileReader();
                  r.onload = async ev => {
                    try {
                      const compressed = await compressPhoto(ev.target.result);
                      setForm(prev=>({...prev,photo_url:compressed}));
                    } catch(err) {
                      console.error("Photo compress error:", err);
                      alert("Could not process that photo. Please try a different image.");
                    } finally {
                      inputEl.value=""; // reset only after the read+compress is fully done, so choosing the same file again works
                    }
                  };
                  r.onerror = () => {
                    alert("Could not read that photo. Please try again.");
                    inputEl.value="";
                  };
                  r.readAsDataURL(f);
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
              {receipt._offline && (
                <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:10,padding:"10px 14px",marginBottom:12,maxWidth:360,marginLeft:"auto",marginRight:"auto"}}>
                  <div style={{fontWeight:800,color:"#991b1b",fontSize:12,marginBottom:3}}>📡 Saved Offline</div>
                  <div style={{fontSize:11,color:"#991b1b"}}>This student's real Matricule will be assigned automatically once this device reconnects to the internet and syncs. The ID shown below is temporary.</div>
                </div>
              )}
              <div style={{background:C.white,borderRadius:11,overflow:"hidden",border:`2px solid ${C.navy}`,maxWidth:360,margin:"0 auto 13px"}}>
                <div style={{background:`linear-gradient(135deg,${C.navy},${C.navyMid})`,color:C.white,padding:"14px",textAlign:"center"}}>
                  <img src={SCHOOL_LOGO} alt="SBCC Bawe" style={{width:54,height:54,objectFit:"contain"}}/>
                  <h2 style={{margin:"4px 0",fontSize:14,fontWeight:900}}>SAKER BAPTIST COMPREHENSIVE COLLEGE BAWE</h2>
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
    exportToExcel(sample, "SBCC-Bulk-Import-Template.xlsx", "Students");
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
        FORMS.forEach(f => { formCounts[f] = 0; }); // rows of this file already given an ID, per form

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
            id = nextMatricule(form, students.map(x=>x.id), formCounts[form]-1);
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
  const [showRollover, setShowRollover] = useState(false);
  const [showIdCards, setShowIdCards] = useState(false);
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
      ...(auth.canFees ? { "Fees Paid (FCFA)":    feesMap?.[s.id]?.paid ?? 0, "Fees Total (FCFA)":   feesMap?.[s.id]?.total ?? TOTAL_FEE, "Balance (FCFA)":      (feesMap?.[s.id]?.total ?? TOTAL_FEE) - (feesMap?.[s.id]?.paid ?? 0) } : {}),
    };
  }

  function exportStudents() {
    if (filter.form) {
      // A single form is already selected — one sheet is enough
      const rows = filtered.map(studentExportRow);
      exportToExcel(rows, `SBCC-Students-${filter.form.replace(" ","")}-${sel_year_safe()}.xlsx`, filter.form);
    } else {
      // No form filter — split into one sheet per class, all in one workbook
      const sheets = FORMS.map(f => ({
        name: f,
        rows: filtered.filter(s=>s.form===f).map(studentExportRow),
      }));
      exportToExcelMultiSheet(sheets, `SBCC-Students-AllForms-${sel_year_safe()}.xlsx`);
    }
  }
  function sel_year_safe(){ return ACTIVE_YEAR.replace("/","-"); }

  async function save() {
    if (!form.name?.trim()) return;
    setSaving(true);
    try {
      if (modal==="add") {
        const fNum = form.form.replace("Form ","");
        const id   = nextMatricule(form.form, students.map(s=>s.id));
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
        {auth.role==="admin" && <SmBtn onClick={()=>setShowRollover(true)} color={C.navy}>🔄 New Academic Year</SmBtn>}
        {auth.role==="admin" && <SmBtn onClick={()=>setShowIdCards(true)} color={C.navyMid}>🪪 ID Cards</SmBtn>}
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
                const inputEl = e.target;
                const r=new FileReader();
                r.onload = async ev => {
                  try {
                    const compressed = await compressPhoto(ev.target.result);
                    setForm(prev=>({...prev,photo_url:compressed}));
                  } catch(err) {
                    console.error("Photo compress error:", err);
                    alert("Could not process that photo. Please try a different image.");
                  } finally {
                    inputEl.value="";
                  }
                };
                r.onerror = () => {
                  alert("Could not read that photo. Please try again.");
                  inputEl.value="";
                };
                r.readAsDataURL(f);
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

      {showPromote && <PromoteStudentsModal students={students} saveStudent={saveStudent} saveFee={ctx.saveFee} canFees={ctx.auth.canFees} onClose={()=>setShowPromote(false)}/>}
      {showRollover && <YearRolloverModal ctx={ctx} onClose={()=>setShowRollover(false)}/>}
      {showIdCards && <IdCardsModal students={students} onClose={()=>setShowIdCards(false)}/>}
    </div>
  );
}

// ─── Promote Students (end-of-year bulk promotion) ──────────────────────────────
function PromoteStudentsModal({ students, saveStudent, saveFee, canFees, onClose }) {
  const [fromForm, setFromForm] = useState("Form 1");
  const toFormOptions = FORMS.slice(FORMS.indexOf(fromForm)+1);
  const isGraduating = fromForm === "Form 5";
  const [toForm, setToForm] = useState(isGraduating ? "" : (FORMS[FORMS.indexOf(fromForm)+1]||""));
  const [resetFees, setResetFees] = useState(!!canFees);
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
              Form 5 is the final form. These students will be marked as <strong>graduated</strong> and moved out of active rolls. Their records (marks, report cards, fee history, candidate number and GCE results) are preserved. Print results slips from Marks Control → Form 5 → Class summary → GCE Result.
            </div>
          )}
          {!isGraduating && (
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:12}}>
              <input type="checkbox" id="resetFees" disabled={!canFees} checked={resetFees} onChange={e=>setResetFees(e.target.checked)}/>
              <label htmlFor="resetFees" style={{fontSize:12,color:C.navy}}>Reset fees to 0 paid for the new academic year{!canFees ? " (only the fee administrator can do this)" : ""}</label>
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

// ─── New academic year (rollover) ──────────────────────────────────────────────
// One guided step: graduates Form 5, moves every other class up, resets fees, switches the app to the new year
// and clears the old "next term begins" date. Marks, remarks and deadlines of the old year stay under that year.
function YearRolloverModal({ ctx, onClose }) {
  const { students, settings, loadAll } = ctx;
  const cur  = ACAD_YEARS.includes(settings?.current_year) ? settings.current_year : ACAD_YEARS[0];
  const next = ACAD_YEARS[ACAD_YEARS.indexOf(cur)+1] || "";
  const prev = parseJSON(settings?.rollover_state, {});
  const interrupted = prev.status==="running" && prev.to===next;
  const [step, setStep]         = useState("plan"); // plan | confirm | running | done
  const [excluded, setExcluded] = useState({});     // student id → true = stays in the class (repeats)
  const [open, setOpen]         = useState("");
  const [resetFees, setResetFees] = useState(!!ctx.auth.canFees);
  const [cutoff, setCutoff]     = useState(next ? `${next.slice(0,4)}-09-30` : "");
  const [typed, setTyped]       = useState("");
  const [ack, setAck]           = useState(false);
  const [progress, setProgress] = useState({ done:0, total:1 });
  const [result, setResult]     = useState(null);

  const classes = [...FORMS].reverse().map(f => {
    const mine = students.filter(s => s.active && s.form===f);
    return { from:f, to: f==="Form 5" ? null : FORMS[FORMS.indexOf(f)+1],
      list: mine.filter(s=>s.reg_status==="registered").sort((a,b)=>a.name.localeCompare(b.name)),
      notReg: mine.filter(s=>s.reg_status!=="registered").length };
  });
  const goCount   = c => c.list.filter(s=>!excluded[s.id]).length;
  const totalGo   = classes.reduce((a,c)=>a+goCount(c),0);
  const totalStay = classes.reduce((a,c)=>a+c.list.length-goCount(c),0);
  const validCutoff = /^\d{4}-\d{2}-\d{2}$/.test(cutoff);
  const chunk = (a,n=80) => { const o=[]; for (let i=0;i<a.length;i+=n) o.push(a.slice(i,i+n)); return o; };
  const put   = (key, value) => supabase.from("app_settings").upsert({ key, value }, { onConflict:"key" });

  // Switches the app to the new year. Returns an error message, or "" when everything saved.
  async function writeYear(status) {
    try {
      for (const [k,v] of [["current_year",next],["late_cutoff",cutoff],["next_term_begins",""],["rollover_state",JSON.stringify({ from:cur, to:next, status, at:new Date().toISOString() })]]) {
        const { error } = await put(k, v); if (error) throw error;
      }
      await loadAll();
      return "";
    } catch(e) { return e.message || "unknown error"; }
  }

  async function run() {
    setStep("running");
    const failures = [];
    let moved = 0, graduated = 0;
    const continuing = [];
    // Written first so an interrupted run is noticed next time (stops a second run from moving classes up twice)
    try { const { error } = await put("rollover_state", JSON.stringify({ from:cur, to:next, status:"running", at:new Date().toISOString() })); if (error) throw error; }
    catch(e) { alert("Could not start the rollover (nothing was changed): "+e.message); setStep("plan"); return; }
    const jobs = classes.filter(c => c.list.length);
    setProgress({ done:0, total:jobs.length+1 });
    for (const c of jobs) {          // Form 5 first, then Form 4 → 5, … Form 1 → 2
      const go   = c.list.filter(s=>!excluded[s.id]).map(s=>s.id);
      const stay = c.list.filter(s=> excluded[s.id]).map(s=>s.id);
      continuing.push(...stay);
      try {
        for (const ids of chunk(go)) {
          const q = c.to ? supabase.from("students").update({ form:c.to }) : supabase.from("students").update({ active:false, graduated:true });
          const { error } = await q.in("id", ids); if (error) throw error;
        }
        if (c.to) { continuing.push(...go); moved += go.length; } else graduated += go.length;
      } catch(e) { failures.push(`${c.from} ${c.to?"promotion":"graduation"}: ${e.message}`); }
      setProgress(p => ({ ...p, done:p.done+1 }));
    }
    if (resetFees && continuing.length) {
      try { for (const ids of chunk(continuing)) { const { error } = await supabase.from("fees").update({ paid:0 }).in("student_id", ids); if (error) throw error; } }
      catch(e) { failures.push("Fee reset: "+e.message); }
    }
    const yErr = await writeYear(failures.length ? "partial" : "done");
    if (yErr) failures.push("Saving the new year: "+yErr);
    await loadAll();
    setProgress(p => ({ ...p, done:p.total }));
    setResult({ moved, graduated, stayed:totalStay, failures, yearOk:!yErr });
    setStep("done");
  }
  async function onlySetYear() {
    const err = await writeYear("done");
    if (err) alert("Could not save: "+err); else { alert(`The app is now on ${next}.`); onClose(); }
  }

  const box = { borderRadius:8, padding:"9px 12px", marginBottom:12, fontSize:12 };
  return (
    <Modal title={`🔄 New academic year${next?` ${next}`:""}`} onClose={step==="running"?()=>{}:onClose}>
      {!next && <div style={{...box,background:"#fffbeb",border:"1px solid #fde68a",color:"#92400e"}}>The app is on {cur}, the last year in the list. Add the next year to ACAD_YEARS at the top of App.jsx first.</div>}

      {next && step==="plan" && (
        <div>
          <div style={{fontSize:13,fontWeight:800,color:C.navy,marginBottom:8}}>{cur} → {next}</div>
          {interrupted && (
            <div style={{...box,background:"#fef2f2",border:"1px solid #fca5a5",color:C.red}}>
              A rollover to {next} started on {fmtDate((prev.at||"").slice(0,10))} and did not finish. Some students may already have moved up. Check the Students page before running it again.
              <div style={{marginTop:6}}><SmBtn onClick={onlySetYear} color={C.navyMid}>Students are already moved — just switch to {next}</SmBtn></div>
              <label style={{display:"flex",alignItems:"center",gap:6,marginTop:8,color:C.navy}}><input type="checkbox" checked={ack} onChange={e=>setAck(e.target.checked)}/>I checked, run it again from the start</label>
            </div>
          )}
          <div style={{fontSize:11,color:C.gray,marginBottom:6}}>Ticked students move up. Untick a student who is repeating the class. Form 5 students who stay are not graduated.</div>
          <div style={{border:`1px solid ${C.grayLight}`,borderRadius:8,marginBottom:12}}>
            {classes.map(c => (
              <div key={c.from} style={{borderBottom:`1px solid ${C.grayBg}`}}>
                <div onClick={()=>setOpen(open===c.from?"":c.from)} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 10px",cursor:"pointer",gap:8}}>
                  <div>
                    <div style={{fontWeight:700,color:C.navy,fontSize:12.5}}>{c.from} → {c.to || "Graduate"}</div>
                    {c.notReg>0 && <div style={{fontSize:10,color:C.gold}}>{c.notReg} not registered, left as they are</div>}
                  </div>
                  <div style={{fontSize:11.5,fontWeight:700,color:C.green,textAlign:"right"}}>{goCount(c)} {c.to?"move up":"graduate"}{c.list.length-goCount(c)>0 && <span style={{color:C.gold}}> · {c.list.length-goCount(c)} stay</span>}</div>
                </div>
                {open===c.from && (
                  <div style={{maxHeight:200,overflowY:"auto",padding:"0 10px 8px"}}>
                    {!c.list.length && <div style={{fontSize:12,color:C.gray}}>No registered students.</div>}
                    {c.list.map(st => (
                      <label key={st.id} style={{display:"flex",alignItems:"center",gap:8,padding:"4px 0",fontSize:12,color:excluded[st.id]?C.gray:C.navy}}>
                        <input type="checkbox" checked={!excluded[st.id]} onChange={e=>setExcluded(x=>({...x,[st.id]:!e.target.checked}))}/>
                        <span style={{flex:1,textDecoration:excluded[st.id]?"line-through":"none"}}>{st.name}</span>
                        <span style={{fontSize:10,fontFamily:"monospace",color:C.gold}}>{st.id}</span>
                      </label>
                    ))}
                  </div>
                )}
              </div>
            ))}
          </div>
          <label style={{display:"flex",alignItems:"center",gap:8,fontSize:12,color:C.navy,marginBottom:10}}><input type="checkbox" disabled={!ctx.auth.canFees} checked={resetFees} onChange={e=>setResetFees(e.target.checked)}/>Reset fees to 0 paid for everyone who continues (moved up or repeating)</label>
          <Fr label={`Last day for normal registration in ${next} (later registrations pay the late fee)`}><input style={inp} type="date" value={cutoff} onChange={e=>setCutoff(e.target.value)}/></Fr>
          <div style={{...box,background:C.grayBg,color:C.navy,lineHeight:1.5}}>
            <b>What changes:</b> classes move up, Form 5 graduates, the app switches to {next} (First Term, SEQ 1), and the old "next term begins" date is cleared. Marks entry for {next} is open until you set deadlines in Marks Control.<br/>
            <b>What stays:</b> all marks, remarks and attendance of {cur} under that year; teachers' classes, subjects and class teachers (they belong to the class); subject lists and candidate numbers of graduating Form 5 candidates, kept for their GCE records.<br/>
            <b>After this:</b> set the subject list and candidate number of the new Form 5 in Marks Control → Form 5 → Setup, and register the new Form 1 intake in Registration.
          </div>
          <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
            <Btn onClick={onClose} outline>Cancel</Btn>
            <Btn onClick={()=>setStep("confirm")} disabled={!validCutoff || (interrupted && !ack) || (!totalGo && !totalStay)} color={C.gold}>Review →</Btn>
          </div>
        </div>
      )}

      {next && step==="confirm" && (
        <div>
          <div style={{...box,background:"#fef2f2",border:"1px solid #fca5a5",textAlign:"center"}}>
            <div style={{fontSize:26,marginBottom:4}}>⚠️</div>
            <div style={{fontWeight:800,color:C.red,fontSize:14,marginBottom:4}}>This cannot be undone in one step</div>
            <div style={{color:C.gray}}>{totalGo} student(s) move up or graduate{totalStay>0?`, ${totalStay} stay in their class`:""}{resetFees?", fees reset to 0":""}. The app switches to {next}.</div>
          </div>
          <Fr label={`Type ${next} to confirm`}><input style={inp} value={typed} onChange={e=>setTyped(e.target.value.trim())} placeholder={next}/></Fr>
          <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:10}}>
            <Btn onClick={()=>setStep("plan")} outline>← Back</Btn>
            <Btn onClick={run} disabled={typed!==next} color={C.red}>Start {next}</Btn>
          </div>
        </div>
      )}

      {step==="running" && (
        <div style={{textAlign:"center",padding:"20px 0"}}>
          <div style={{background:C.grayBg,borderRadius:6,height:10,overflow:"hidden",marginBottom:10}}><div style={{width:`${progress.done/Math.max(progress.total,1)*100}%`,height:"100%",background:C.gold,transition:"width .3s"}}/></div>
          <div style={{fontSize:13,color:C.gray}}>Working… keep this screen open ({progress.done}/{progress.total})</div>
        </div>
      )}

      {step==="done" && result && (
        <div>
          <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
            <Pill color={C.green}>✓ {result.moved} moved up</Pill>
            <Pill color={C.green}>🎓 {result.graduated} graduated</Pill>
            {result.stayed>0 && <Pill color={C.gold}>{result.stayed} stayed</Pill>}
            {result.yearOk && <Pill color={C.navyMid}>Now on {next}</Pill>}
            {result.failures.length>0 && <Pill color={C.red}>✕ {result.failures.length} problem(s)</Pill>}
          </div>
          {result.failures.length>0 && (
            <div style={{...box,background:"#fef2f2",border:"1px solid #fca5a5",color:C.red}}>
              {result.failures.map((f,i)=><div key={i} style={{marginBottom:3}}>{f}</div>)}
              <div style={{color:C.gray,marginTop:6}}>Do not run this again. Fix the classes listed above with 🎓 Promote Students, one class at a time.</div>
              {!result.yearOk && <div style={{marginTop:8}}><SmBtn onClick={async()=>{ const err=await writeYear("done"); if(err) alert("Could not save: "+err); else { alert(`The app is now on ${next}.`); setResult(r=>({...r,yearOk:true})); } }} color={C.navyMid}>Retry switching to {next}</SmBtn></div>}
            </div>
          )}
          <Btn onClick={onClose}>Done</Btn>
        </div>
      )}
    </Modal>
  );
}

// ─── Student ID cards: 10 per A4 page (credit-card size) ───────────────────────
function idCardDocs(list, year) {
  const card = st => {
    const photo = st.photo_url
      ? `<img src="${st.photo_url}" alt="" style="width:20mm;height:25mm;object-fit:cover;border:0.3mm solid #0D2340;display:block">`
      : `<div style="width:20mm;height:25mm;border:0.3mm dashed #999;box-sizing:border-box;display:flex;align-items:center;justify-content:center;font-size:6pt;color:#888;text-align:center">Photo</div>`;
    return `<div class="c">
      <div class="top"><img src="${SCHOOL_LOGO}" alt="" style="width:9mm;height:9mm;object-fit:contain"><div><div class="sc">SAKER BAPTIST COMPREHENSIVE COLLEGE BAWE</div><div class="tt">Student ID Card · ${escHtml(year)}</div></div></div>
      <div class="bd">${photo}<div class="in"><div class="nm">${escHtml(st.name)}</div><div class="ln"><b>Matricule:</b> ${escHtml(st.id)}</div><div class="ln"><b>Class:</b> ${escHtml(st.form)}</div>${st.gender?`<div class="ln"><b>Sex:</b> ${escHtml(st.gender)}</div>`:""}</div></div>
      <div class="ft"><span>Holder's signature</span><span>Principal</span></div>
    </div>`;
  };
  const pages = [];
  for (let i=0; i<list.length; i+=10) pages.push(list.slice(i,i+10));
  return pages.map(g => `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    body{margin:0;padding:8mm 15mm;font-family:Arial,Helvetica,sans-serif;color:#111;background:#fff}
    .grid{display:grid;grid-template-columns:85.6mm 85.6mm;grid-auto-rows:54mm;gap:2.5mm 4mm}
    .c{border:0.3mm solid #0D2340;border-radius:2.5mm;box-sizing:border-box;overflow:hidden;display:flex;flex-direction:column}
    .top{background:#0D2340;color:#fff;display:flex;align-items:center;gap:2mm;padding:1mm 2.5mm;height:11mm;box-sizing:border-box;border-bottom:0.8mm solid #C9962A}
    .top img{background:#fff;border-radius:50%} .sc{font-size:6.3pt;font-weight:700} .tt{font-size:6pt;color:#F0C050}
    .bd{display:flex;gap:3mm;padding:2.5mm 3mm;flex:1} .in{flex:1;min-width:0}
    .nm{font-weight:700;font-size:9.5pt;margin-bottom:1.5mm;line-height:4.2mm} .ln{font-size:7.5pt;margin-bottom:0.8mm}
    .ft{display:flex;justify-content:space-between;padding:0 3mm 2mm;font-size:5.5pt;color:#555} .ft span{border-top:0.2mm solid #555;width:38%;text-align:center;padding-top:0.4mm}
  </style></head><body><div class="grid">${g.map(card).join("")}</div></body></html>`);
}

function IdCardsModal({ students, onClose }) {
  const [form, setForm] = useState("");
  const [busy, setBusy] = useState("");
  const list = students.filter(s => s.active && s.reg_status==="registered" && (!form || s.form===form))
    .sort((a,b)=>FORMS.indexOf(a.form)-FORMS.indexOf(b.form) || a.name.localeCompare(b.name));
  const noPhoto = list.filter(s=>!s.photo_url).length;
  async function print() {
    if (!list.length) { alert("No students to print."); return; }
    const docs = idCardDocs(list, ACTIVE_YEAR);
    setBusy(`0/${docs.length}`);
    try { downloadBlob(await htmlDocsToPdfBlob(docs, (d,t)=>setBusy(`${d}/${t}`), 1.5, 0.85), `ID-Cards-${(form||"All-classes").replace(/\s+/g,"")}-${ACTIVE_YEAR.replace("/","-")}.pdf`); }
    catch(e) { alert("Could not create the PDF: "+e.message); }
    setBusy("");
  }
  return (
    <Modal title="🪪 Student ID cards" onClose={onClose}>
      <Fr label="Class">
        <select style={inp} value={form} onChange={e=>setForm(e.target.value)}><option value="">All classes</option>{FORMS.map(f=><option key={f}>{f}</option>)}</select>
      </Fr>
      <div style={{fontSize:12,color:C.navy,marginBottom:4}}>{list.length} card(s) · {Math.ceil(list.length/10)} page(s) · year {ACTIVE_YEAR}</div>
      {noPhoto>0 && <div style={{fontSize:11,color:C.gold,marginBottom:8}}>{noPhoto} student(s) have no photo yet. Their card prints with an empty photo box.</div>}
      <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:10}}>
        <Btn onClick={onClose} outline>Close</Btn>
        <Btn onClick={print} disabled={!!busy}>{busy?`Making PDF… ${busy}`:"Print ID cards PDF"}</Btn>
      </div>
    </Modal>
  );
}

// ─── Teachers ──────────────────────────────────────────────────────────────────
// Name of the Supabase Edge Function that creates / resets / deletes teacher logins
const TEACHER_FUNCTION = "manage-teacher-v2";

// Turns a failed call to the manage-teacher function into a message an admin can act on
async function functionErrorMessage(error) {
  try { const body = await error.context.json(); if (body?.error) return body.error; } catch { /* not JSON */ }
  const status = error?.context?.status;
  if (status === 404 || /failed to send|failed to fetch/i.test(error?.message || ""))
    return `The teacher-account service (${TEACHER_FUNCTION}) could not be reached. It may not be deployed in Supabase yet.`;
  return error?.message || "Unknown error";
}

function TeachersPage({ ctx }) {
  const { teachers, saveTeacher } = ctx;
  const [modal,  setModal]  = useState(null);
  const [form,   setForm]   = useState({ name:"", email:"", password:"", assignments:{}, active:true });
  const [saving, setSaving] = useState(false);
  const [err,    setErr]    = useState("");
  const [reset,  setReset]  = useState(null); // { t, password, done }

  // Per-class assignments: { "Form 1": ["Mathematics"], "Form 3": ["Physics"] }
  const setAssign   = fn => setForm(f => ({ ...f, assignments: fn({ ...(f.assignments||{}) }) }));
  const toggleClass = fm => setAssign(a => { if (a[fm]) delete a[fm]; else a[fm] = []; return a; });
  const toggleSub   = (fm, sub) => setAssign(a => { const cur = a[fm]||[]; a[fm] = cur.includes(sub) ? cur.filter(x=>x!==sub) : [...cur, sub]; return a; });
  const cleanAssign = a => Object.fromEntries(Object.entries(a||{}).filter(([,v]) => v.length));
  const derive      = a => ({ forms: FORMS.filter(f=>a[f]), subjects: [...new Set(Object.values(a).flat())] });

  function randomPassword() {
    // Easy-to-read 8-character password: avoids ambiguous characters (0/O, 1/l/I)
    const chars = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789";
    let p = "";
    for (let i=0;i<8;i++) p += chars[Math.floor(Math.random()*chars.length)];
    return p;
  }

  async function createTeacher() {
    if (!form.name?.trim()||!form.email?.trim()||!form.password) { setErr("Please fill in the name, email and password."); return; }
    if (form.password.length < 6) { setErr("The password must be at least 6 characters."); return; }
    setErr(""); setSaving(true);
    try {
      const a = cleanAssign(form.assignments), d = derive(a);
      const { data, error } = await supabase.functions.invoke(TEACHER_FUNCTION, {
        body: { action:"create", name:form.name.trim(), email:form.email.trim().toLowerCase(), password:form.password, subjects:d.subjects, forms:d.forms },
      });
      if (error) throw new Error(await functionErrorMessage(error));
      if (data?.error) throw new Error(data.error);
      const extra = { ...(Object.keys(a).length ? { assignments:a } : {}), ...(form.class_teacher_of ? { class_teacher_of:form.class_teacher_of } : {}) };
      if (Object.keys(extra).length) {
        const { error: e2 } = await supabase.from("teachers").update(extra).ilike("email", form.email.trim());
        if (e2) throw new Error("Account created, but saving the classes failed: "+e2.message+". Run migration2.sql in Supabase, then edit this teacher.");
      }
      setModal({ justCreated:true, name:form.name, email:form.email.trim().toLowerCase(), password:form.password });
    } catch(e) { setErr("Could not create teacher: " + e.message); }
    setSaving(false);
  }

  async function resetPassword() {
    if (!reset.password || reset.password.length < 6) { setErr("The password must be at least 6 characters."); return; }
    setErr(""); setSaving(true);
    try {
      const { data, error } = await supabase.functions.invoke(TEACHER_FUNCTION, {
        body: { action:"reset_password", teacherId:reset.t.id, userId:reset.t.user_id, newPassword:reset.password },
      });
      if (error) throw new Error(await functionErrorMessage(error));
      if (data?.error) throw new Error(data.error);
      setReset(r => ({ ...r, done:true }));
    } catch(e) { setErr("Could not reset the password: " + e.message); }
    setSaving(false);
  }

  function copyLogin(email, password) {
    const text = `SBCC Bawe school portal\nEmail: ${email}\nPassword: ${password}\n${window.location.origin}`;
    navigator.clipboard?.writeText(text).then(() => alert("Login details copied.")).catch(() => alert("Could not copy. Please write them down."));
  }

  async function editTeacher() {
    if (!form.name?.trim()) return;
    setSaving(true);
    try { const a = cleanAssign(form.assignments); await saveTeacher({ id:modal.id, name:form.name, email:form.email, ...derive(a), assignments:a, ...("class_teacher_of" in form ? { class_teacher_of: form.class_teacher_of||null } : {}), active:form.active, joined:form.joined }); setModal(null); }
    catch(e) { setErr("Error: "+e.message); }
    setSaving(false);
  }

  async function deleteTeacherFully(t) {
    if (!window.confirm(`Permanently delete ${t.name}'s account? They will no longer be able to log in. This cannot be undone.`)) return;
    try {
      const { data, error } = await supabase.functions.invoke(TEACHER_FUNCTION, {
        body: { action:"delete", teacherId:t.id, userId:t.user_id },
      });
      if (error) throw new Error(await functionErrorMessage(error));
      if (data?.error) throw new Error(data.error);
      alert(`${t.name}'s account has been removed.`);
    } catch(e) { alert("Could not delete: " + e.message); }
  }

  return (
    <div>
      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:9,gap:8}}>
        <Btn onClick={()=>{setForm({name:"",email:"",password:randomPassword(),assignments:{},active:true,joined:todayStr()});setErr("");setModal("add");}}>+ Create Teacher Account</Btn>
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
          {t.class_teacher_of && <div style={{marginTop:6,fontSize:11.5,color:C.green,fontWeight:700}}>🗒️ Class teacher of {t.class_teacher_of}</div>}
          {FORMS.filter(f=>teacherSubjectsIn(t,f).length).map(f=>(
            <div key={f} style={{marginTop:6,display:"flex",flexWrap:"wrap",gap:4,alignItems:"center"}}>
              <Pill color={C.gold}>{f}</Pill>
              {teacherSubjectsIn(t,f).map(sb=><Pill key={sb} color={C.navyMid}>{sb}</Pill>)}
            </div>
          ))}
          <div style={{display:"flex",gap:7,marginTop:8,flexWrap:"wrap"}}>
            <SmBtn onClick={()=>{setForm({...t,password:"",assignments:teacherAssignments(t)});setErr("");setModal(t);}} color={C.green}>Edit</SmBtn>
            <SmBtn onClick={()=>{setErr("");setReset({t,password:randomPassword(),done:false});}} color={C.navyMid}>Reset Password</SmBtn>
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
          <div style={{display:"flex",gap:8}}><Btn outline onClick={()=>copyLogin(modal.email,modal.password)}>Copy login details</Btn><Btn onClick={()=>setModal(null)}>Done</Btn></div>
        </Modal>
      )}

      {reset && (
        <Modal title={reset.done ? "Password reset" : `Reset password · ${reset.t.name}`} onClose={()=>!saving&&setReset(null)}>
          {reset.done ? (<>
            <div style={{background:"#f0fdf4",border:"1px solid #86efac",borderRadius:10,padding:12,marginBottom:12,fontSize:13,color:C.green,fontWeight:700}}>New password saved for {reset.t.name}.</div>
            <div style={{background:C.grayBg,borderRadius:8,padding:"10px 12px",marginBottom:12,fontSize:13}}>
              <div>Email: <strong>{reset.t.email}</strong></div>
              <div>Password: <strong style={{fontFamily:"monospace"}}>{reset.password}</strong></div>
            </div>
            <div style={{display:"flex",gap:8}}><Btn outline onClick={()=>copyLogin(reset.t.email,reset.password)}>Copy login details</Btn><Btn onClick={()=>setReset(null)}>Done</Btn></div>
          </>) : (<>
            <Fr label="New password">
              <div style={{display:"flex",gap:6}}>
                <input style={{...inp,flex:1,fontFamily:"monospace"}} value={reset.password} onChange={e=>setReset(r=>({...r,password:e.target.value}))}/>
                <SmBtn onClick={()=>setReset(r=>({...r,password:randomPassword()}))} color={C.navyMid}>New</SmBtn>
              </div>
            </Fr>
            {err && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"8px 12px",color:C.red,fontSize:12,marginTop:10}}>{err}</div>}
            <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:13}}>
              <Btn outline onClick={()=>setReset(null)} disabled={saving}>Cancel</Btn>
              <Btn onClick={resetPassword} disabled={saving}>{saving?"Saving…":"Save new password"}</Btn>
            </div>
          </>)}
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
            <label style={lbl}>Classes taught — tap a class, then tick only the subjects taught in it</label>
            <div style={{display:"flex",flexWrap:"wrap",gap:5,marginTop:5}}>
              {FORMS.map(f => { const on = !!form.assignments?.[f]; return (
                <button key={f} onClick={()=>toggleClass(f)} style={{padding:"5px 10px",borderRadius:12,border:`1px solid ${on?C.gold:C.grayLight}`,background:on?C.gold:C.white,color:on?C.white:C.gray,fontSize:12,fontWeight:700,cursor:"pointer"}}>{f}</button>
              ); })}
            </div>
            {FORMS.filter(f=>form.assignments?.[f]).map(f => (
              <div key={f} style={{marginTop:9,padding:"8px 10px",background:C.grayBg,borderRadius:8}}>
                <div style={{fontSize:11.5,fontWeight:800,color:C.navy,marginBottom:5}}>{f}</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
                  {SUBJECTS_BY_FORM[f].map(sb => { const on = form.assignments[f].includes(sb); return (
                    <button key={sb} onClick={()=>toggleSub(f,sb)} style={{padding:"4px 8px",borderRadius:12,border:`1px solid ${on?C.navy:C.grayLight}`,background:on?C.navy:C.white,color:on?C.white:C.gray,fontSize:11,cursor:"pointer"}}>{sb}</button>
                  ); })}
                </div>
              </div>
            ))}
          </div>
          <div style={{marginTop:10}}>
            <Fr label="Class teacher of (writes conduct & remarks for that form)">
              <select style={inp} value={form.class_teacher_of||""} onChange={e=>setForm(f=>({...f,class_teacher_of:e.target.value||null}))}>
                <option value="">— none —</option>{FORMS.map(f=><option key={f}>{f}</option>)}
              </select>
            </Fr>
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

// Shown to a teacher whose account is not linked to a profile or has no classes/subjects assigned yet
function TeacherSetupNotice({ teacher, email, needSubjects }) {
  let msg = "";
  if (!teacher) msg = `Your login (${email}) is not linked to a teacher profile. Ask the administrator to open Teachers and check that your email matches.`;
  else if (!(teacher.forms||[]).length) msg = "No classes are assigned to your profile yet. Ask the administrator to open Teachers, edit your profile and choose your classes and subjects.";
  else if (needSubjects && !(teacher.subjects||[]).length) msg = "No subjects are assigned to your profile yet. Ask the administrator to open Teachers, edit your profile and tick your subjects.";
  if (!msg) return null;
  return <div style={{background:"#fffbeb",border:"1px solid #fcd34d",borderRadius:10,padding:12,marginBottom:11,fontSize:12.5,color:"#78350f"}}><strong>⚠️ Setup needed.</strong> {msg}</div>;
}

// ─── Marks Entry ───────────────────────────────────────────────────────────────
function MarksPage({ ctx }) {
  const { students, marksMap, saveMark, auth, teachers, loadAll, markDeadlines, markUnlocks } = ctx;
  const isTeacher  = auth.role==="teacher";
  const teacher    = findTeacher(teachers, auth.user.email);
  const availForms = isTeacher ? (teacher?.forms||[]) : FORMS;

  const [filter, setFilter] = useState({ form:availForms[0]||(isTeacher?"":"Form 1"), seq:"SEQ 1", year:ACTIVE_YEAR, subject:"" });
  // If the teacher profile loads after this page opened, or the chosen class is not theirs, jump to their first class
  useEffect(() => {
    if (isTeacher && availForms.length && !availForms.includes(filter.form)) setFilter(f => ({ ...f, form:availForms[0], subject:"" }));
  }, [availForms.join("|")]);
  // The assessment list depends on the class: Form 5 has Pre Mock / Regional Mock / GCE Result
  const seqOptions = seqsForForm(filter.form);
  const seqV       = seqOptions.includes(filter.seq) ? filter.seq : seqOptions[0];
  const availSubs  = isTeacher
    ? teacherSubjectsIn(teacher, filter.form).filter(s=>SUBJECTS_BY_FORM[filter.form]?.includes(s))
    : (SUBJECTS_BY_FORM[filter.form]||[]);
  const subject    = filter.subject || availSubs[0] || "";
  const lock       = lockStatus(filter.form, seqV, filter.year, markDeadlines, markUnlocks);
  const gceAdminOnly = isTeacher && seqV===GCE_SEQ; // GCE grades are decided and entered by the admin
  const blocked    = (isTeacher && lock.locked) || gceAdminOnly;
  const paperMode  = isF5(filter.form) && seqV!==GCE_SEQ; // Form 5 mocks: enter each paper, score /20 is worked out

  const [coeff,  setCoeff]  = useState(getCoeff(subject));
  const [local,  setLocal]  = useState({});
  const [saving, setSaving] = useState(false);
  const [saved,  setSaved]  = useState(false);
  const [paperLocal, setPaperLocal] = useState({}); // studentId → {P1,P2,P3}

  useEffect(()=>setCoeff(getCoeff(subject)),[subject]);

  // Form 5: only the students who write this subject
  const formStudents = students.filter(s=>s.active&&s.form===filter.form&&s.reg_status==="registered"&&f5Writes(ctx.settings, s, subject));

  useEffect(() => {
    const init = {};
    formStudents.forEach(s => {
      const key = `${s.id}-${subject}-${seqV}-${filter.year}`;
      init[s.id] = marksMap[key]?.score ?? "";
    });
    const pinit = {};
    formStudents.forEach(s => {
      const o = {};
      f5PapersFor(subject).forEach(pp => { o[pp.k] = marksMap[`${s.id}-${subject}-${seqV} ${pp.k}-${filter.year}`]?.score ?? ""; });
      pinit[s.id] = o;
    });
    setLocal(init); setPaperLocal(pinit); setSaved(false);
  }, [filter.form, subject, seqV, filter.year, students.length]);

  async function saveAll() {
    if (blocked) return;
    setSaving(true);
    try {
      const rows = [];
      const mk = (sid, sq, score) => rows.push({ student_id:sid, subject, seq:sq, acad_year:filter.year, score, coeff:Number(coeff)||1, teacher_id:teacher?.id||null });
      for (const s of formStudents) {
        if (paperMode) {
          const pv = paperLocal[s.id] || {};
          for (const pp of f5PapersFor(subject)) {
            const x = pv[pp.k];
            if (x===""||x===undefined) continue;
            if (Number(x)<0 || Number(x)>pp.max) throw new Error(`${s.name}: Paper ${pp.k.slice(1)} must be between 0 and ${pp.max}`);
            mk(s.id, `${seqV} ${pp.k}`, Number(x));
          }
          const comb = f5Combine(subject, pv); // only once every paper is entered
          if (comb!==null) mk(s.id, seqV, comb);
          continue;
        }
        const v = local[s.id];
        if (v===""||v===undefined) continue;
        mk(s.id, seqV, Number(v));
      }
      const res = await writeMarkRows(rows, ctx.isOnline, auth.user.name, {});
      if (res.queued) { ctx.setOfflineQueue(loadOfflineQueue()); ctx.applyQueuedMarks(); alert("No internet. The marks are saved on this phone and will upload by themselves when you are back online."); }
      else await ctx.loadAll(); // refresh all data
      setSaved(true);
    } catch(e) { alert("Save error: "+e.message); }
    setSaving(false);
  }

  return (
    <div>
      {isTeacher && <TeacherSetupNotice teacher={teacher} email={auth.user.email} needSubjects/>}
      {(lock.deadline || gceAdminOnly) && (
        <div style={{background:blocked?"#fef2f2":"#fffbeb",border:`1px solid ${blocked?"#fca5a5":"#fde68a"}`,borderRadius:8,padding:"8px 12px",marginBottom:11,fontSize:12,fontWeight:600,color:blocked?C.red:"#92400e"}}>
          {gceAdminOnly ? "🔒 GCE grades are entered by the administrator only."
            : blocked ? `🔒 ${seqV} marks for ${filter.form} closed on ${fmtDate(lock.deadline)}. Ask the administrator to unlock this class if a correction is needed.`
            : lock.locked ? `🔒 Deadline passed (${fmtDate(lock.deadline)}). Teachers are locked out; you can still edit.`
            : lock.unlocked && lock.passed ? `🔓 Unlocked for corrections (deadline was ${fmtDate(lock.deadline)}).`
            : `⏰ Deadline for ${seqV}: ${fmtDate(lock.deadline)}.`}
        </div>
      )}
      <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><label style={lbl}>Form</label><select style={inp} value={filter.form} onChange={e=>setFilter(f=>({...f,form:e.target.value,subject:""}))}>{availForms.map(f=><option key={f}>{f}</option>)}</select></div>
          <div><label style={lbl}>{isF5(filter.form)?"Assessment":"Sequence"}</label><select style={inp} value={seqV} onChange={e=>setFilter(f=>({...f,seq:e.target.value}))}>{seqOptions.map(s=><option key={s}>{s}</option>)}</select></div>
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
          <div style={{fontSize:12}}><strong style={{color:C.navy}}>{filter.form}</strong><span style={{color:C.gray,marginLeft:5}}>· {subject} · {seqV} · Coeff {coeff}</span></div>
          <div style={{display:"flex",gap:7,alignItems:"center"}}>
            {saved && <span style={{fontSize:12,color:C.green,fontWeight:700}}>✓ Saved to Supabase</span>}
            <Btn onClick={saveAll} disabled={saving||blocked}>{saving?"Saving…":blocked?"Locked":"Save Marks"}</Btn>
          </div>
        </div>
        <div style={{padding:"5px 11px",background:"#fffbeb",borderBottom:`1px solid ${C.grayLight}`,fontSize:11,color:"#92400e"}}>
          {seqV===GCE_SEQ ? <>Select the <strong>GCE grade</strong> (A, B, C = pass · D, E, U = fail)</> : paperMode ? <>{f5PaperNote(subject)} → score out of <strong>20</strong> · Weighted = score × {coeff}</> : <>Score out of <strong>20</strong> · Weighted = score × {coeff} = max {20*coeff} pts</>}
        </div>
        {!formStudents.length ? <Empty text="No registered students in this form."/> : formStudents.map((s,i) => {
          const pv = paperLocal[s.id] || {};
          const v = paperMode ? (f5Combine(subject, pv) ?? "") : (local[s.id] ?? "");
          const gceG = seqV===GCE_SEQ ? gceFromPoints(v) : "";
          const { grade: gradeRaw } = scoreToGrade(v);
          const f5g = isF5(filter.form) && seqV!==GCE_SEQ ? f5Grade(ctx.settings, v) : null; // admin's Form 5 scale
          const grade = seqV===GCE_SEQ ? (gceG||"—") : f5g ? f5g.grade : gradeRaw;
          const gc2 = seqV===GCE_SEQ ? (gceG==="" ? C.gray : gcePassed(v) ? C.green : C.red)
                    : f5g ? (f5g.grade==="—" ? C.gray : f5g.pass ? C.green : C.red) : gradeCol(grade);
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
                  <div style={{fontSize:9,color:C.gray,marginBottom:2}}>{seqV===GCE_SEQ?"GCE grade":paperMode?"Papers → /20":"Score /20"}</div>
                  {seqV===GCE_SEQ
                    ? <select disabled={blocked} style={{width:62,padding:"5px",border:`1px solid ${C.border}`,borderRadius:7,fontSize:14,fontWeight:700,textAlign:"center",color:C.navy}} value={gceFromPoints(v)} onChange={e=>{setLocal(p=>({...p,[s.id]:gceToPoints(e.target.value)}));setSaved(false);}}><option value="">—</option>{GCE_GRADES.map(g=><option key={g} value={g}>{g}</option>)}</select>
                    : paperMode
                    ? <div style={{display:"flex",alignItems:"flex-end",gap:4}}>
                        {f5PapersFor(subject).map(pp => (
                          <div key={pp.k} style={{textAlign:"center"}}>
                            <div style={{fontSize:8,color:C.gray}}>P{pp.k.slice(1)} /{pp.max}</div>
                            <input type="number" min={0} max={pp.max} disabled={blocked} style={{width:44,padding:"5px 2px",border:`1px solid ${C.border}`,borderRadius:7,fontSize:13,fontWeight:700,textAlign:"center",color:C.navy}} value={pv[pp.k] ?? ""} onChange={e=>{const val=e.target.value;setPaperLocal(x=>({...x,[s.id]:{...(x[s.id]||{}),[pp.k]:val}}));setSaved(false);}}/>
                          </div>
                        ))}
                        <div style={{fontSize:12,fontWeight:800,color:C.navyMid,minWidth:42,textAlign:"center"}}>{v!=="" ? Number(v).toFixed(2) : "—"}</div>
                      </div>
                    : <input type="number" min={0} max={20} disabled={blocked} style={{width:52,padding:"5px",border:`1px solid ${C.border}`,borderRadius:7,fontSize:14,fontWeight:700,textAlign:"center",color:C.navy}} value={v} onChange={e=>{setLocal(p=>({...p,[s.id]:e.target.value}));setSaved(false);}}/>}
                </div>
                <div style={{textAlign:"center",minWidth:36}}>
                  <div style={{fontSize:9,color:C.gray,marginBottom:2}}>{seqV===GCE_SEQ?"":"×"+coeff}</div>
                  <span style={{fontSize:12,fontWeight:700,color:C.navyMid}}>{seqV===GCE_SEQ?"":(v!==""?Number(v)*coeff:"—")}</span>
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

// ─── Marks Control (admin): who has entered marks, deadlines & locks, change history ───
// Admin: Form 5 grade scale + the subjects (and candidate number) of each Form 5 student
function F5SetupPanel({ ctx }) {
  const { students, settings, saveSetting } = ctx;
  const scale = f5Scale(settings);
  const [sc, setSc]   = useState({ A:String(scale.A), B:String(scale.B), C:String(scale.C), D:String(scale.D), E:String(scale.E) });
  const [busy, setBusy] = useState("");
  const [open, setOpen] = useState("");
  const [draft, setDraft] = useState({ subjects:[], candNo:"" });
  const [dlDate, setDlDate] = useState(String(settings?.f5_subject_deadline||""));
  const [showMissing, setShowMissing] = useState(false);
  const [bulkSubs, setBulkSubs] = useState([]);
  const [bulkSel,  setBulkSel]  = useState([]);
  const [bulkMode, setBulkMode] = useState("replace"); // replace | add | remove
  const [tmplId,   setTmplId]   = useState("");
  const [cnPrefix, setCnPrefix] = useState("");
  const [cnStart,  setCnStart]  = useState("1");
  const [cnWidth,  setCnWidth]  = useState("3");
  const [cnOnlyEmpty, setCnOnlyEmpty] = useState(true);
  const all = SUBJECTS_BY_FORM["Form 5"];
  const pool = students.filter(x => x.active && x.form==="Form 5" && x.reg_status==="registered").sort((a,b)=>a.name.localeCompare(b.name));
  const cfg = f5Students(settings);

  const win = f5SubjectWindow(settings);
  const submittedMap = f5ParentSubmitted(settings);
  const submittedCount = pool.filter(x => submittedMap[x.id]).length;
  const notSubmitted = pool.filter(x => !submittedMap[x.id]);
  const closedByAdmin = win.reason==="closed_by_admin";
  async function saveDeadline() {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dlDate)) { alert("Pick a closing date first."); return; }
    setBusy("dl");
    try {
      await saveSetting("f5_subject_deadline", dlDate);
      if (closedByAdmin) await saveSetting("f5_subject_closed", "false"); // a new date reopens it
      alert("Closing date saved. Parents of Form 5 students can choose subjects in the Parent Portal until " + fmtDate(dlDate) + ".");
    } catch(e) { alert("Error: "+e.message); }
    setBusy("");
  }
  async function toggleClosed() {
    const closing = !closedByAdmin;
    if (!window.confirm(closing ? "Close subject selection now? No parent will be able to fill or change subjects." : "Reopen subject selection? Parents can fill again until the closing date.")) return;
    setBusy("dl");
    try { await saveSetting("f5_subject_closed", closing ? "true" : "false"); }
    catch(e) { alert("Error: "+e.message); }
    setBusy("");
  }
  async function saveScale() {
    const n = Object.fromEntries(Object.entries(sc).map(([k,v]) => [k, Number(v)]));
    const order = ["A","B","C","D","E"];
    if (order.some(k => !Number.isFinite(n[k]) || n[k]<0 || n[k]>20) || order.some((k,i) => i>0 && n[k] >= n[order[i-1]])) {
      alert("Each grade must start lower than the one above it (A highest), between 0 and 20."); return;
    }
    setBusy("scale");
    try { await saveSetting("f5_grade_scale", JSON.stringify(n)); alert("Form 5 grade scale saved."); }
    catch(e) { alert("Error: "+e.message); }
    setBusy("");
  }
  async function applyBulk() {
    if (!bulkSubs.length) { alert("Tick at least one subject."); return; }
    if (!bulkSel.length)  { alert("Select at least one student."); return; }
    const verb = bulkMode==="replace" ? "set their subjects to exactly these" : bulkMode==="add" ? "add these subjects to their lists" : "remove these subjects from their lists";
    if (!window.confirm(`Apply to ${bulkSel.length} student(s) and ${verb}?`)) return;
    setBusy("bulk");
    try {
      const cur = f5Students(settings), next = { ...cur };
      let skipped = 0;
      bulkSel.forEach(id => {
        const old  = cur[id] || {};
        const base = old.subjects?.length ? old.subjects : all; // no list yet = writes everything
        let subjects = bulkMode==="replace" ? bulkSubs : bulkMode==="add" ? [...new Set([...base, ...bulkSubs])] : base.filter(x => !bulkSubs.includes(x));
        subjects = all.filter(x => subjects.includes(x));
        if (!subjects.length) { skipped++; return; }
        next[id] = { subjects, candNo: old.candNo || "" };
      });
      await saveSetting("f5_students", JSON.stringify(next));
      setBulkSel([]);
      alert(`Done.${skipped ? " "+skipped+" student(s) skipped because they would have no subjects left." : ""}`);
    } catch(e) { alert("Error: "+e.message); }
    setBusy("");
  }
  const cnMake = n => `${cnPrefix.trim()}${String(n).padStart(Math.max(1, Number(cnWidth)||1), "0")}`;
  async function applyCandNos() {
    const start = parseInt(cnStart, 10);
    if (!Number.isFinite(start) || start<0) { alert("Enter a valid starting number."); return; }
    const cur = f5Students(settings);
    const targets = pool.filter(st => !cnOnlyEmpty || !cur[st.id]?.candNo);
    if (!targets.length) { alert("Every student already has a candidate number."); return; }
    if (!window.confirm(`Give ${targets.length} student(s) candidate numbers ${cnMake(start)} to ${cnMake(start+targets.length-1)} (alphabetical order)?`)) return;
    setBusy("cn");
    try {
      const next = { ...cur };
      targets.forEach((st,i) => { const old = cur[st.id] || {}; next[st.id] = { subjects: old.subjects || [], candNo: cnMake(start+i) }; });
      await saveSetting("f5_students", JSON.stringify(next));
      alert("Candidate numbers saved.");
    } catch(e) { alert("Error: "+e.message); }
    setBusy("");
  }
  function openStudent(st) {
    if (open===st.id) { setOpen(""); return; }
    const c = cfg[st.id];
    setDraft({ subjects: c?.subjects?.length ? [...c.subjects] : [...all], candNo: c?.candNo || "" });
    setOpen(st.id);
  }
  async function saveStudent(st) {
    if (!draft.subjects.length) { alert("Tick at least one subject."); return; }
    setBusy(st.id);
    try {
      const next = { ...f5Students(settings), [st.id]: { subjects: draft.subjects, candNo: draft.candNo.trim() } };
      await saveSetting("f5_students", JSON.stringify(next));
      setOpen("");
    } catch(e) { alert("Error: "+e.message); }
    setBusy("");
  }
  const card = { background:C.white, borderRadius:10, padding:12, marginBottom:11, boxShadow:"0 1px 3px rgba(0,0,0,0.06)" };
  return (
    <div>
      <div style={card}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:4}}>Form 5 grade scale (set by admin)</div>
        <div style={{fontSize:11,color:C.gray,marginBottom:8}}>Minimum score out of 20 for each grade. Below E is U. A, B and C are passes. Used for the Pre Mock and Regional Mock.</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:6}}>
          {["A","B","C","D","E"].map(g => (
            <div key={g} style={{textAlign:"center"}}>
              <div style={{fontSize:11,fontWeight:800,color:C.navy,marginBottom:3}}>{g} ≥</div>
              <input type="number" min={0} max={20} step="0.5" style={{...inp,padding:"6px 2px",textAlign:"center"}} value={sc[g]} onChange={e=>setSc(x=>({...x,[g]:e.target.value}))}/>
            </div>
          ))}
        </div>
        <div style={{marginTop:9}}><Btn onClick={saveScale} disabled={busy==="scale"}>{busy==="scale"?"Saving…":"Save grade scale"}</Btn></div>
      </div>

      <div style={card}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:4}}>Parents choose subjects (Parent Portal)</div>
        <div style={{fontSize:11,color:C.gray,marginBottom:8}}>Set a closing date. Until then, parents of Form 5 students pick their child's subjects in the Parent Portal. Once it closes, nobody can fill or change them there. You can still edit any student below.</div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <input type="date" style={{...inp,maxWidth:170}} value={dlDate} onChange={e=>setDlDate(e.target.value)}/>
          <Btn onClick={saveDeadline} disabled={busy==="dl"}>{busy==="dl"?"Saving…":"Save closing date"}</Btn>
        </div>
        <div style={{marginTop:9,display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <Pill color={win.open?C.green:C.red}>{win.open ? `Open until ${fmtDate(win.deadline)}` : win.reason==="closed_by_admin" ? "Closed by admin" : win.reason==="deadline_passed" ? `Closed (${fmtDate(win.deadline)})` : "Not opened — set a date"}</Pill>
          {(win.open || closedByAdmin) && <SmBtn onClick={toggleClosed} color={closedByAdmin?C.green:C.red}>{closedByAdmin?"Reopen":"Close now"}</SmBtn>}
        </div>
        <div style={{fontSize:12,color:C.navy,marginTop:9}}>{submittedCount} of {pool.length} parents have submitted
          {notSubmitted.length>0 && <> · <span onClick={()=>setShowMissing(v=>!v)} style={{color:C.navyMid,textDecoration:"underline",cursor:"pointer"}}>{showMissing?"hide":"show"} {notSubmitted.length} not yet</span></>}
        </div>
        {showMissing && <div style={{maxHeight:160,overflowY:"auto",border:`1px solid ${C.grayLight}`,borderRadius:8,padding:6,marginTop:6,fontSize:12,color:C.navy}}>{notSubmitted.map(x=><div key={x.id} style={{padding:"2px 0"}}>{x.name}</div>)}</div>}
        <div style={{fontSize:10.5,color:C.gray,marginTop:8}}>Needs migration_f5_subjects.sql to have been run once in Supabase. Students who never submit keep the default list until you set it.</div>
      </div>

      <div style={card}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:4}}>Candidate numbers</div>
        <div style={{fontSize:11,color:C.gray,marginBottom:8}}>Numbers are given in alphabetical order of name. You can still edit any single number below.</div>
        <div style={{display:"grid",gridTemplateColumns:"1.3fr 1fr 0.8fr",gap:6}}>
          <div><label style={lbl}>Prefix</label><input style={inp} value={cnPrefix} onChange={e=>setCnPrefix(e.target.value)} placeholder="e.g. 0123"/></div>
          <div><label style={lbl}>Start at</label><input style={inp} type="number" min={0} value={cnStart} onChange={e=>setCnStart(e.target.value)}/></div>
          <div><label style={lbl}>Digits</label><input style={inp} type="number" min={1} max={6} value={cnWidth} onChange={e=>setCnWidth(e.target.value)}/></div>
        </div>
        <div style={{fontSize:11,color:C.navyMid,marginTop:6}}>Example: {cnMake(parseInt(cnStart,10)||0)}, {cnMake((parseInt(cnStart,10)||0)+1)}, {cnMake((parseInt(cnStart,10)||0)+2)} …</div>
        <label style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:C.navy,marginTop:8}}><input type="checkbox" checked={cnOnlyEmpty} onChange={e=>setCnOnlyEmpty(e.target.checked)}/>Only students who don't have a number yet</label>
        <div style={{marginTop:9}}><Btn onClick={applyCandNos} disabled={busy==="cn"}>{busy==="cn"?"Saving…":"Generate numbers"}</Btn></div>
      </div>

      <div style={card}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:4}}>Bulk set subjects</div>
        <div style={{fontSize:11,color:C.gray,marginBottom:8}}>Tick the subjects, select the students (for example all Science students), then apply once.</div>
        <label style={lbl}>Copy ticks from a student (optional)</label>
        <select style={inp} value={tmplId} onChange={e=>{ const id=e.target.value; setTmplId(id); const c=cfg[id]; setBulkSubs(c?.subjects?.length ? [...c.subjects] : id ? [...all] : []); }}>
          <option value="">— None —</option>
          {pool.map(st=><option key={st.id} value={st.id}>{st.name}</option>)}
        </select>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:8}}>
          <label style={{...lbl,marginBottom:0}}>Subjects ({bulkSubs.length})</label>
          <div style={{display:"flex",gap:6}}><SmBtn onClick={()=>setBulkSubs([...all])} color={C.navyMid}>All</SmBtn><SmBtn onClick={()=>setBulkSubs([])} color={C.gray}>None</SmBtn></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:4,marginTop:6}}>
          {all.map(sub => (
            <label key={sub} style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:C.navy}}>
              <input type="checkbox" checked={bulkSubs.includes(sub)} onChange={e=>setBulkSubs(b=>e.target.checked?[...b,sub]:b.filter(x=>x!==sub))}/>{sub}
            </label>
          ))}
        </div>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:10}}>
          <label style={{...lbl,marginBottom:0}}>Students ({bulkSel.length} selected)</label>
          <div style={{display:"flex",gap:6}}>
            <SmBtn onClick={()=>setBulkSel(pool.map(x=>x.id))} color={C.navyMid}>All</SmBtn>
            <SmBtn onClick={()=>setBulkSel(pool.filter(x=>!cfg[x.id]?.subjects?.length).map(x=>x.id))} color={C.navyMid}>No list yet</SmBtn>
            <SmBtn onClick={()=>setBulkSel([])} color={C.gray}>None</SmBtn>
          </div>
        </div>
        <div style={{maxHeight:200,overflowY:"auto",border:`1px solid ${C.grayLight}`,borderRadius:8,padding:6,marginTop:6}}>
          {!pool.length && <div style={{fontSize:12,color:C.gray}}>No registered Form 5 students.</div>}
          {pool.map(st => (
            <label key={st.id} style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:C.navy,padding:"3px 0"}}>
              <input type="checkbox" checked={bulkSel.includes(st.id)} onChange={e=>setBulkSel(b=>e.target.checked?[...b,st.id]:b.filter(x=>x!==st.id))}/>
              {st.name} <span style={{color:C.gray,fontSize:10}}>({cfg[st.id]?.subjects?.length ? cfg[st.id].subjects.length+" subjects" : "not set"})</span>
            </label>
          ))}
        </div>
        <div style={{display:"flex",gap:6,marginTop:10}}>
          {[["replace","Replace"],["add","Add"],["remove","Remove"]].map(([k,l])=>(
            <button key={k} onClick={()=>setBulkMode(k)} style={{flex:1,padding:"7px 4px",borderRadius:8,border:`2px solid ${bulkMode===k?C.navy:C.grayLight}`,background:bulkMode===k?C.navy:C.white,color:bulkMode===k?C.white:C.gray,fontWeight:700,fontSize:11.5,cursor:"pointer"}}>{l}</button>
          ))}
        </div>
        <div style={{marginTop:9}}><Btn onClick={applyBulk} disabled={busy==="bulk"}>{busy==="bulk"?"Applying…":`Apply to ${bulkSel.length} student(s)`}</Btn></div>
      </div>

      <div style={card}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:4}}>Subjects written by each Form 5 student</div>
        <div style={{fontSize:11,color:C.gray,marginBottom:8}}>Tap a student to choose the subjects they write and give them a candidate number. A student with no list yet is counted as writing every subject.</div>
        {!pool.length && <Empty text="No registered Form 5 students."/>}
        {pool.map(st => {
          const c = cfg[st.id];
          const n = c?.subjects?.length || 0;
          return (
            <div key={st.id} style={{borderTop:`1px solid ${C.grayBg}`,padding:"8px 0"}}>
              <div onClick={()=>openStudent(st)} style={{display:"flex",justifyContent:"space-between",alignItems:"center",cursor:"pointer",gap:8}}>
                <div>
                  <div style={{fontWeight:700,color:C.navy,fontSize:13}}>{st.name}</div>
                  <div style={{fontSize:10,color:C.gray,fontFamily:"monospace"}}>{st.id}{c?.candNo ? ` · Cand. No ${c.candNo}` : ""}</div>
                </div>
                <div style={{fontSize:11,fontWeight:700,color:n?C.green:C.gold}}>{n ? `${n} subjects` : "not set"}</div>
              </div>
              {open===st.id && (
                <div style={{marginTop:8}}>
                  <label style={lbl}>Candidate number</label>
                  <input style={inp} value={draft.candNo} onChange={e=>setDraft(d=>({...d,candNo:e.target.value}))} placeholder="e.g. 12345"/>
                  <label style={{...lbl,marginTop:8}}>Subjects written ({draft.subjects.length})</label>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:4}}>
                    {all.map(sub => (
                      <label key={sub} style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:C.navy}}>
                        <input type="checkbox" checked={draft.subjects.includes(sub)} onChange={e=>setDraft(d=>({...d,subjects:e.target.checked?[...d.subjects,sub]:d.subjects.filter(x=>x!==sub)}))}/>{sub}
                      </label>
                    ))}
                  </div>
                  <div style={{marginTop:9,display:"flex",gap:8}}>
                    <Btn onClick={()=>saveStudent(st)} disabled={busy===st.id}>{busy===st.id?"Saving…":"Save"}</Btn>
                    <SmBtn onClick={()=>setOpen("")} color={C.gray}>Cancel</SmBtn>
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

const GRADE_PTS = { A:5, B:4, C:3, D:2, E:1, U:0 };

// One row per Form 5 student for the chosen assessment
function f5SummaryRows(ctx, year, assess) {
  const { students, settings, marksMap, teachers } = ctx;
  const env = { marksMap, teachers, settings };
  const isG = assess===GCE_SEQ;
  const pool = f5PoolFor(students, marksMap, year);
  const rows = pool.map(st => {
    const rep = computeReport(st, env, { year, mode:"term", term:"First Term", f5Assess:assess });
    const written = f5SubjectsOf(settings, st).length;
    const entered = rep.modeRows.length;
    const subs = {};
    let passed = 0, pts = 0; const failed = []; const dist = {};
    rep.modeRows.forEach(r => {
      if (isG) {
        const g = gceFromPoints(r.termAvg); subs[r.sub] = g; dist[g] = (dist[g]||0)+1;
        if (gcePassed(r.termAvg)) passed++; else failed.push(r.sub);
        pts += Number(r.termAvg)||0;
      } else {
        const g = f5Grade(settings, r.termAvg); subs[r.sub] = Number(r.termAvg).toFixed(2);
        if (g.pass) passed++; else failed.push(r.sub);
      }
    });
    const avg = isG ? null : rep.overallAvg;
    return { st, written, entered, subs, passed, failed, pts, avg,
             grade: avg!==null && !isG ? f5Grade(settings, avg).grade : "—",
             distStr: GCE_GRADES.filter(g=>dist[g]).map(g=>g+dist[g]).join(" ") };
  });
  const ranked = rows.filter(r => r.entered>0).sort((a,b) => isG ? (b.passed-a.passed)||(b.pts-a.pts) : (b.avg??-1)-(a.avg??-1));
  let last = null, pos = 0;
  ranked.forEach((r,i) => { const key = isG ? r.passed+"|"+r.pts : r.avg; if (key!==last) { pos = i+1; last = key; } r.rank = pos; });
  return { rows, ranked };
}

function f5SummaryDocs(ranked, meta) {
  const per = 26, pages = [];
  for (let i=0; i<Math.max(ranked.length,1); i+=per) pages.push(ranked.slice(i, i+per));
  return pages.map((chunk, pi) => `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:14mm 11mm;color:#111;font-size:10pt}
    h1{font-size:14pt;margin:0 0 1mm;text-align:center} h2{font-size:10.5pt;margin:0 0 4mm;text-align:center;font-weight:600;color:#333}
    table{width:100%;border-collapse:collapse;font-size:8.5pt} th,td{border:1px solid #999;padding:1.4mm 1.6mm;text-align:center}
    th{background:#E8EEF7} td.l{text-align:left} .bad{color:#B91C1C} .ok{color:#15803D;font-weight:700}
    .foot{margin-top:4mm;font-size:8pt;color:#555;text-align:center}
  </style></head><body>
    <h1>Saker Baptist Comprehensive College Bawe</h1>
    <h2>Form 5 · ${escHtml(meta.assess)} · ${escHtml(meta.year)} — Class Summary</h2>
    <table><thead><tr><th>Rank</th><th>Name</th><th>Cand. No</th><th>Subj. written</th><th>${meta.isG?"Grades":"Average /20"}</th><th>${meta.isG?"Passed":"Grade"}</th><th>${meta.isG?"Failed subjects":"Passed (A–C)"}</th>${meta.isG?"":"<th>Failed subjects</th>"}</tr></thead><tbody>
    ${chunk.map(r => `<tr><td>${r.rank}</td><td class="l">${escHtml(r.st.name)}</td><td>${escHtml(f5CandNo(meta.settings, r.st))}</td><td>${r.entered}/${r.written}</td>
      ${meta.isG
        ? `<td>${escHtml(r.distStr)}</td><td class="ok">${r.passed}/${r.entered}</td><td class="l bad">${escHtml(r.failed.join(", "))}</td>`
        : `<td><b>${r.avg!==null?r.avg.toFixed(2):"—"}</b></td><td><b>${escHtml(r.grade)}</b></td><td class="ok">${r.passed}/${r.entered}</td><td class="l bad">${escHtml(r.failed.join(", "))}</td>`}</tr>`).join("")}
    </tbody></table>
    <div class="foot">${escHtml(meta.note)} · Page ${pi+1} of ${pages.length} · Printed ${new Date().toLocaleDateString("en-GB")}</div>
  </body></html>`);
}

// One statement-of-results page per student (GCE grades)
function f5SlipDocs(ranked, year, settings) {
  return ranked.map(r => {
    const rows = Object.entries(r.subs).map(([sub,g]) => { const ok = ["A","B","C"].includes(g); return `<tr><td class="l">${escHtml(sub)}</td><td><b>${escHtml(g)}</b></td><td class="${ok?"ok":"bad"}">${ok?"Pass":"Fail"}</td></tr>`; }).join("");
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
      body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:22mm 18mm;color:#111;font-size:11pt}
      h1{font-size:16pt;margin:0;text-align:center} h2{font-size:12pt;margin:2mm 0 8mm;text-align:center;font-weight:600}
      table{width:100%;border-collapse:collapse;margin-top:6mm} th,td{border:1px solid #888;padding:2.2mm 3mm;text-align:center} th{background:#E8EEF7}
      td.l{text-align:left} .ok{color:#15803D;font-weight:700} .bad{color:#B91C1C;font-weight:700}
      .meta td{border:none;text-align:left;padding:1.2mm 0} .foot{margin-top:14mm;font-size:9pt;color:#555;text-align:center}
    </style></head><body>
      <h1>Saker Baptist Comprehensive College Bawe</h1>
      <h2>Form 5 · GCE Statement of Results · ${escHtml(year)}</h2>
      <table class="meta"><tr><td><b>Name:</b> ${escHtml(r.st.name)}</td><td><b>Matricule:</b> ${escHtml(r.st.id)}</td></tr>
      <tr><td><b>Candidate No:</b> ${escHtml(f5CandNo(settings, r.st))}</td><td><b>Subjects passed (A–C):</b> ${r.passed} / ${r.entered}</td></tr></table>
      <table><thead><tr><th>Subject</th><th>Grade</th><th>Result</th></tr></thead><tbody>${rows}</tbody></table>
      <div class="foot">Printed ${new Date().toLocaleDateString("en-GB")} · A, B, C = Pass · D, E, U = Fail</div>
    </body></html>`;
  });
}

function F5Summary({ ctx }) {
  const { settings } = ctx;
  const [year, setYear]     = useState(ACTIVE_YEAR);
  const [assess, setAssess] = useState("Pre Mock");
  const [busy, setBusy]     = useState(false);
  const isG = assess===GCE_SEQ;
  const { rows, ranked } = useMemo(() => f5SummaryRows(ctx, year, assess), [ctx.students, ctx.marksMap, ctx.settings, year, assess]);
  const avgs = ranked.filter(r=>r.avg!==null).map(r=>r.avg);
  const classAvg = avgs.length ? (avgs.reduce((a,b)=>a+b,0)/avgs.length).toFixed(2) : "—";
  const k = f5Scale(settings);
  const note = isG ? "GCE grades: A, B, C = Pass · D, E, U = Fail" : `Grading: A ≥${k.A} · B ≥${k.B} · C ≥${k.C} · D ≥${k.D} · E ≥${k.E} · U below · A–C = Pass`;

  function excel() {
    const allSubs = SUBJECTS_BY_FORM["Form 5"];
    const out = ranked.map(r => {
      const o = { Rank:r.rank, Name:r.st.name, Matricule:r.st.id, "Candidate No":f5CandNo(settings, r.st), "Subjects written":r.written, "Subjects entered":r.entered };
      allSubs.forEach(sb => { o[sb] = r.subs[sb] ?? (f5Writes(settings, r.st, sb) ? "" : "—"); });
      if (isG) { o["Passed (A–C)"] = r.passed; o["Failed subjects"] = r.failed.join(", "); }
      else { o["Average /20"] = r.avg!==null ? Number(r.avg.toFixed(2)) : ""; o["Grade"] = r.grade; o["Passed (A–C)"] = r.passed; o["Failed subjects"] = r.failed.join(", "); }
      return o;
    });
    exportToExcel(out, `SBCC-Form5-${assess.replace(/\s+/g,"")}-${year.replace("/","-")}.xlsx`, "Form 5 "+assess);
  }
  async function pdf() {
    if (!ranked.length) { alert("No marks entered yet for this assessment."); return; }
    setBusy(true);
    try {
      const blob = await htmlDocsToPdfBlob(f5SummaryDocs(ranked, { assess, year, isG, settings, note }), null, 1.5, 0.85);
      downloadBlob(blob, `Form5-${assess.replace(/\s+/g,"")}-Summary.pdf`);
    } catch(e) { alert("Could not create the PDF: "+e.message); }
    setBusy(false);
  }
  async function slips() {
    if (!ranked.length) { alert("No GCE grades entered yet."); return; }
    setBusy(true);
    try { downloadBlob(await htmlDocsToPdfBlob(f5SlipDocs(ranked.slice().sort((a,b)=>a.st.name.localeCompare(b.st.name)), year, settings), null, 1.5, 0.85), `Form5-GCE-Results-Slips-${year.replace("/","-")}.pdf`); }
    catch(e) { alert("Could not create the PDF: "+e.message); }
    setBusy(false);
  }
  const card = { background:C.white, borderRadius:10, padding:12, marginBottom:11, boxShadow:"0 1px 3px rgba(0,0,0,0.06)" };
  const passTotal = ranked.reduce((a,r)=>a+r.passed,0), enteredTotal = ranked.reduce((a,r)=>a+r.entered,0);
  return (
    <div>
      <div style={card}>
        <div style={{display:"flex",gap:6,marginBottom:8}}>
          {F5_ASSESSMENTS.map(t=>(
            <button key={t} onClick={()=>setAssess(t)} style={{flex:1,padding:"8px 4px",borderRadius:8,border:`2px solid ${assess===t?C.navy:C.grayLight}`,background:assess===t?C.navy:C.white,color:assess===t?C.white:C.gray,fontWeight:700,fontSize:11,cursor:"pointer"}}>{t}</button>
          ))}
        </div>
        <select style={inp} value={year} onChange={e=>setYear(e.target.value)}>{ACAD_YEARS.map(y=><option key={y}>{y}</option>)}</select>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginTop:9,textAlign:"center"}}>
          <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>Students</div><div style={{fontSize:17,fontWeight:900,color:C.navy}}>{ranked.length}/{rows.length}</div></div>
          <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>{isG?"Subject passes":"Class average"}</div><div style={{fontSize:17,fontWeight:900,color:C.navy}}>{isG ? `${passTotal}/${enteredTotal}` : classAvg}</div></div>
          <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>Pass rate</div><div style={{fontSize:17,fontWeight:900,color:C.navy}}>{enteredTotal ? Math.round(passTotal/enteredTotal*100)+"%" : "—"}</div></div>
        </div>
        <div style={{display:"flex",gap:8,marginTop:10,flexWrap:"wrap"}}><Btn onClick={pdf} disabled={busy}>{busy?"Making PDF…":"📄 PDF"}</Btn><Btn onClick={excel} outline>📊 Excel (all subjects)</Btn>{isG && <Btn onClick={slips} outline disabled={busy}>🧾 Results slips</Btn>}</div>
      </div>
      {!ranked.length && <Empty text="No marks entered yet for this assessment and year."/>}
      {ranked.map(r => {
        const good = r.passed===r.entered;
        return (
          <div key={r.st.id} style={{background:C.white,borderRadius:10,padding:"9px 12px",marginBottom:7,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:30,height:30,borderRadius:15,background:C.navy,color:C.white,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:12,flexShrink:0}}>{r.rank}</div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontWeight:700,color:C.navy,fontSize:13}}>{r.st.name}</div>
                <div style={{fontSize:10,color:C.gray}}>{f5CandNo(settings, r.st) ? "Cand. "+f5CandNo(settings, r.st)+" · " : ""}{r.entered}/{r.written} subjects entered</div>
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{fontWeight:900,fontSize:15,color:C.navy}}>{isG ? r.distStr||"—" : `${r.avg!==null?r.avg.toFixed(2):"—"} · ${r.grade}`}</div>
                <div style={{fontSize:10.5,fontWeight:700,color:good?C.green:C.gold}}>{r.passed}/{r.entered} passed</div>
              </div>
            </div>
            {r.failed.length>0 && <div style={{fontSize:10.5,color:C.red,marginTop:5}}>Failed: {r.failed.join(", ")}</div>}
          </div>
        );
      })}
    </div>
  );
}

// GCE results: pass rates, grade spread and movement from the mocks
function F5Gce({ ctx }) {
  const { students, settings, marksMap, teachers } = ctx;
  const [year, setYear] = useState(ACTIVE_YEAR);
  const env = { marksMap, teachers, settings };
  const pool = f5PoolFor(students, marksMap, year);
  const data = useMemo(() => {
    const subj = {}; const stu = [];
    pool.forEach(st => {
      let entered = 0, passed = 0;
      f5SubjectsOf(settings, st).forEach(sub => {
        const o = subj[sub] || (subj[sub] = { sub, written:0, entered:0, passed:0, dist:{}, up:0, same:0, down:0 });
        o.written++;
        const gp = marksMap[`${st.id}-${sub}-${GCE_SEQ}-${year}`]?.score;
        if (gp===undefined || gp===null) return;
        o.entered++; entered++;
        const g = gceFromPoints(gp); o.dist[g] = (o.dist[g]||0)+1;
        if (gcePassed(gp)) { o.passed++; passed++; }
        const rs = marksMap[`${st.id}-${sub}-Regional Mock-${year}`]?.score;
        if (rs!==undefined && rs!==null) {
          const rp = GRADE_PTS[f5Grade(settings, rs).grade];
          if (Number(gp) > rp) o.up++; else if (Number(gp) < rp) o.down++; else o.same++;
        }
      });
      const pre = computeReport(st, env, { year, mode:"term", term:"First Term", f5Assess:"Pre Mock" }).overallAvg;
      const reg = computeReport(st, env, { year, mode:"term", term:"First Term", f5Assess:"Regional Mock" }).overallAvg;
      stu.push({ st, entered, passed, pre, reg });
    });
    return { subjects: Object.values(subj).filter(o=>o.entered>0).sort((a,b)=>(b.passed/b.entered)-(a.passed/a.entered)), stu: stu.filter(x=>x.entered>0).sort((a,b)=>(b.passed/b.entered)-(a.passed/a.entered)||b.passed-a.passed) };
  }, [students, marksMap, settings, year]);
  const totE = data.subjects.reduce((a,o)=>a+o.entered,0), totP = data.subjects.reduce((a,o)=>a+o.passed,0);
  const allPass = data.stu.filter(x=>x.passed===x.entered).length;
  const best = data.subjects[0], worst = data.subjects[data.subjects.length-1];
  const card = { background:C.white, borderRadius:10, padding:12, marginBottom:11, boxShadow:"0 1px 3px rgba(0,0,0,0.06)" };
  const pct = (a,b) => b ? Math.round(a/b*100)+"%" : "—";
  function xlSubjects() {
    exportToExcel(data.subjects.map(o => ({ Subject:o.sub, Written:o.written, Entered:o.entered, Passed:o.passed, "Pass rate":pct(o.passed,o.entered), A:o.dist.A||0, B:o.dist.B||0, C:o.dist.C||0, D:o.dist.D||0, E:o.dist.E||0, U:o.dist.U||0, "Up vs Regional":o.up, "Same":o.same, "Down":o.down })), `SBCC-GCE-Subjects-${year.replace("/","-")}.xlsx`, "GCE subjects");
  }
  function xlStudents() {
    exportToExcel(data.stu.map(x => ({ Name:x.st.name, Matricule:x.st.id, "Candidate No":f5CandNo(settings, x.st), "Pre Mock avg /20":x.pre!==null?Number(x.pre.toFixed(2)):"", "Regional Mock avg /20":x.reg!==null?Number(x.reg.toFixed(2)):"", "GCE subjects entered":x.entered, "GCE passed (A–C)":x.passed })), `SBCC-GCE-Students-${year.replace("/","-")}.xlsx`, "GCE students");
  }
  return (
    <div>
      <div style={card}>
        <select style={inp} value={year} onChange={e=>setYear(e.target.value)}>{ACAD_YEARS.map(y=><option key={y}>{y}</option>)}</select>
        {!data.subjects.length ? <div style={{marginTop:8}}><Empty text="No GCE grades entered for this year yet. Enter them in Marks → Form 5 → GCE Result."/></div> : <>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginTop:9,textAlign:"center"}}>
            <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>Overall pass rate (A–C)</div><div style={{fontSize:19,fontWeight:900,color:C.navy}}>{pct(totP,totE)}</div><div style={{fontSize:10,color:C.gray}}>{totP}/{totE} subject results</div></div>
            <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>Passed every subject</div><div style={{fontSize:19,fontWeight:900,color:C.navy}}>{allPass}/{data.stu.length}</div><div style={{fontSize:10,color:C.gray}}>students</div></div>
            <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>Strongest subject</div><div style={{fontSize:12,fontWeight:800,color:C.green}}>{best?.sub}</div><div style={{fontSize:10,color:C.gray}}>{best ? pct(best.passed,best.entered) : ""}</div></div>
            <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>Weakest subject</div><div style={{fontSize:12,fontWeight:800,color:C.red}}>{worst?.sub}</div><div style={{fontSize:10,color:C.gray}}>{worst ? pct(worst.passed,worst.entered) : ""}</div></div>
          </div>
          <div style={{display:"flex",gap:8,marginTop:10}}><SmBtn onClick={xlSubjects} color={C.navyMid}>📊 Subjects Excel</SmBtn><SmBtn onClick={xlStudents} color={C.navyMid}>📊 Students Excel</SmBtn></div>
        </>}
      </div>

      {data.subjects.length>0 && <div style={card}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:6}}>By subject (best pass rate first)</div>
        {data.subjects.map(o => {
          const mv = o.up+o.same+o.down;
          return (
            <div key={o.sub} style={{borderTop:`1px solid ${C.grayBg}`,padding:"7px 0"}}>
              <div style={{display:"flex",justifyContent:"space-between",gap:8}}>
                <div style={{fontWeight:700,color:C.navy,fontSize:12.5}}>{o.sub}</div>
                <div style={{fontWeight:800,fontSize:12.5,color:o.passed/o.entered>=0.5?C.green:C.red}}>{pct(o.passed,o.entered)} <span style={{fontWeight:600,color:C.gray,fontSize:10.5}}>({o.passed}/{o.entered})</span></div>
              </div>
              <div style={{fontSize:10.5,color:C.gray,marginTop:2}}>{GCE_GRADES.filter(g=>o.dist[g]).map(g=>g+": "+o.dist[g]).join("  ·  ")}</div>
              {mv>0 && <div style={{fontSize:10.5,color:C.gray,marginTop:1}}>vs Regional Mock: <span style={{color:C.green,fontWeight:700}}>▲ {o.up}</span> · ▬ {o.same} · <span style={{color:C.red,fontWeight:700}}>▼ {o.down}</span></div>}
            </div>
          );
        })}
      </div>}

      {data.stu.length>0 && <div style={card}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:6}}>By student</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 46px 46px 56px",gap:4,fontSize:9.5,color:C.gray,fontWeight:700,paddingBottom:4}}><div>Name</div><div style={{textAlign:"center"}}>Pre</div><div style={{textAlign:"center"}}>Reg.</div><div style={{textAlign:"center"}}>GCE</div></div>
        {data.stu.map(x => (
          <div key={x.st.id} style={{display:"grid",gridTemplateColumns:"1fr 46px 46px 56px",gap:4,alignItems:"center",borderTop:`1px solid ${C.grayBg}`,padding:"6px 0",fontSize:12}}>
            <div style={{fontWeight:700,color:C.navy}}>{x.st.name}</div>
            <div style={{textAlign:"center",color:C.navyMid}}>{x.pre!==null?x.pre.toFixed(1):"—"}</div>
            <div style={{textAlign:"center",color:C.navyMid}}>{x.reg!==null?x.reg.toFixed(1):"—"}</div>
            <div style={{textAlign:"center",fontWeight:800,color:x.passed===x.entered?C.green:C.gold}}>{x.passed}/{x.entered}</div>
          </div>
        ))}
        <div style={{fontSize:10,color:C.gray,marginTop:6}}>Pre and Reg. are mock averages out of 20. GCE is subjects passed (A–C) out of subjects entered.</div>
      </div>}
    </div>
  );
}

// Exam cards: 6 per A4 page (2 × 3). Each card carries the photo, candidate number and the subjects the candidate writes.
function f5CardDocs(list, year, settings) {
  const card = st => {
    const subs = f5SubjectsOf(settings, st);
    const half = Math.ceil(subs.length/2);
    const col  = (arr, off) => arr.map((x,i) => `<div class="s">${off+i+1}. ${escHtml(x)}</div>`).join("");
    const photo = st.photo_url
      ? `<img src="${st.photo_url}" alt="" style="width:22mm;height:28mm;object-fit:cover;border:0.4mm solid #0D2340;display:block">`
      : `<div style="width:22mm;height:28mm;border:0.4mm dashed #999;box-sizing:border-box;display:flex;align-items:center;justify-content:center;font-size:7pt;color:#888;text-align:center">Paste<br>photo</div>`;
    const cn = f5CandNo(settings, st);
    return `<div class="card">
      <div class="hd"><img src="${SCHOOL_LOGO}" alt="" style="width:11mm;height:11mm;object-fit:contain"><div><div class="sc">Saker Baptist Comprehensive College Bawe</div><div class="tt">GCE Examination Card · ${escHtml(year)}</div></div></div>
      <div class="id">${photo}<div class="info"><div class="nm">${escHtml(st.name)}</div><div class="ln"><b>Matricule:</b> ${escHtml(st.id)}</div><div class="ln"><b>Class:</b> Form 5</div><div class="cn"><span>Candidate No</span><b>${cn ? escHtml(cn) : "&nbsp;"}</b></div></div></div>
      <div class="sh">Subjects written (${subs.length})</div>
      <div class="subs"><div>${col(subs.slice(0,half),0)}</div><div>${col(subs.slice(half),half)}</div></div>
      <div class="sg"><span>Candidate's signature</span><span>Principal</span></div>
    </div>`;
  };
  const pages = [];
  for (let i=0; i<list.length; i+=6) pages.push(list.slice(i,i+6));
  return pages.map(group => `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    body{margin:0;padding:8mm 10mm;font-family:Arial,Helvetica,sans-serif;color:#111;background:#fff}
    .grid{display:grid;grid-template-columns:93mm 93mm;grid-auto-rows:90mm;gap:3mm 4mm}
    .card{border:0.4mm solid #0D2340;border-radius:2mm;padding:2.5mm 3mm;box-sizing:border-box;overflow:hidden;display:flex;flex-direction:column}
    .hd{display:flex;gap:2mm;align-items:center;border-bottom:0.3mm solid #C9962A;padding-bottom:1.5mm}
    .sc{font-weight:700;font-size:8pt;color:#0D2340} .tt{font-size:7.5pt;color:#555}
    .id{display:flex;gap:3mm;margin-top:2mm} .info{flex:1;min-width:0}
    .nm{font-weight:700;font-size:10pt;margin-bottom:1mm} .ln{font-size:8pt;margin-bottom:0.6mm}
    .cn{margin-top:1.5mm;border:0.3mm solid #0D2340;padding:1mm 2mm;display:flex;justify-content:space-between;align-items:center;font-size:8pt} .cn b{font-size:11pt;letter-spacing:0.3mm}
    .sh{font-weight:700;font-size:7.5pt;margin:2mm 0 0.8mm;color:#0D2340;border-bottom:0.2mm solid #ccc}
    .subs{display:grid;grid-template-columns:1fr 1fr;gap:0 2mm;flex:1} .s{font-size:7.5pt;line-height:3.6mm;white-space:nowrap;overflow:hidden}
    .sg{display:flex;justify-content:space-between;font-size:6.5pt;color:#555;padding-top:1mm} .sg span{border-top:0.2mm solid #555;width:40%;padding-top:0.5mm;text-align:center}
  </style></head><body><div class="grid">${group.map(card).join("")}</div></body></html>`);
}

function F5Cards({ ctx }) {
  const { students, settings } = ctx;
  const [year, setYear] = useState(ACTIVE_YEAR);
  const [skip, setSkip] = useState({});   // student id → true when left out of the print
  const [busy, setBusy] = useState("");
  const pool = students.filter(x => x.active && x.form==="Form 5" && x.reg_status==="registered").sort((a,b)=>a.name.localeCompare(b.name));
  const cfg  = f5Students(settings);
  const chosen = pool.filter(x => !skip[x.id]);
  const issues = st => [!st.photo_url && "no photo", !f5CandNo(settings, st) && "no candidate number", !cfg[st.id]?.subjects?.length && "subjects not set"].filter(Boolean);
  const noPhoto = chosen.filter(x=>!x.photo_url).length, noCand = chosen.filter(x=>!f5CandNo(settings,x)).length, noSubs = chosen.filter(x=>!cfg[x.id]?.subjects?.length).length;
  async function print() {
    if (!chosen.length) { alert("Select at least one candidate."); return; }
    if ((noCand || noSubs) && !window.confirm(`${noCand} selected candidate(s) have no candidate number and ${noSubs} have no subject list yet (they print with every Form 5 subject). Print anyway?`)) return;
    const docs = f5CardDocs(chosen, year, settings);
    setBusy(`0/${docs.length}`);
    try { downloadBlob(await htmlDocsToPdfBlob(docs, (d,t)=>setBusy(`${d}/${t}`), 1.5, 0.85), `Form5-Exam-Cards-${year.replace("/","-")}.pdf`); }
    catch(e) { alert("Could not create the PDF: "+e.message); }
    setBusy("");
  }
  const card = { background:C.white, borderRadius:10, padding:12, marginBottom:11, boxShadow:"0 1px 3px rgba(0,0,0,0.06)" };
  return (
    <div>
      <div style={card}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:4}}>Candidate exam cards</div>
        <div style={{fontSize:11,color:C.gray,marginBottom:8}}>One card per candidate with photo, matricule, candidate number and the subjects they write. Six cards per A4 page. Photos come from Students; numbers and subjects from Setup.</div>
        <select style={inp} value={year} onChange={e=>setYear(e.target.value)}>{ACAD_YEARS.map(y=><option key={y}>{y}</option>)}</select>
        {(noPhoto>0 || noCand>0 || noSubs>0) && <div style={{fontSize:11,color:C.gold,marginTop:8}}>Among the selected: {noPhoto} without photo · {noCand} without candidate number · {noSubs} without a subject list</div>}
        <div style={{marginTop:9}}><Btn onClick={print} disabled={!!busy}>{busy?`Making PDF… ${busy}`:`🪪 Exam cards PDF (${chosen.length})`}</Btn></div>
      </div>
      <div style={card}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6}}>
          <div style={{fontWeight:800,color:C.navy,fontSize:13}}>Candidates ({chosen.length}/{pool.length})</div>
          <div style={{display:"flex",gap:6}}><SmBtn onClick={()=>setSkip({})} color={C.navyMid}>All</SmBtn><SmBtn onClick={()=>setSkip(Object.fromEntries(pool.map(x=>[x.id,true])))} color={C.gray}>None</SmBtn></div>
        </div>
        {!pool.length && <Empty text="No registered Form 5 students."/>}
        {pool.map(st => {
          const is = issues(st);
          return (
            <label key={st.id} style={{display:"flex",alignItems:"center",gap:8,borderTop:`1px solid ${C.grayBg}`,padding:"7px 0"}}>
              <input type="checkbox" checked={!skip[st.id]} onChange={e=>setSkip(k=>({...k,[st.id]:!e.target.checked}))}/>
              <PhotoBox photo={st.photo_url} size={[26,32]}/>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontWeight:700,color:C.navy,fontSize:12.5}}>{st.name}</div>
                <div style={{fontSize:10,color:C.gray}}>{f5CandNo(settings, st) ? "Cand. No "+f5CandNo(settings, st) : "—"}{cfg[st.id]?.subjects?.length ? ` · ${cfg[st.id].subjects.length} subjects` : ""}</div>
                {is.length>0 && <div style={{fontSize:10,color:C.red}}>{is.join(" · ")}</div>}
              </div>
            </label>
          );
        })}
      </div>
    </div>
  );
}

function F5Tools({ ctx }) {
  const [sub, setSub] = useState("summary");
  return (
    <div>
      <div style={{display:"flex",gap:6,marginBottom:10}}>
        {[["summary","Class summary"],["gce","GCE analysis"],["cards","Exam cards"],["setup","Setup"]].map(([k,l])=>(
          <button key={k} onClick={()=>setSub(k)} style={{flex:1,padding:"8px 4px",borderRadius:8,border:`2px solid ${sub===k?C.navy:C.grayLight}`,background:sub===k?C.navy:C.white,color:sub===k?C.white:C.gray,fontWeight:700,fontSize:11.5,cursor:"pointer"}}>{l}</button>
        ))}
      </div>
      {sub==="summary" && <F5Summary ctx={ctx}/>}
      {sub==="gce"     && <F5Gce ctx={ctx}/>}
      {sub==="cards"   && <F5Cards ctx={ctx}/>}
      {sub==="setup"   && <F5SetupPanel ctx={ctx}/>}
    </div>
  );
}

// ─── Honour roll and certificates (Forms 1–4) ──────────────────────────────────
const ordinal = n => n+(["th","st","nd","rd"][(n%100>10&&n%100<14)?0:Math.min(n%10,4)%4]||"th");
// Best overall student and best score per subject in each class for a period. Ties share the award.
function honourData(ctx, forms, year, period) {
  return forms.map(form => {
    const { ranked, subs } = classSummaryRows(ctx, form, year, period);
    const top = ranked.length ? ranked.filter(r => r.rank===1) : [];
    const subjects = [];
    subs.forEach(sub => {
      let best = -1, winners = [];
      ranked.forEach(r => { const v = r.vals[sub]; if (v===undefined) return; if (v>best+1e-9) { best=v; winners=[r]; } else if (Math.abs(v-best)<=1e-9) winners.push(r); });
      if (winners.length) subjects.push({ sub, score:best, winners:winners.map(w=>w.st) });
    });
    return { form, top, subjects };
  });
}
function certificateDocs(items, year, period) {
  return items.map(it => `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    body{margin:0;padding:12mm;font-family:Georgia,'Times New Roman',serif;color:#111;background:#fff}
    .b{border:2mm double #0D2340;height:271mm;box-sizing:border-box;padding:14mm 16mm;text-align:center;position:relative}
    .in{border:0.4mm solid #C9962A;position:absolute;inset:3mm}
    .c{position:relative}
    h1{font-size:19pt;margin:4mm 0 0;color:#0D2340;letter-spacing:0.3mm} .sub{font-size:11pt;color:#555;margin-top:1mm}
    h2{font-size:30pt;margin:16mm 0 4mm;color:#C9962A;font-weight:normal} .t{font-size:13pt;margin:5mm 0} .nm{font-size:25pt;font-weight:bold;color:#0D2340;margin:6mm 0;border-bottom:0.4mm solid #999;display:inline-block;padding:0 10mm 1mm}
    .what{font-size:15pt;line-height:1.6;margin:6mm 10mm} .sg{display:flex;justify-content:space-around;margin-top:38mm;font-size:11pt} .sg span{border-top:0.3mm solid #111;width:60mm;padding-top:1.5mm}
  </style></head><body><div class="b"><div class="in"></div><div class="c">
    <img src="${SCHOOL_LOGO}" alt="" style="width:30mm;height:30mm;object-fit:contain">
    <h1>SAKER BAPTIST COMPREHENSIVE COLLEGE BAWE</h1>
    <h2>Certificate of Excellence</h2>
    <div class="t">This certificate is proudly awarded to</div>
    <div class="nm">${escHtml(it.st.name)}</div>
    <div class="t">${escHtml(it.st.form)} · Matricule ${escHtml(it.st.id)}</div>
    <div class="what">${it.kind==="class"
      ? `for ranking <b>${ordinal(1)}</b> in ${escHtml(it.st.form)} with an average of <b>${it.score.toFixed(2)}/20</b> in ${escHtml(period)}, ${escHtml(year)}.`
      : `for the best performance in <b>${escHtml(it.sub)}</b> in ${escHtml(it.st.form)} with <b>${it.score.toFixed(2)}/20</b> in ${escHtml(period)}, ${escHtml(year)}.`}</div>
    <div class="sg"><span>Class Teacher</span><span>Principal</span></div>
    <div class="t" style="margin-top:12mm;font-size:10pt;color:#555">Date: ${new Date().toLocaleDateString("en-GB")}</div>
  </div></div></body></html>`);
}

function HonourRoll({ ctx }) {
  const [year, setYear]     = useState(ACTIVE_YEAR);
  const [period, setPeriod] = useState("Annual");
  const [form, setForm]     = useState("");
  const [busy, setBusy]     = useState("");
  const forms = FORMS.filter(f => !isF5(f) && (!form || f===form));
  const data = useMemo(() => honourData(ctx, forms, year, period), [ctx.students, ctx.marksMap, year, period, form]);
  async function make(kind) {
    const items = [];
    data.forEach(d => {
      if (kind==="class") d.top.forEach(r => items.push({ kind:"class", st:r.st, score:r.avg }));
      else d.subjects.forEach(x => x.winners.forEach(st => items.push({ kind:"subject", st, sub:x.sub, score:x.score })));
    });
    if (!items.length) { alert("No marks entered for this selection."); return; }
    if (items.length>40 && !window.confirm(`${items.length} certificates will be made. This can be slow on a phone. Choose one class to make fewer. Continue?`)) return;
    setBusy(kind);
    try { downloadBlob(await htmlDocsToPdfBlob(certificateDocs(items, year, period), null, 1.2, 0.8), `Certificates-${kind==="class"?"Best-student":"Best-in-subject"}-${(form||"All").replace(/\s+/g,"")}-${period.replace(/\s+/g,"")}.pdf`); }
    catch(e) { alert("Could not create the PDF: "+e.message); }
    setBusy("");
  }
  const card = { background:C.white, borderRadius:10, padding:12, marginBottom:11, boxShadow:"0 1px 3px rgba(0,0,0,0.06)" };
  return (
    <div>
      <div style={card}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:4}}>Honour roll and certificates</div>
        <div style={{fontSize:11,color:C.gray,marginBottom:8}}>Best student per class and best score per subject, for Forms 1 to 4. A tie shares the award.</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
          <div><label style={lbl}>Academic year</label><select style={inp} value={year} onChange={e=>setYear(e.target.value)}>{ACAD_YEARS.map(y=><option key={y}>{y}</option>)}</select></div>
          <div><label style={lbl}>Period</label><select style={inp} value={period} onChange={e=>setPeriod(e.target.value)}>{PERIODS.map(x=><option key={x}>{x}</option>)}</select></div>
        </div>
        <label style={{...lbl,marginTop:8}}>Class</label>
        <select style={inp} value={form} onChange={e=>setForm(e.target.value)}><option value="">All classes (Forms 1–4)</option>{FORMS.filter(f=>!isF5(f)).map(f=><option key={f}>{f}</option>)}</select>
        <div style={{display:"flex",gap:8,marginTop:10,flexWrap:"wrap"}}>
          <Btn onClick={()=>make("class")} disabled={!!busy}>{busy==="class"?"Making PDF…":"🏅 Best student certificates"}</Btn>
          <Btn onClick={()=>make("subject")} outline disabled={!!busy}>{busy==="subject"?"Making PDF…":"🎖 Best in subject certificates"}</Btn>
        </div>
      </div>
      {data.map(d => (
        <div key={d.form} style={card}>
          <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:6}}>{d.form}</div>
          {!d.top.length ? <div style={{fontSize:12,color:C.gray}}>No marks entered for this period.</div> : <>
            <div style={{fontSize:12.5,color:C.navy}}>🏅 <b>{d.top.map(r=>r.st.name).join(" & ")}</b> — {d.top[0].avg.toFixed(2)}/20</div>
            <div style={{marginTop:7}}>
              {d.subjects.map(x => (
                <div key={x.sub} style={{display:"flex",justifyContent:"space-between",gap:8,borderTop:`1px solid ${C.grayBg}`,padding:"4px 0",fontSize:11.5}}>
                  <span style={{color:C.gray}}>{x.sub}</span>
                  <span style={{color:C.navy,fontWeight:600,textAlign:"right"}}>{x.winners.map(w=>w.name).join(", ")} <span style={{color:C.gold}}>{x.score.toFixed(1)}</span></span>
                </div>
              ))}
            </div>
          </>}
        </div>
      ))}
    </div>
  );
}

// ─── Class summary (Forms 1–4) and Subject results ─────────────────────────────
const PERIODS = [...SEQ_LABELS, ...TERMS, "Annual"];
const GRADE_ORDER = ["A","B+","B","C+","C","D","F"];

// A student's score in one subject for a period: a sequence, a term (avg of its 2 sequences) or the year (avg of the 3 terms)
function subjectPeriodScore(marksMap, st, sub, year, period) {
  const g = sq => { const v = marksMap[`${st.id}-${sub}-${sq}-${year}`]?.score; return v===null||v===undefined ? null : Number(v); };
  const avg = a => { const v = a.filter(x=>x!==null); return v.length ? v.reduce((x,y)=>x+y,0)/v.length : null; };
  if (SEQ_LABELS.includes(period)) return g(period);
  if (TERMS.includes(period)) return avg(TERM_SEQS[period].map(g));
  return avg(TERMS.map(t => avg(TERM_SEQS[t].map(g))));
}

function classSummaryRows(ctx, form, year, period) {
  const { students, marksMap } = ctx;
  const subs = SUBJECTS_BY_FORM[form] || [];
  const pool = students.filter(x => x.active && x.reg_status==="registered" && x.form===form).sort((a,b)=>a.name.localeCompare(b.name));
  const rows = pool.map(st => {
    let tw=0, tc=0, passed=0; const failed=[], vals={};
    subs.forEach(sub => {
      const v = subjectPeriodScore(marksMap, st, sub, year, period);
      if (v===null) return;
      const c = marksMap[`${st.id}-${sub}-SEQ 1-${year}`]?.coeff || getCoeff(sub);
      vals[sub] = v; tw += v*c; tc += c;
      if (v>=PASS_MARK) passed++; else failed.push(sub);
    });
    const avg = tc ? tw/tc : null;
    return { st, vals, entered:Object.keys(vals).length, avg, grade: avg!==null ? scoreToGrade(avg).grade : "—", passed, failed,
             decision: period==="Annual" && avg!==null ? (avg>=PASS_MARK ? "Promoted" : "Repeat") : "" };
  });
  const ranked = rows.filter(r=>r.avg!==null).sort((a,b)=>b.avg-a.avg);
  let last=null, pos=0;
  ranked.forEach((r,i) => { if (r.avg!==last) { pos=i+1; last=r.avg; } r.rank=pos; });
  return { rows, ranked, subs };
}

function classSummaryDocs(ranked, meta) {
  const per = 26, pages = [];
  for (let i=0; i<Math.max(ranked.length,1); i+=per) pages.push(ranked.slice(i, i+per));
  const dec = meta.period==="Annual";
  return pages.map((chunk, pi) => `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:14mm 11mm;color:#111;font-size:10pt}
    h1{font-size:14pt;margin:0 0 1mm;text-align:center} h2{font-size:10.5pt;margin:0 0 4mm;text-align:center;font-weight:600;color:#333}
    table{width:100%;border-collapse:collapse;font-size:8.5pt} th,td{border:1px solid #999;padding:1.4mm 1.6mm;text-align:center}
    th{background:#E8EEF7} td.l{text-align:left} .bad{color:#B91C1C} .ok{color:#15803D;font-weight:700}
    .foot{margin-top:4mm;font-size:8pt;color:#555;text-align:center}
  </style></head><body>
    <h1>Saker Baptist Comprehensive College Bawe</h1>
    <h2>${escHtml(meta.form)} · ${escHtml(meta.period)} · ${escHtml(meta.year)} — Class Summary</h2>
    <table><thead><tr><th>Rank</th><th>Name</th><th>Subj.</th><th>Average /20</th><th>Grade</th><th>Passed</th><th>Failed subjects</th>${dec?"<th>Decision</th>":""}</tr></thead><tbody>
    ${chunk.map(r => `<tr><td>${r.rank}</td><td class="l">${escHtml(r.st.name)}</td><td>${r.entered}</td><td><b>${r.avg.toFixed(2)}</b></td><td><b>${escHtml(r.grade)}</b></td><td class="ok">${r.passed}/${r.entered}</td><td class="l bad">${escHtml(r.failed.join(", "))}</td>${dec?`<td class="${r.decision==="Promoted"?"ok":"bad"}"><b>${r.decision}</b></td>`:""}</tr>`).join("")}
    </tbody></table>
    <div class="foot">Pass mark ${PASS_MARK}/20 · Page ${pi+1} of ${pages.length} · Printed ${new Date().toLocaleDateString("en-GB")}</div>
  </body></html>`);
}

function ClassSummaryPanel({ ctx }) {
  const [form, setForm]     = useState("Form 1");
  const [year, setYear]     = useState(ACTIVE_YEAR);
  const [period, setPeriod] = useState("First Term");
  const [busy, setBusy]     = useState(false);
  const { rows, ranked, subs } = useMemo(() => classSummaryRows(ctx, form, year, period), [ctx.students, ctx.marksMap, form, year, period]);
  const avgs = ranked.map(r=>r.avg);
  const classAvg = avgs.length ? avgs.reduce((a,b)=>a+b,0)/avgs.length : null;
  const passing = ranked.filter(r=>r.avg>=PASS_MARK).length;
  const card = { background:C.white, borderRadius:10, padding:12, marginBottom:11, boxShadow:"0 1px 3px rgba(0,0,0,0.06)" };
  function excel() {
    const out = ranked.map(r => {
      const o = { Rank:r.rank, Name:r.st.name, Matricule:r.st.id };
      subs.forEach(sb => { o[sb] = r.vals[sb]!==undefined ? Number(r.vals[sb].toFixed(2)) : ""; });
      o["Average /20"] = Number(r.avg.toFixed(2)); o["Grade"] = r.grade; o["Passed"] = r.passed; o["Failed subjects"] = r.failed.join(", ");
      if (period==="Annual") o["Decision"] = r.decision;
      return o;
    });
    exportToExcel(out, `SBCC-${form.replace(/\s+/g,"")}-${period.replace(/\s+/g,"")}-${year.replace("/","-")}.xlsx`, `${form} ${period}`.slice(0,31));
  }
  async function pdf() {
    if (!ranked.length) { alert("No marks entered yet for this selection."); return; }
    setBusy(true);
    try { downloadBlob(await htmlDocsToPdfBlob(classSummaryDocs(ranked, { form, year, period }), null, 1.5, 0.85), `${form.replace(/\s+/g,"")}-${period.replace(/\s+/g,"")}-Summary.pdf`); }
    catch(e) { alert("Could not create the PDF: "+e.message); }
    setBusy(false);
  }
  return (
    <div>
      <div style={card}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
          <div><label style={lbl}>Class</label><select style={inp} value={form} onChange={e=>setForm(e.target.value)}>{FORMS.filter(f=>!isF5(f)).map(f=><option key={f}>{f}</option>)}</select></div>
          <div><label style={lbl}>Academic year</label><select style={inp} value={year} onChange={e=>setYear(e.target.value)}>{ACAD_YEARS.map(y=><option key={y}>{y}</option>)}</select></div>
        </div>
        <label style={{...lbl,marginTop:8}}>Period</label>
        <select style={inp} value={period} onChange={e=>setPeriod(e.target.value)}>{PERIODS.map(x=><option key={x}>{x}</option>)}</select>
        <div style={{fontSize:10.5,color:C.gray,marginTop:5}}>Form 5 has its own summary under the Form 5 tab.</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginTop:9,textAlign:"center"}}>
          <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>Assessed</div><div style={{fontSize:17,fontWeight:900,color:C.navy}}>{ranked.length}/{rows.length}</div></div>
          <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>Class average</div><div style={{fontSize:17,fontWeight:900,color:classAvg!==null?avgColor(classAvg):C.gray}}>{classAvg!==null?classAvg.toFixed(2):"—"}</div></div>
          <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>Pass rate</div><div style={{fontSize:17,fontWeight:900,color:C.navy}}>{ranked.length ? Math.round(passing/ranked.length*100)+"%" : "—"}</div></div>
        </div>
        <div style={{display:"flex",gap:8,marginTop:10}}><Btn onClick={pdf} disabled={busy}>{busy?"Making PDF…":"📄 PDF"}</Btn><Btn onClick={excel} outline>📊 Excel (all subjects)</Btn></div>
      </div>
      {!ranked.length && <Empty text="No marks entered yet for this class and period."/>}
      {ranked.map(r => (
        <div key={r.st.id} style={{background:C.white,borderRadius:10,padding:"9px 12px",marginBottom:7,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:30,height:30,borderRadius:15,background:C.navy,color:C.white,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:12,flexShrink:0}}>{r.rank}</div>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontWeight:700,color:C.navy,fontSize:13}}>{r.st.name}</div>
              <div style={{fontSize:10,color:C.gray}}>{r.entered}/{subs.length} subjects entered{r.decision ? " · " : ""}{r.decision && <b style={{color:r.decision==="Promoted"?C.green:C.red}}>{r.decision}</b>}</div>
            </div>
            <div style={{textAlign:"right"}}>
              <div style={{fontWeight:900,fontSize:15,color:avgColor(r.avg)}}>{r.avg.toFixed(2)} · {r.grade}</div>
              <div style={{fontSize:10.5,fontWeight:700,color:r.passed===r.entered?C.green:C.gold}}>{r.passed}/{r.entered} passed</div>
            </div>
          </div>
          {r.failed.length>0 && <div style={{fontSize:10.5,color:C.red,marginTop:5}}>Below {PASS_MARK}: {r.failed.join(", ")}</div>}
        </div>
      ))}
    </div>
  );
}

// Subject results: a teacher (read-only) sees how a subject is doing in each class they teach it in. Admin sees every subject.
function SubjectResultsPage({ ctx }) {
  const { students, marksMap, teachers, auth, settings } = ctx;
  const isTeacher = auth.role==="teacher";
  const teacher   = findTeacher(teachers, auth.user.email);
  const combos = useMemo(() => {
    const out = [];
    (isTeacher ? (teacher?.forms||[]) : FORMS).forEach(form => {
      const have = SUBJECTS_BY_FORM[form] || [];
      (isTeacher ? teacherSubjectsIn(teacher, form) : have).filter(sb => have.includes(sb)).forEach(sub => out.push({ form, sub }));
    });
    return out;
  }, [teachers, auth.user.email, isTeacher]);
  const subjectList = [...new Set(combos.map(c=>c.sub))].sort();
  const [f, setF] = useState({ sub:"", form:"", year:ACTIVE_YEAR, period:"First Term" });
  const subV  = subjectList.includes(f.sub) ? f.sub : subjectList[0];
  const formOpts = combos.filter(c=>c.sub===subV).map(c=>c.form);
  const formV = formOpts.includes(f.form) ? f.form : formOpts[0];
  const f5 = isF5(formV);
  const periodOpts = f5 ? F5_ASSESSMENTS : PERIODS;
  const periodV = periodOpts.includes(f.period) ? f.period : periodOpts[0];
  const isGce = f5 && periodV===GCE_SEQ;

  const data = useMemo(() => {
    if (!subV || !formV) return null;
    const pool = students.filter(x => x.active && x.reg_status==="registered" && x.form===formV && f5Writes(settings, x, subV)).sort((a,b)=>a.name.localeCompare(b.name));
    const rows = pool.map(st => {
      let v = null, grade = "—", pass = false, note = "";
      if (f5) {
        const raw = marksMap[`${st.id}-${subV}-${periodV}-${f.year}`]?.score;
        v = raw===null||raw===undefined ? null : Number(raw);
        if (v!==null) {
          if (isGce) { grade = gceFromPoints(v); pass = gcePassed(v); }
          else { const g = f5Grade(settings, v); grade = g.grade; pass = g.pass;
                 note = f5PapersFor(subV).map(pp => { const x = marksMap[`${st.id}-${subV}-${periodV} ${pp.k}-${f.year}`]?.score; return x!==undefined&&x!==null ? "P"+pp.k.slice(1)+" "+x : ""; }).filter(Boolean).join(" · "); }
        }
      } else {
        v = subjectPeriodScore(marksMap, st, subV, f.year, periodV);
        if (v!==null) { grade = scoreToGrade(v).grade; pass = v>=PASS_MARK; }
      }
      return { st, v, grade, pass, note };
    });
    const done = rows.filter(r=>r.v!==null).sort((a,b)=>b.v-a.v);
    const missing = rows.filter(r=>r.v===null);
    const dist = {}; done.forEach(r => { dist[r.grade] = (dist[r.grade]||0)+1; });
    const nums = done.map(r=>r.v);
    return { done, missing, dist, total:rows.length,
      passed: done.filter(r=>r.pass).length,
      avg: !isGce && nums.length ? nums.reduce((a,b)=>a+b,0)/nums.length : null,
      hi: !isGce && nums.length ? Math.max(...nums) : null, lo: !isGce && nums.length ? Math.min(...nums) : null };
  }, [students, marksMap, settings, subV, formV, f.year, periodV]);

  if (isTeacher && !combos.length) return <TeacherSetupNotice teacher={teacher} email={auth.user.email} needSubjects/>;
  if (!subV) return <Empty text="No subjects to show."/>;
  const card = { background:C.white, borderRadius:10, padding:12, marginBottom:11, boxShadow:"0 1px 3px rgba(0,0,0,0.06)" };
  const gradeKeys = f5 ? GCE_GRADES : GRADE_ORDER;
  const passRate = data && data.done.length ? Math.round(data.passed/data.done.length*100) : null;
  const weak = data ? data.done.filter(r=>!r.pass).slice().reverse() : [];
  function excel() {
    if (!data?.done.length) { alert("Nothing to export yet."); return; }
    exportToExcel(data.done.map((r,i) => ({ Rank:i+1, Name:r.st.name, Matricule:r.st.id, [isGce?"Grade":"Score /20"]: isGce ? r.grade : Number(r.v.toFixed(2)), Grade:r.grade, Result:r.pass?"Pass":"Fail" })),
      `SBCC-${subV.replace(/[^A-Za-z]/g,"")}-${formV.replace(/\s+/g,"")}-${periodV.replace(/\s+/g,"")}.xlsx`, subV.slice(0,31));
  }
  return (
    <div>
      <div style={card}>
        <label style={lbl}>Subject</label>
        <select style={inp} value={subV} onChange={e=>setF(x=>({...x,sub:e.target.value}))}>{subjectList.map(x=><option key={x}>{x}</option>)}</select>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginTop:8}}>
          <div><label style={lbl}>Class</label><select style={inp} value={formV} onChange={e=>setF(x=>({...x,form:e.target.value}))}>{formOpts.map(x=><option key={x}>{x}</option>)}</select></div>
          <div><label style={lbl}>Academic year</label><select style={inp} value={f.year} onChange={e=>setF(x=>({...x,year:e.target.value}))}>{ACAD_YEARS.map(y=><option key={y}>{y}</option>)}</select></div>
        </div>
        <label style={{...lbl,marginTop:8}}>{f5?"Assessment":"Period"}</label>
        <select style={inp} value={periodV} onChange={e=>setF(x=>({...x,period:e.target.value}))}>{periodOpts.map(x=><option key={x}>{x}</option>)}</select>
      </div>

      {!data || !data.done.length ? <div style={card}><Empty text={`No ${subV} marks entered yet for ${formV}, ${periodV}.`}/></div> : <>
        <div style={card}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,textAlign:"center"}}>
            <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>{isGce?"Passed (A–C)":"Class average"}</div><div style={{fontSize:17,fontWeight:900,color:isGce?C.navy:avgColor(data.avg)}}>{isGce ? `${data.passed}/${data.done.length}` : fmt2(data.avg)}</div></div>
            <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>Pass rate</div><div style={{fontSize:17,fontWeight:900,color:passRate>=50?C.green:C.red}}>{passRate}%</div></div>
            <div style={{background:C.grayBg,borderRadius:8,padding:"7px 2px"}}><div style={{fontSize:10,color:C.gray}}>Marks in</div><div style={{fontSize:17,fontWeight:900,color:C.navy}}>{data.done.length}/{data.total}</div></div>
          </div>
          {!isGce && <div style={{fontSize:11,color:C.gray,marginTop:8,textAlign:"center"}}>Highest <b style={{color:C.green}}>{fmt1(data.hi)}</b> · Lowest <b style={{color:C.red}}>{fmt1(data.lo)}</b> · Pass mark {f5 ? "grade C" : PASS_MARK+"/20"}</div>}
          <div style={{fontSize:11,color:C.navyMid,marginTop:6,textAlign:"center"}}>{gradeKeys.filter(g=>data.dist[g]).map(g=>g+": "+data.dist[g]).join("  ·  ")}</div>
          <div style={{marginTop:9}}><SmBtn onClick={excel} color={C.navyMid}>📊 Excel</SmBtn></div>
        </div>

        {weak.length>0 && <div style={{...card,borderLeft:`4px solid ${C.red}`}}>
          <div style={{fontWeight:800,color:C.red,fontSize:13,marginBottom:6}}>Needs attention ({weak.length})</div>
          {weak.map(r => <div key={r.st.id} style={{display:"flex",justifyContent:"space-between",fontSize:12.5,padding:"4px 0",borderTop:`1px solid ${C.grayBg}`}}><span style={{color:C.navy,fontWeight:600}}>{r.st.name}</span><span style={{color:C.red,fontWeight:800}}>{isGce ? r.grade : r.v.toFixed(1)+" · "+r.grade}</span></div>)}
        </div>}

        <div style={card}>
          <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:6}}>All students (highest first)</div>
          {data.done.map((r,i) => (
            <div key={r.st.id} style={{display:"flex",alignItems:"center",gap:10,borderTop:`1px solid ${C.grayBg}`,padding:"6px 0"}}>
              <div style={{width:22,fontSize:11,color:C.gray,fontWeight:700}}>{i+1}</div>
              <div style={{flex:1,minWidth:0}}><div style={{fontSize:12.5,fontWeight:600,color:C.navy}}>{r.st.name}</div>{r.note && <div style={{fontSize:9.5,color:C.gray}}>{r.note}</div>}</div>
              <div style={{fontWeight:800,fontSize:13,color:r.pass?C.green:C.red}}>{isGce ? r.grade : r.v.toFixed(1)+" · "+r.grade}</div>
            </div>
          ))}
        </div>
        {data.missing.length>0 && <div style={card}>
          <div style={{fontWeight:800,color:C.gold,fontSize:13,marginBottom:4}}>No mark yet ({data.missing.length})</div>
          <div style={{fontSize:12,color:C.gray}}>{data.missing.map(r=>r.st.name).join(", ")}</div>
        </div>}
      </>}
    </div>
  );
}

function MarksControlPage({ ctx }) {
  const { students, teachers, marksMap, markDeadlines, markUnlocks, saveDeadline, setClassUnlocked } = ctx;
  const [tab,  setTab]  = useState("progress"); // progress | lock | history
  const [year, setYear] = useState(ACTIVE_YEAR);
  const [seq,  setSeq]  = useState("SEQ 1");
  const [open, setOpen] = useState({});
  const [dlInput, setDlInput] = useState("");
  const [busy, setBusy] = useState(false);
  const [histForm, setHistForm] = useState("");
  const [hist, setHist] = useState({ rows:[], loading:false, error:"" });

  const currentDeadline = markDeadlines.find(d=>d.acad_year===year && d.seq===seq)?.deadline || "";
  useEffect(() => { setDlInput(currentDeadline); }, [currentDeadline, year, seq]);

  const formsForSeq = F5_ASSESSMENTS.includes(seq) ? ["Form 5"] : FORMS.filter(f => !isF5(f));
  const progress = useMemo(() => formsForSeq.map(form => {
    const pool = students.filter(s=>s.active && s.reg_status==="registered" && s.form===form);
    const subs = (SUBJECTS_BY_FORM[form]||[]).map(sub => {
      const subPool = pool.filter(s => f5Writes(ctx.settings, s, sub)); // Form 5: only students writing the subject
      const entered = subPool.filter(s => marksMap[`${s.id}-${sub}-${seq}-${year}`]).length;
      const tch = teachers.filter(t=>t.active && teacherSubjectsIn(t, form).includes(sub)).map(t=>t.name);
      const status = !subPool.length ? "na" : entered>=subPool.length ? "done" : entered>0 ? "partial" : "none";
      return { sub, entered, total:subPool.length, teachers:tch, status };
    });
    return { form, n:pool.length, subs, done:subs.filter(x=>x.status==="done").length };
  }), [students, teachers, marksMap, seq, year, ctx.settings]);

  function copyMissing() {
    const lines = [];
    progress.forEach(p => {
      const miss = p.subs.filter(x=>x.status==="partial"||x.status==="none");
      if (!miss.length) return;
      lines.push(p.form);
      miss.forEach(x => lines.push(`- ${x.sub} (${x.teachers.join(", ")||"no teacher assigned"}): ${x.entered}/${x.total}`));
    });
    const text = lines.length ? `Missing ${seq} marks, ${year}:\n`+lines.join("\n") : "All marks are in.";
    navigator.clipboard?.writeText(text).then(()=>alert("Copied. Paste it into WhatsApp."), ()=>alert(text));
  }

  async function run(fn) {
    setBusy(true);
    try { await fn(); }
    catch(e) { alert("Error: "+e.message+"\n\nHave you run migration.sql in Supabase?"); }
    setBusy(false);
  }

  async function loadHistory() {
    setHist(h => ({ ...h, loading:true, error:"" }));
    const { data, error } = await supabase.from("marks_history").select("*").eq("acad_year", year).eq("seq", seq).order("changed_at", { ascending:false }).limit(300);
    setHist({ rows:data||[], loading:false, error:error?error.message:"" });
  }
  useEffect(() => { if (tab==="history") loadHistory(); }, [tab, year, seq]);

  const statusColor = { done:C.green, partial:C.gold, none:C.red, na:C.gray };
  const stuById = Object.fromEntries(students.map(s=>[s.id,s]));
  const histRows = hist.rows.filter(r => !histForm || stuById[r.student_id]?.form===histForm);

  return (
    <div>
      <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><label style={lbl}>Academic Year</label><select style={inp} value={year} onChange={e=>setYear(e.target.value)}>{ACAD_YEARS.map(y=><option key={y}>{y}</option>)}</select></div>
          <div><label style={lbl}>Sequence</label><select style={inp} value={seq} onChange={e=>setSeq(e.target.value)}>{ALL_SEQ_LABELS.map(x=><option key={x}>{x}</option>)}</select></div>
        </div>
        <div style={{display:"flex",gap:6,marginTop:10}}>
          {[["progress","Progress"],["lock","Locks"],["history","History"],["class","Class"],["honour","Honours"],["f5","Form 5"]].map(([k,l])=>(
            <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"8px 4px",borderRadius:8,border:`2px solid ${tab===k?C.navy:C.grayLight}`,background:tab===k?C.navy:C.white,color:tab===k?C.white:C.gray,fontWeight:700,fontSize:11.5,cursor:"pointer"}}>{l}</button>
          ))}
        </div>
      </div>

      {tab==="f5" && <F5Tools ctx={ctx}/>}
      {tab==="class" && <ClassSummaryPanel ctx={ctx}/>}
      {tab==="honour" && <HonourRoll ctx={ctx}/>}

      {tab==="progress" && <>
        <div style={{display:"flex",justifyContent:"flex-end",marginBottom:8}}><SmBtn onClick={copyMissing} color={C.navyMid}>📋 Copy missing list</SmBtn></div>
        {progress.map(p => (
          <div key={p.form} style={{background:C.white,borderRadius:10,marginBottom:9,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",overflow:"hidden"}}>
            <div onClick={()=>setOpen(o=>({...o,[p.form]:!o[p.form]}))} style={{padding:"11px 12px",cursor:"pointer"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div style={{fontWeight:800,color:C.navy,fontSize:13}}>{p.form} <span style={{fontWeight:400,color:C.gray,fontSize:11}}>· {p.n} students</span></div>
                <Pill color={p.done===p.subs.length&&p.n?C.green:C.gold}>{p.done}/{p.subs.length} subjects {open[p.form]?"▲":"▼"}</Pill>
              </div>
              <div style={{marginTop:7,background:C.grayBg,borderRadius:5,height:7,overflow:"hidden"}}>
                <div style={{width:`${p.subs.length?p.done/p.subs.length*100:0}%`,height:"100%",background:C.green}}/>
              </div>
            </div>
            {open[p.form] && p.subs.map((x,i) => (
              <div key={x.sub} style={{padding:"7px 12px",borderTop:`1px solid ${C.grayBg}`,display:"flex",justifyContent:"space-between",gap:8,background:i%2?C.grayBg:C.white}}>
                <div style={{minWidth:0}}>
                  <div style={{fontSize:12.5,fontWeight:600,color:C.navy}}>{x.sub}</div>
                  <div style={{fontSize:10.5,color:x.teachers.length?C.gray:C.red}}>{x.teachers.join(", ")||"No teacher assigned"}</div>
                </div>
                <Pill color={statusColor[x.status]}>{x.status==="na"?"—":`${x.entered}/${x.total}`}</Pill>
              </div>
            ))}
          </div>
        ))}
      </>}

      {tab==="lock" && <>
        <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <label style={lbl}>Deadline for {seq} ({year}) — teachers are locked out the day after</label>
          <div style={{display:"flex",gap:8}}>
            <input type="date" style={{...inp,flex:1}} value={dlInput} onChange={e=>setDlInput(e.target.value)}/>
            <Btn disabled={busy} onClick={()=>run(()=>saveDeadline(year, seq, dlInput))}>{dlInput?"Save":"Clear"}</Btn>
          </div>
          <div style={{fontSize:10.5,color:C.gray,marginTop:6}}>Admin is never locked out. Leave empty and save to remove the deadline.</div>
        </div>
        {formsForSeq.map(f => {
          const st = lockStatus(f, seq, year, markDeadlines, markUnlocks);
          return (
            <div key={f} style={{background:C.white,borderRadius:10,padding:"10px 12px",marginBottom:7,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",display:"flex",justifyContent:"space-between",alignItems:"center",gap:8}}>
              <div>
                <div style={{fontWeight:700,color:C.navy,fontSize:13}}>{f}</div>
                <div style={{fontSize:11,color:st.locked?C.red:C.gray}}>
                  {!st.deadline ? "Open (no deadline set)" : !st.passed ? `Open until ${fmtDate(st.deadline)}` : st.unlocked ? "Unlocked for correction" : `Locked since ${fmtDate(st.deadline)}`}
                </div>
              </div>
              {st.passed && (st.unlocked
                ? <SmBtn color={C.red} onClick={()=>run(()=>setClassUnlocked(f, year, seq, false))}>🔒 Lock again</SmBtn>
                : <SmBtn color={C.green} onClick={()=>run(()=>setClassUnlocked(f, year, seq, true))}>🔓 Unlock</SmBtn>)}
            </div>
          );
        })}
      </>}

      {tab==="history" && <>
        <div style={{marginBottom:9}}>
          <select style={{...inp,maxWidth:160}} value={histForm} onChange={e=>setHistForm(e.target.value)}><option value="">All forms</option>{FORMS.map(f=><option key={f}>{f}</option>)}</select>
        </div>
        {hist.loading && <Spin text="Loading history…"/>}
        {hist.error && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"10px 12px",color:C.red,fontSize:12,marginBottom:9}}>{hist.error}<br/>Run migration.sql in Supabase if you haven't.</div>}
        {!hist.loading && !hist.error && !histRows.length && <Empty text="No changes recorded for this sequence."/>}
        {histRows.map(r => {
          const stu = stuById[r.student_id];
          return (
            <div key={r.id} style={{background:C.white,borderRadius:10,padding:"9px 12px",marginBottom:6,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
              <div style={{display:"flex",justifyContent:"space-between",gap:8}}>
                <div style={{fontWeight:700,color:C.navy,fontSize:12.5}}>{stu?.name||r.student_id} <span style={{fontWeight:400,color:C.gray}}>· {stu?.form}</span></div>
                <div style={{fontWeight:800,fontSize:13,color:C.navy,whiteSpace:"nowrap"}}>{r.old_score??"—"} → {r.new_score??"deleted"}</div>
              </div>
              <div style={{fontSize:11,color:C.gray}}>{r.subject} · by {r.changed_by_name||"unknown"} · {new Date(r.changed_at).toLocaleString("en-GB",{timeZone:"Africa/Douala",day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"})}</div>
            </div>
          );
        })}
      </>}
    </div>
  );
}

// ─── Class Remarks: the class teacher writes conduct + remark for every student ───
const CONDUCT_OPTIONS = ["Excellent","Very Good","Good","Satisfactory","Needs Improvement"];

function ClassRemarksPage({ ctx }) {
  const { students, teachers, auth, remarksMap, saveRemarks } = ctx;
  const isTeacher = auth.role==="teacher";
  const teacher   = findTeacher(teachers, auth.user.email);
  const myForms   = isTeacher ? (teacher?.class_teacher_of ? [teacher.class_teacher_of] : []) : FORMS;
  const [sel, setSel] = useState({ form:myForms[0]||"Form 1", year:ACTIVE_YEAR, term:"First Term" });
  const [local, setLocal] = useState({});
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);
  useEffect(() => { if (myForms.length && !myForms.includes(sel.form)) setSel(x => ({ ...x, form:myForms[0] })); }, [myForms.join("|")]);

  const pool = students.filter(s => s.active && s.reg_status==="registered" && s.form===sel.form);
  useEffect(() => {
    const init = {};
    pool.forEach(s => { const r = remarksMap[`${s.id}|${sel.year}|${sel.term}`]; init[s.id] = { conduct:r?.conduct||"", remark:r?.remark||"" }; });
    setLocal(init); setSaved(false);
  }, [sel.form, sel.year, sel.term, students.length]);

  if (isTeacher && !myForms.length) return (
    <div style={{background:"#fffbeb",border:"1px solid #fcd34d",borderRadius:10,padding:12,fontSize:12.5,color:"#78350f"}}>
      <strong>⚠️ Not a class teacher yet.</strong> Ask the administrator to open Teachers, edit your profile and choose the form you are class teacher of.
    </div>
  );

  const setField = (id, k, v) => { setLocal(p => ({ ...p, [id]: { ...p[id], [k]:v } })); setSaved(false); };
  async function saveAll() {
    const rows = pool
      .filter(s => local[s.id] && (local[s.id].conduct || local[s.id].remark.trim()))
      .map(s => ({ student_id:s.id, acad_year:sel.year, term:sel.term, conduct:local[s.id].conduct||null, remark:local[s.id].remark.trim()||null }));
    if (!rows.length) { alert("Nothing to save."); return; }
    setSaving(true);
    try { await saveRemarks(rows); setSaved(true); }
    catch(e) { alert("Save error: "+e.message+"\n\nHave you run migration3.sql in Supabase?"); }
    setSaving(false);
  }
  const done = pool.filter(s => local[s.id]?.remark?.trim()).length;

  return (
    <div>
      <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><label style={lbl}>Form</label><select style={inp} value={sel.form} onChange={e=>setSel(x=>({...x,form:e.target.value}))}>{myForms.map(f=><option key={f}>{f}</option>)}</select></div>
          <div><label style={lbl}>Academic Year</label><select style={inp} value={sel.year} onChange={e=>setSel(x=>({...x,year:e.target.value}))}>{ACAD_YEARS.map(y=><option key={y}>{y}</option>)}</select></div>
        </div>
        <div style={{display:"flex",gap:6,marginTop:10}}>
          {TERMS.map(t => (
            <button key={t} onClick={()=>setSel(x=>({...x,term:t}))} style={{flex:1,padding:"8px 4px",borderRadius:8,border:`2px solid ${sel.term===t?C.navy:C.grayLight}`,background:sel.term===t?C.navy:C.white,color:sel.term===t?C.white:C.gray,fontWeight:700,fontSize:11,cursor:"pointer"}}>{t}</button>
          ))}
        </div>
        <div style={{fontSize:10.5,color:C.gray,marginTop:8}}>Appears on the report card. The Third Term remark is also used on the annual report.</div>
      </div>

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8,gap:8}}>
        <span style={{fontSize:12,color:C.gray}}>{done}/{pool.length} remarks written</span>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          {saved && <span style={{fontSize:12,color:C.green,fontWeight:700}}>✓ Saved</span>}
          <Btn onClick={saveAll} disabled={saving}>{saving?"Saving…":"Save all"}</Btn>
        </div>
      </div>

      {!pool.length ? <Empty text="No registered students in this form."/> : pool.map(s => (
        <div key={s.id} style={{background:C.white,borderRadius:10,padding:"10px 12px",marginBottom:7,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,marginBottom:7}}>
            <div style={{minWidth:0}}><div style={{fontWeight:700,color:C.navy,fontSize:13}}>{s.name}</div><div style={{fontSize:10,fontFamily:"monospace",color:C.gray}}>{s.id}</div></div>
            <select style={{...inp,width:140,padding:"6px 8px",fontSize:12}} value={local[s.id]?.conduct||""} onChange={e=>setField(s.id,"conduct",e.target.value)}>
              <option value="">Conduct…</option>{CONDUCT_OPTIONS.map(o=><option key={o}>{o}</option>)}
            </select>
          </div>
          <input style={inp} placeholder="Class teacher's remark" value={local[s.id]?.remark||""} onChange={e=>setField(s.id,"remark",e.target.value)}/>
        </div>
      ))}
    </div>
  );
}

// ─── Attendance ────────────────────────────────────────────────────────────────
const ATT_STATUSES = ["present","absent","late"];
const ATT_LABELS = { present:"Present", absent:"Absent", late:"Late" };
const ATT_COLORS = { present:C.green, absent:C.red, late:C.gold };
const ATT_ICONS  = { present:"✓", absent:"✕", late:"⏰" };

function AttendancePage({ ctx }) {
  const { students, attendanceMap, auth, teachers, saveAttendanceBulk } = ctx;
  const isTeacher  = auth.role==="teacher";
  const teacher    = findTeacher(teachers, auth.user.email);
  const availForms = isTeacher ? (teacher?.forms||[]) : FORMS;

  const [tab,   setTab]   = useState("mark"); // "mark" | "summary"
  const [form,  setForm]  = useState(availForms[0]||(isTeacher?"":"Form 1"));
  useEffect(() => { if (isTeacher && availForms.length && !availForms.includes(form)) setForm(availForms[0]); }, [availForms.join("|")]);
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
      {isTeacher && <TeacherSetupNotice teacher={teacher} email={auth.user.email}/>}
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

// ─── Report Cards ──────────────────────────────────────────────────────────────
// Shows the report card exactly as it will print: one A4 portrait page (210 × 297 mm), shrunk to fit the screen
function ReportSheet({ html }) {
  const wrapRef = useRef(null);
  const [fit, setFit]   = useState(0.45);
  const [zoom, setZoom] = useState(false);
  useEffect(() => {
    const el = wrapRef.current; if (!el) return;
    const update = () => setFit(Math.min(1, Math.max(0.2, el.clientWidth / 794)));
    update();
    const ro = typeof ResizeObserver !== "undefined" ? new ResizeObserver(update) : null;
    ro?.observe(el); window.addEventListener("resize", update);
    return () => { ro?.disconnect(); window.removeEventListener("resize", update); };
  }, []);
  const scale = zoom ? 1 : fit;
  return (
    <div style={{marginBottom:13}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:7}}>
        <span style={{fontSize:11.5,color:C.gray}}>A4 portrait · exactly as it will print</span>
        <SmBtn onClick={()=>setZoom(z=>!z)} color={C.navyMid}>{zoom ? "Fit to screen" : "🔍 Zoom in"}</SmBtn>
      </div>
      <div ref={wrapRef} style={{width:"100%",overflowX:zoom?"auto":"hidden",background:"#e5e7eb",borderRadius:8,padding:zoom?0:"0"}}>
        <div style={{width:794*scale,height:1123*scale,margin:"0 auto",background:"#fff",boxShadow:"0 2px 12px rgba(0,0,0,0.2)"}}>
          <iframe title="Report card preview" srcDoc={html} scrolling="no"
            style={{width:794,height:1123,border:"none",display:"block",transform:`scale(${scale})`,transformOrigin:"top left",pointerEvents:"none"}}/>
        </div>
      </div>
    </div>
  );
}

// Report cards are for the administrator only (teachers never see this page)
function ReportsPage({ ctx }) {
  if (ctx.auth?.role !== "admin") return <Empty text="Report cards are available to the administrator only."/>;
  return <ReportsPageInner ctx={ctx}/>;
}


// ─── Report-card engine (shared by single print, whole-class print and WhatsApp PDF) ───
const SBC_SUBJECTS_LIST = ALL_SUBJECTS; // same master list as everywhere else

function attendanceTotals(student, attendanceMap) {
  let present=0, absent=0, late=0;
  if (!student) return { present, absent, late };
  Object.keys(attendanceMap||{}).forEach(key => {
    if (!key.startsWith(student.id+"-")) return;
    const st = attendanceMap[key].status;
    if (st==="present") present++; else if (st==="absent") absent++; else if (st==="late") late++;
  });
  return { present, absent, late };
}

// Everything the card needs for one student. env = { marksMap, teachers, ... }
function computeReport(student, env, sel) {
  const { marksMap, teachers, settings } = env;
  const f5       = isF5(student.form);
  const assess   = sel.f5Assess || F5_ASSESSMENTS[0];
  // Form 5: only the subjects this student writes; everyone else: the whole form's list
  const subjects = f5 ? f5SubjectsOf(settings, student) : (SUBJECTS_BY_FORM[student.form] || []);
  const termSeqs = f5 ? [assess] : TERM_SEQS[sel.term];
  const isTerm   = sel.mode==="term";
  const getScore = (sub,seq) => marksMap[`${student.id}-${sub}-${seq}-${sel.year}`]?.score ?? null;
  const getC     = sub => marksMap[`${student.id}-${sub}-${f5?assess:"SEQ 1"}-${sel.year}`]?.coeff || getCoeff(sub);
  const teacherFor = sub => teachers.find(t=>t.active&&teacherSubjectsIn(t, student.form).includes(sub))?.name || "—";
  const avg2 = arr => { const v=arr.filter(x=>x!==null); return v.length ? v.reduce((a,b)=>a+Number(b),0)/v.length : null; };

  const subjectRows = subjects.map(sub => {
    if (f5) {
      // one assessment at a time: the score (or GCE grade points) is the result for that subject
      const papers = {};
      f5PapersFor(sub).forEach(pp => { papers[pp.k] = getScore(sub, `${assess} ${pp.k}`); });
      const comb = assess===GCE_SEQ ? null : f5Combine(sub, papers);
      const sc = comb!==null ? comb : getScore(sub, assess);
      return { sub, coeff:getC(sub), teacher:teacherFor(sub), s1:null,s2:null,s3:null,s4:null,s5:null,s6:null, t1Avg:sc,t2Avg:null,t3Avg:null, annualAvg:sc, termScores:[sc], termAvg:sc, papers };
    }
    const s1=getScore(sub,"SEQ 1"),s2=getScore(sub,"SEQ 2"),s3=getScore(sub,"SEQ 3"),s4=getScore(sub,"SEQ 4"),s5=getScore(sub,"SEQ 5"),s6=getScore(sub,"SEQ 6");
    const t1Avg=avg2([s1,s2]), t2Avg=avg2([s3,s4]), t3Avg=avg2([s5,s6]);
    const annualAvg  = avg2([t1Avg,t2Avg,t3Avg]);
    const termScores = termSeqs.map(sq=>[s1,s2,s3,s4,s5,s6][parseInt(sq.replace("SEQ ",""))-1]);
    return { sub, coeff:getC(sub), teacher:teacherFor(sub), s1,s2,s3,s4,s5,s6, t1Avg,t2Avg,t3Avg, annualAvg, termScores, termAvg:avg2(termScores) };
  });

  const pick    = r => (f5 || isTerm) ? r.termAvg : r.annualAvg;
  const modeRows = subjectRows.filter(r => pick(r)!==null);
  const wAvg = (rows, f) => { let tw=0,tc=0; rows.forEach(r=>{ const a=f(r); if(a!==null){tw+=a*r.coeff;tc+=r.coeff;} }); return tc ? tw/tc : null; };
  const overallAvg   = wAvg(modeRows, pick);
  const annualOverall = wAvg(subjectRows.filter(r=>r.annualAvg!==null), r=>r.annualAvg);
  const totalCoeff   = modeRows.reduce((a,r)=>a+r.coeff,0);
  const overallGrd   = overallAvg!==null ? (f5 ? { ...f5Grade(settings, overallAvg), remark:f5Grade(settings, overallAvg).grade } : scoreToGrade(overallAvg)) : { grade:"—", remark:"—" };
  return { subjectRows, modeRows, overallAvg, annualOverall, totalCoeff, overallGrd };
}

function reportCardParts(student, env, sel, conduct) {
    const { students, marksMap, feesMap, settings } = env;
    const { modeRows, overallAvg, totalCoeff, overallGrd, annualOverall } = computeReport(student, env, sel);
    const assess = sel.f5Assess || F5_ASSESSMENTS[0];
    const isGce  = isF5(student.form) && assess===GCE_SEQ;
    const paperMode = isF5(student.form) && !isGce; // Form 5 mock cards show Paper 1 / 2 / 3
    const f5 = isF5(student.form);
    const termSeqs = f5 ? [assess] : TERM_SEQS[sel.term];

    // Auto-filled boxes at the bottom of the card
    const hasAnnual = !f5 && (sel.mode!=="term" || sel.term==="Third Term");
    const decision  = hasAnnual && annualOverall!==null ? (annualOverall>=PASS_MARK ? "Promoted" : "Repeat") : "";
    const owing     = TOTAL_FEE - (feesMap?.[student.id]?.paid || 0);
    const owingTxt  = env.showFees===false ? "" : owing>0 ? owing.toLocaleString()+" FCFA" : "Nil"; // no fee data for admins without fee access
    const nextBegin = settings?.next_term_begins ? fmtDate(settings.next_term_begins) : "";
    const fill = v => v ? "<b>"+escHtml(v)+"</b>" : "________";

    const SBC_SUBJECTS = SBC_SUBJECTS_LIST;
    const isTerm = f5 ? true : sel.mode==="term";

    // ── Class ranking (same rules as before) ─────────────────────────────────
    const peers = students.filter(s=>s.active&&s.form===student.form&&s.reg_status==="registered");
    const peerAvgs = peers.map(ps=>{
      let tw=0,tc=0;
      SBC_SUBJECTS.forEach(sub=>{
        if (f5 && !f5Writes(settings, ps, sub)) return;
        const seqs = f5
          ? [assess].map(sq=>{ const k=ps.id+"-"+sub+"-"+sq+"-"+sel.year; return marksMap[k]?.score??null; })
          : isTerm
          ? TERM_SEQS[sel.term].map(sq=>{ const k=ps.id+"-"+sub+"-"+sq+"-"+sel.year; return marksMap[k]?.score??null; })
          : SEQ_LABELS.map(sq=>{ const k=ps.id+"-"+sub+"-"+sq+"-"+sel.year; return marksMap[k]?.score??null; });
        const vals=seqs.filter(x=>x!==null);
        const c=getCoeff(sub);
        if(vals.length){tw+=vals.reduce((a,b)=>a+Number(b),0)/vals.length*c;tc+=c;}
      });
      return{id:ps.id,avg:tc?tw/tc:0};
    }).sort((a,b)=>b.avg-a.avg);
    const myRank  = peerAvgs.findIndex(x=>x.id===student.id)+1;
    const rankStr = isGce ? "—" : myRank>0 ? myRank+" / "+peers.length : "—";
    const classAvg = peerAvgs.length ? (peerAvgs.reduce((a,x)=>a+x.avg,0)/peerAvgs.length).toFixed(2) : "—";

    // ── Table columns (percent widths add up to 100) ─────────────────────────
    const seqNums = f5 ? [assess] : isTerm ? termSeqs.map(sq=>sq.replace("SEQ ","")) : ["1","2","3","4","5","6"];
    const wSeq = f5 ? 12 : isTerm ? 9 : 5.6, wAvg = f5 ? 10 : isTerm ? 9 : 7, wCoef = isTerm ? 8 : 6;
    const wScore = isTerm ? 10 : 8, wRmk = isTerm ? 13 : 11, wTch = f5 ? 22 : isTerm ? 18 : 14;
    const wSub = +(100 - seqNums.length*wSeq - wAvg - wCoef - wScore - wRmk - wTch).toFixed(1);
    const cols = [wSub, wTch, ...seqNums.map(()=>wSeq), wAvg, wCoef, wScore, wRmk].map(w=>`<col style="width:${w}%">`).join("");
    const colsP = [18,14,7,7,7,8,7,6,9,17].map(w=>`<col style="width:${w}%">`).join("");
    const paperHead = `<thead><tr><th style="text-align:left;padding-left:1.5mm">SUBJECT</th><th>TEACHER</th><th>P1 /50</th><th>P2 /100</th><th>P3 /100</th><th>TOTAL /20</th><th>GRADE</th><th>COEFF</th><th>WEIGHTED</th><th>REMARK</th></tr></thead>`;
    const seqHeaders = seqNums.map(n=> f5 ? `<th>${isGce?"GRADE":"SCORE /20"}</th>` : `<th>SQ${n}</th>`).join("");

    const findRow = sub => modeRows.find(x=>
      x.sub===sub ||
      (sub==="French/ Français"&&x.sub==="French") ||
      (sub==="Computer Studies"&&x.sub==="Computer Science") ||
      (sub==="Sport/Physical education"&&x.sub==="Physical Education") ||
      (sub==="Literature in English"&&x.sub==="Literature")
    );

    let sumScore = 0;
    const f5Rows = !f5 ? "" : modeRows.map(r=>{
      const sc = r.termAvg;
      const g  = isGce ? { grade:gceFromPoints(sc), pass:gcePassed(sc), remark:gcePassed(sc)?"Pass":"Fail" } : f5Grade(settings, sc);
      const pond = !isGce && sc!==null ? sc*r.coeff : null;
      if (pond!==null) sumScore += pond;
      const col = g.pass ? "#15803D" : "#B91C1C";
      if (paperMode) {
        const pp = r.papers || {};
        const cell = k => pp[k]!=null ? pp[k] : (f5PapersFor(r.sub).some(q=>q.k===k) ? "" : "—");
        const tnP = r.teacher && !/^[?\u2014\u2013-]$/.test(String(r.teacher).trim()) ? r.teacher : "";
        return `<tr><td class="s">${escHtml(r.sub)}</td><td class="t">${escHtml(tnP)}</td>
        <td class="c">${cell("P1")}</td><td class="c">${cell("P2")}</td><td class="c">${cell("P3")}</td>
        <td class="c b">${Number(sc).toFixed(2)}</td>
        <td class="c b" style="color:${col}">${escHtml(g.grade)}</td>
        <td class="c">${r.coeff}</td>
        <td class="c b">${pond!==null?pond.toFixed(1):""}</td>
        <td class="c" style="color:${col};font-weight:700">${escHtml(g.remark)}</td></tr>`;
      }
      const tn  = r.teacher && !/^[?\u2014\u2013-]$/.test(String(r.teacher).trim()) ? r.teacher : "";
      return `<tr><td class="s">${escHtml(r.sub)}</td><td class="t">${escHtml(tn)}</td>
        <td class="c b">${isGce ? escHtml(g.grade) : sc}</td>
        <td class="c b" style="color:${col}">${escHtml(g.grade)}</td>
        <td class="c">${r.coeff}</td>
        <td class="c b">${pond!==null?pond.toFixed(1):""}</td>
        <td class="c" style="color:${col};font-weight:700">${escHtml(g.remark)}</td></tr>`;
    }).join("");
    const rows0 = SBC_SUBJECTS.map(sub=>{
      const r = findRow(sub);
      const empty = `<td class="c"></td>`;
      if(!r){
        return `<tr><td class="s">${escHtml(sub)}</td><td class="t"></td>${seqNums.map(()=>empty).join("")}${empty}${empty}${empty}${empty}</tr>`;
      }
      const avg = isTerm ? r.termAvg : r.annualAvg;
      const {remark} = scoreToGrade(avg);
      const pass = avg!==null && avg>=10;
      const pond = avg!==null ? avg*r.coeff : null;
      if (pond!==null) sumScore += pond;
      const scoreVals = isTerm ? r.termScores : [r.s1,r.s2,r.s3,r.s4,r.s5,r.s6]; // Form 5 uses termScores (3 assessments)
      const scoreCols = scoreVals.map((s,qi)=> (f5 && F5_ASSESSMENTS[qi]===GCE_SEQ)
        ? `<td class="c b" style="color:${gcePassed(s)?"#15803D":"#B91C1C"}">${s!=null?escHtml(gceFromPoints(s)):""}</td>`
        : `<td class="c">${s!=null?s:""}</td>`).join("");
      const teacherName = r.teacher && !/^[?\u2014\u2013-]$/.test(String(r.teacher).trim()) ? r.teacher : "";
      return `<tr>
        <td class="s">${escHtml(sub)}</td>
        <td class="t">${escHtml(teacherName)}</td>
        ${scoreCols}
        <td class="c b">${avg!=null?avg.toFixed(1):""}</td>
        <td class="c">${r.coeff}</td>
        <td class="c b">${pond!==null?pond.toFixed(1):""}</td>
        <td class="c" style="color:${pass?"#15803D":"#B91C1C"};font-weight:700">${avg!=null?escHtml(remark):""}</td>
      </tr>`;
    }).join("");
    const rows = f5 ? f5Rows : rows0;
    const f5Passed = f5 ? modeRows.filter(r => (isGce ? gcePassed(r.termAvg) : f5Grade(settings, r.termAvg).pass)).length : 0;

    const totalRow = paperMode ? `<tr class="tot">
      <td colspan="5" style="text-align:right;padding-right:3mm">TOTAL / AVERAGE</td>
      <td class="c">${overallAvg!==null?overallAvg.toFixed(2):""}</td>
      <td class="c">${overallAvg!==null ? f5Grade(settings, overallAvg).grade : ""}</td>
      <td class="c">${totalCoeff||""}</td>
      <td class="c">${sumScore?sumScore.toFixed(1):""}</td>
      <td class="c"></td>
    </tr>` : f5 ? `<tr class="tot">
      <td colspan="3" style="text-align:right;padding-right:3mm">${isGce ? "SUBJECTS PASSED (A–C)" : "TOTAL / AVERAGE"}</td>
      <td class="c">${isGce ? "" : (overallAvg!==null ? f5Grade(settings, overallAvg).grade : "")}</td>
      <td class="c">${isGce ? "" : (totalCoeff||"")}</td>
      <td class="c">${isGce ? "" : (sumScore?sumScore.toFixed(1):"")}</td>
      <td class="c">${isGce ? f5Passed+" / "+modeRows.length : (overallAvg!==null?overallAvg.toFixed(2):"")}</td>
    </tr>` : `<tr class="tot">
      <td colspan="${2+seqNums.length}" style="text-align:right;padding-right:3mm">TOTAL / AVERAGE</td>
      <td class="c">${overallAvg!==null?overallAvg.toFixed(2):""}</td>
      <td class="c">${totalCoeff||""}</td>
      <td class="c">${sumScore?sumScore.toFixed(1):""}</td>
      <td class="c">${overallAvg!==null?escHtml(overallGrd.remark):""}</td>
    </tr>`;

    const photoHtml = student.photo_url
      ? `<img src="${student.photo_url}" alt="" style="width:22mm;height:28mm;object-fit:cover;border:1px solid #1a56a0;display:block;margin:0 auto">`
      : `<div style="width:22mm;height:28mm;border:1px solid #999;margin:0 auto;display:flex;align-items:center;justify-content:center;font-size:7pt;color:#999">PHOTO</div>`;

    const termLabel = f5 ? assess.toUpperCase() : isTerm ? sel.term.replace(" Term","").toUpperCase()+" TERM" : "ANNUAL REPORT";
    const blank = "&nbsp;";
    const avgLine  = (label, on) => `${label}: <b>${on&&overallAvg!==null ? overallAvg.toFixed(2)+" / 20" : "—"}</b>`;

    const css = `
  *{box-sizing:border-box;margin:0;padding:0}
  @page{size:A4 portrait;margin:8mm}
  html,body{font-family:Arial,Helvetica,sans-serif;color:#000;background:#fff;
    -webkit-print-color-adjust:exact;print-color-adjust:exact}
  @media print{ html,body{width:194mm} }
  @media screen{ html,body{width:210mm;height:297mm} body{padding:8mm} }
  .sheet{width:194mm;height:276mm;display:flex;flex-direction:column;gap:2mm;overflow:hidden;break-inside:avoid;page-break-after:avoid}
  table{width:100%;border-collapse:collapse}
  .gov{font-size:7.5pt;line-height:1.45;vertical-align:top;width:34%}
  .info td{border:1px solid #1a56a0;padding:0 2mm;height:5.6mm;font-size:8.5pt}
  .mk{table-layout:fixed;border:1px solid #1a56a0}
  .mk th{background:#1a56a0;color:#fff;border:1px solid #fff;height:6.2mm;font-size:7.5pt;text-align:center}
  .mk td{border:1px solid #1a56a0;height:5.5mm;padding:0 1.2mm;font-size:8pt;overflow:hidden;white-space:nowrap}
  .mk tbody tr:nth-child(even) td{background:#F5F8FC}
  .mk td.c{text-align:center} .mk td.b{font-weight:700} .mk td.s{text-align:left}
  .mk td.t{font-size:6.5pt;font-style:italic;color:#1a56a0;text-overflow:ellipsis}
  .mk tr.tot td{background:#E8F0FA!important;font-weight:800;font-size:8.5pt;border-top:2px solid #1a56a0}
  .sum td{border:1px solid #000;padding:0 2.5mm;height:6.2mm;font-size:8.5pt}
  .sum td.rem{height:8mm;vertical-align:top;padding-top:1mm}
  .sig div.line{height:10mm;border-bottom:1px solid #000;margin-bottom:1mm}
  .sig td{text-align:center;padding:0 5mm;font-size:8pt}
`;
    const sheet = `<div class="sheet">

  <table><tr>
    <td class="gov"><b>REPUBLIC OF CAMEROON</b><br>Peace – Work – Fatherland<br>Ministry of Secondary Education</td>
    <td style="text-align:center;vertical-align:middle"><img src="${SCHOOL_LOGO}" alt="SBCC" style="width:19mm;height:19mm;object-fit:contain"></td>
    <td class="gov" style="text-align:right"><b>REPUBLIQUE DU CAMEROUN</b><br>Paix – Travail – Patrie<br>Ministère De l'Enseignement</td>
  </tr></table>

  <div style="background:#1a56a0;color:#fff;text-align:center;padding:2mm 2mm">
    <div style="font-size:14.5pt;font-weight:900;letter-spacing:.4px">SAKER BAPTIST COMPREHENSIVE COLLEGE BAWE</div>
    <div style="font-size:8pt;margin-top:.5mm;opacity:.9">Motto: Quality, Discipline, and Excellence education</div>
  </div>

  <div style="text-align:center">
    <span style="font-size:12pt;font-weight:800;text-decoration:underline">ACADEMIC REPORT CARD</span>
    <span style="font-size:11pt;font-weight:700"> / BULLETIN DE NOTES</span>
  </div>

  <table class="info"><tr>
    <td colspan="2" style="background:#EEF4FB"><b>${termLabel}</b> &nbsp;|&nbsp; Class: <b>${escHtml(student.form)}</b> &nbsp;|&nbsp; No. on Roll: <b>${peers.length}</b> &nbsp;|&nbsp; Academic Year: <b>${escHtml(sel.year)}</b>${f5 ? ` &nbsp;|&nbsp; Subjects written: <b>${modeRows.length}</b>` : ""}</td>
    <td rowspan="5" style="width:27mm;text-align:center;vertical-align:middle;padding:1mm">${photoHtml}</td>
  </tr><tr>
    <td style="width:62%;height:7.5mm">Name / Nom: <b style="font-size:12.5pt;letter-spacing:.2px">${escHtml(student.name)}</b></td>
    <td>Matricule: <b>${escHtml(student.id)}</b></td>
  </tr><tr>
    <td>Sex / Sexe: <b>${escHtml(student.gender)}</b></td>
    <td>Date of birth: <b>${fmtDate(student.dob)}</b></td>
  </tr><tr>
    <td>Parent / Guardian: <b>${escHtml(student.parent||"—")}</b></td>
    <td>Phone: <b>${escHtml(student.phone||"—")}</b></td>
  </tr><tr>
    <td colspan="2">Address / Domicile: <b>${escHtml(student.address||"—")}</b></td>
  </tr></table>

  <table class="mk"><colgroup>${paperMode ? colsP : cols}</colgroup>
    ${paperMode ? paperHead : `<thead><tr><th style="text-align:left;padding-left:1.5mm">SUBJECT</th><th>TEACHER</th>${seqHeaders}${f5 ? "<th>Grade</th>" : "<th>Avg /20</th>"}<th>Coeff</th><th>${f5 ? "Weighted" : "Score"}</th><th>Remark</th></tr></thead>`}
    <tbody>${rows}${totalRow}</tbody>
  </table>

  <table class="sum">
    <tr>
      ${f5
        ? `<td style="width:25%">${isGce ? "Subjects passed: <b>"+f5Passed+" / "+modeRows.length+"</b>" : avgLine("Average", true)}</td>
      <td style="width:25%">Subjects written: <b>${modeRows.length}</b></td>
      <td style="width:25%">${isGce ? "&nbsp;" : "Class Avg: <b>"+classAvg+" / 20</b>"}</td>`
        : `<td style="width:25%">${avgLine("Terminal Avg", isTerm)}</td>
      <td style="width:25%">${avgLine("Annual Avg", !isTerm)}</td>
      <td style="width:25%">Class Avg: <b>${classAvg} / 20</b></td>`}
      <td style="width:25%">Position / Rang: <b>${rankStr}</b></td>
    </tr><tr>
      <td>Days Present: <b>${escHtml(conduct.present||"—")}</b></td>
      <td>Days Absent: <b>${escHtml(conduct.absent||"—")}</b></td>
      <td>Times Late: <b>${escHtml(conduct.late||"—")}</b></td>
      <td>Conduct: <b>${escHtml(conduct.conduct||"—")}</b></td>
    </tr><tr>
      <td>${f5 ? "Candidate No: "+fill(f5CandNo(settings, student)) : "Promoted / Repeat: "+fill(decision)}</td>
      <td>No. of Warnings: ________</td>
      <td>Fees Owing: ${fill(owingTxt)}</td>
      <td>Next Term Begins: ${fill(nextBegin)}</td>
    </tr><tr>
      <td colspan="4" class="rem">Class Teacher's Remark: <b>${escHtml(conduct.classTeacherRemark)||blank}</b></td>
    </tr><tr>
      <td colspan="4" class="rem">Principal's Remark: <b>${escHtml(conduct.principalRemark)||blank}</b></td>
    </tr>
  </table>

  <table class="sig" style="margin-top:2mm"><tr>
    <td style="width:33%"><div class="line"></div>Class Teacher / Prof. Principal</td>
    <td style="width:34%"><div class="line"></div>Principal's Signature / Stamp</td>
    <td style="width:33%"><div class="line"></div>Parent / Guardian</td>
  </tr></table>

  <div style="font-size:6.5pt;color:#444;text-align:center;line-height:1.4">
    ${f5 ? (isGce ? "GCE grades: A, B, C = Pass · D, E, U = Fail" : (()=>{ const k=f5Scale(settings); return "Papers: P1 MCQ /50 (40%) + P2 Structural /100 (60%) · Computer Studies: P1 30% + P2 40% + P3 Practical 30% · Total /20 · Grading: A "+k.A+"–20 · B "+k.B+"–"+k.A+" · C "+k.C+"–"+k.B+" · D "+k.D+"–"+k.C+" · E "+k.E+"–"+k.D+" · U below "+k.E+" · A, B, C = Pass"; })())
    : "Grading (out of 20): A Excellent 17–20 · B+ Very Good 15–16.9 · B Good 13–14.9 · C+ Fairly Good 11–12.9 · C Average 10–10.9 · D Pass 9–9.9 · F Fail below 9"}
  </div>
  <div style="margin-top:auto;font-size:7pt;color:#888;text-align:center;border-top:1px dashed #ccc;padding-top:1.2mm">
    Saker Baptist Comprehensive College Bawe · ${escHtml(sel.year)} · Printed: ${new Date().toLocaleDateString("en-GB")}
  </div>
</div>`;
    return { css, sheet };
}



function reportCardDoc(parts, title) {
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<title>${escHtml(title||"Report Card")}</title>
<style>${parts.css}</style></head><body>${parts.sheet}</body></html>`;
}

// Several cards in one printable document, one per A4 page
function reportCardBatchDoc(partsList, title) {
  const extra = `
@media print{.sheet{break-after:page;page-break-after:always}.sheet:last-child{break-after:auto;page-break-after:auto}}
@media screen{html,body{height:auto!important}.sheet{margin-bottom:8mm}}`;
  return reportCardDoc({ css: partsList[0].css + extra, sheet: partsList.map(p=>p.sheet).join("\n") }, title);
}

// Parent phone → international digits for wa.me / sms links (Cameroon: 9 digits → +237)
function waNumber(phone) {
  let d = String(phone||"").replace(/\D/g,"");
  if (d.startsWith("00")) d = d.slice(2);
  if (d.length===9) d = "237"+d;
  return d.length>=11 ? d : "";
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 10000);
}

// Renders full HTML pages (one A4 page each) into a real PDF file. Needs: npm i html2canvas jspdf
async function htmlDocsToPdfBlob(docs, onProgress, scale=2, quality=0.92) {
  const [{ default: html2canvas }, { jsPDF }] = await Promise.all([import("html2canvas"), import("jspdf")]);
  const pdf = new jsPDF({ unit:"mm", format:"a4", orientation:"portrait" });
  for (let i=0; i<docs.length; i++) {
    const frame = document.createElement("iframe");
    frame.style.cssText = "position:fixed;left:-10000px;top:0;width:794px;height:1123px;border:0";
    document.body.appendChild(frame);
    await new Promise(res => { frame.onload = res; frame.srcdoc = docs[i]; });
    const d = frame.contentDocument;
    await Promise.all([...d.images].map(im => im.complete ? 0 : new Promise(r => { im.onload = im.onerror = r; })));
    await new Promise(r => setTimeout(r, 120));
    const canvas = await html2canvas(d.documentElement, { scale, backgroundColor:"#ffffff", width:794, height:1123, windowWidth:794, windowHeight:1123, useCORS:true });
    if (i>0) pdf.addPage();
    pdf.addImage(canvas.toDataURL("image/jpeg", quality), "JPEG", 0, 0, 210, 297);
    canvas.width = canvas.height = 0; // free memory
    frame.remove();
    onProgress?.(i+1, docs.length);
    await new Promise(r => setTimeout(r, 30));
  }
  return pdf.output("blob");
}

function ReportsPageInner({ ctx }) {
  const { students, marksMap, teachers, attendanceMap, feesMap, settings, saveSetting, remarksMap } = ctx;
  const remTerm = sel0 => sel0.mode==="term" ? sel0.term : "Third Term";
  const [sel,     setSel]     = useState({ studentId:"", year:ACTIVE_YEAR, mode:"term", term:"First Term", f5Assess:"Pre Mock" });
  const [conduct, setConduct] = useState({ present:"", absent:"", late:"", conduct:"Very Good", classTeacherRemark:"", principalRemark:"" });
  const [showC,   setShowC]   = useState(false);
  const [sharing, setSharing] = useState(false);
  const [batchProg, setBatchProg] = useState(null); // null | {done,total}
  const [batchForm, setBatchForm] = useState("Form 1");
  const [nextTerm, setNextTerm] = useState(settings?.next_term_begins || "");
  const [savingNT, setSavingNT] = useState(false);

  const student  = students.find(s=>s.id===sel.studentId);
  const subjects = student ? (SUBJECTS_BY_FORM[student.form]||[]) : [];
  const termSeqs = TERM_SEQS[sel.term];

  // Auto-compute real attendance totals for this student from the Attendance page's data
  const attTotals = attendanceTotals(student, attendanceMap);

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

  // The fee administrator already has every fee row in feesMap. Other admins can't read the fees table,
  // so they get just the balances needed for the report cards from a database function.
  const [rcFees, setRcFees] = useState(null);
  useEffect(() => {
    if (ctx.auth.canFees) return;
    let alive = true;
    (async () => {
      const { data, error } = await supabase.rpc("fee_balances_for_report_cards");
      if (!alive || error) return;
      const fm = {};
      (data||[]).forEach(f => { fm[f.student_id] = { paid: f.paid||0, total: f.total||TOTAL_FEE }; });
      setRcFees(fm);
    })();
    return () => { alive = false; };
  }, [ctx.auth.canFees]);
  const rcFeesMap = ctx.auth.canFees ? feesMap : rcFees;
  const env = { students, marksMap, teachers, feesMap: rcFeesMap || {}, settings, showFees: !!(ctx.auth.canFees || rcFees) };
  const curRem = student ? remarksMap?.[`${student.id}|${sel.year}|${remTerm(sel)}`] : null;
  useEffect(() => {
    if (!student) return;
    setConduct(c => ({ ...c, conduct: curRem?.conduct || "Very Good", classTeacherRemark: curRem?.remark || "" }));
  }, [student?.id, sel.year, sel.term, sel.mode, curRem?.conduct, curRem?.remark]);
  const modeRows = student ? computeReport(student, env, sel).modeRows : [];
  const periodLabel = student && isF5(student.form) ? sel.f5Assess : sel.mode==="term" ? sel.term : "Annual";
  const buildPrintHTML = () => (!student || !modeRows.length) ? "" : reportCardDoc(reportCardParts(student, env, sel, conduct), "Report Card – "+student.name);

  // One document with every card of a class, alphabetical, one per A4 page
  async function printClass() {
    const pool = students.filter(s=>s.active && s.form===batchForm && s.reg_status==="registered");
    const list = []; let skipped = 0;
    pool.forEach(s => {
      if (!computeReport(s, env, sel).modeRows.length) { skipped++; return; }
      const att = attendanceTotals(s, attendanceMap);
      const rm = remarksMap?.[`${s.id}|${sel.year}|${remTerm(sel)}`];
      list.push(reportCardParts(s, env, sel, { present:String(att.present), absent:String(att.absent), late:String(att.late), conduct:rm?.conduct||"", classTeacherRemark:rm?.remark||"", principalRemark:"" }));
    });
    if (!list.length) { alert("No student in "+batchForm+" has marks for this period."); return; }
    if (skipped && !window.confirm(skipped+" student(s) have no marks for this period and will be left out. Print "+list.length+" report cards?")) return;
    setBatchProg({ done:0, total:list.length });
    try {
      const docs = list.map(p => reportCardDoc(p, "Report Card"));
      const blob = await htmlDocsToPdfBlob(docs, (done,total) => setBatchProg({ done, total }), 1.5, 0.85);
      downloadBlob(blob, "Report-Cards-"+batchForm.replace(/\s+/g,"")+"-"+(isF5(batchForm)?sel.f5Assess:periodLabel).replace(/\s+/g,"")+".pdf");
    } catch(e) { alert("Could not create the PDF: "+e.message); }
    setBatchProg(null);
  }

  async function shareWhatsApp() {
    const num = waNumber(student.phone);
    if (!num) { alert("This student has no valid parent phone number. Add it on the Students page first."); return; }
    setSharing(true);
    const win = window.open("", "_blank"); // opened now, while the tap still counts, so the browser allows it
    try {
      const blob = await htmlDocsToPdfBlob([buildPrintHTML()]);
      downloadBlob(blob, "Report-Card-"+student.name.replace(/[^A-Za-z0-9]+/g,"-")+"-"+periodLabel.replace(/\s+/g,"")+".pdf");
      const text = "Good day "+(student.parent||"Sir/Madam")+", please find attached "+student.name+"'s "+periodLabel+" report card ("+sel.year+") from Saker Baptist Comprehensive College Bawe.";
      const url = "https://wa.me/"+num+"?text="+encodeURIComponent(text);
      if (win) win.location.href = url; else window.location.href = url;
    } catch(e) {
      win?.close();
      alert("Could not create the PDF: "+e.message);
    }
    setSharing(false);
  }

  async function saveNextTerm() {
    setSavingNT(true);
    try { await saveSetting("next_term_begins", nextTerm || ""); alert("Saved."); }
    catch(e) { alert("Could not save: "+e.message+"\n\nRun migration.sql in Supabase first."); }
    setSavingNT(false);
  }

  function printReport() {
    if (!student||!modeRows.length) return;
    domPrint("sbc-report", buildPrintHTML(), "A4 portrait", "8mm");
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
          {student && isF5(student.form) && (
            <div style={{gridColumn:"1/-1"}}>
              <label style={lbl}>Assessment (Form 5)</label>
              <div style={{display:"flex",gap:6}}>
                {F5_ASSESSMENTS.map(t=>(
                  <button key={t} onClick={()=>setSel(x=>({...x,f5Assess:t}))} style={{flex:1,padding:"8px 4px",borderRadius:8,border:`2px solid ${sel.f5Assess===t?C.navy:C.grayLight}`,background:sel.f5Assess===t?C.navy:C.white,color:sel.f5Assess===t?C.white:C.gray,fontWeight:700,fontSize:11,cursor:"pointer"}}>{t}</button>
                ))}
              </div>
            </div>
          )}
          {!(student && isF5(student.form)) && <div><label style={lbl}>Report Type</label><select style={inp} value={sel.mode} onChange={e=>setSel(s=>({...s,mode:e.target.value}))}><option value="term">Term Report</option><option value="annual">Annual Report</option></select></div>}
          {sel.mode==="term" && !(student && isF5(student.form)) && (
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

      <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:8}}>📄 Whole class (one PDF)</div>
        <div style={{display:"flex",gap:8}}>
          <select style={{...inp,flex:1}} value={batchForm} onChange={e=>setBatchForm(e.target.value)}>{FORMS.map(f=><option key={f}>{f}</option>)}</select>
          <Btn onClick={printClass} disabled={!!batchProg}>{batchProg?`${batchProg.done}/${batchProg.total}…`:"Make PDF"}</Btn>
        </div>
        <div style={{fontSize:10.5,color:C.gray,marginTop:6}}>Uses the year / report type / term above. Downloads one PDF. Keep this screen open until it finishes.</div>
        <div style={{borderTop:`1px solid ${C.grayLight}`,marginTop:11,paddingTop:10}}>
          <label style={lbl}>Next term begins (printed on every card)</label>
          <div style={{display:"flex",gap:8}}>
            <input type="date" style={{...inp,flex:1}} value={nextTerm} onChange={e=>setNextTerm(e.target.value)}/>
            <Btn onClick={saveNextTerm} disabled={savingNT}>{savingNT?"Saving…":"Save"}</Btn>
          </div>
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

          {modeRows.length>0 && <ReportSheet html={buildPrintHTML()}/>}

          {modeRows.length>0
            ? <button onClick={printReport} style={{width:"100%",padding:"13px",background:`linear-gradient(90deg,${C.navy},${C.navyMid})`,color:C.white,border:"none",borderRadius:10,fontWeight:800,fontSize:15,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
                🖨 &nbsp;Print A4 Report Card — {student.name}
              </button>
            : <div style={{background:"#fef3cd",borderRadius:10,padding:"12px 16px",textAlign:"center",fontSize:13,color:"#92400e",border:"1px solid #fde68a"}}>⚠️ No marks found for this period. Enter marks first.</div>
          }
          {modeRows.length>0 && (
            <button onClick={shareWhatsApp} disabled={sharing} style={{width:"100%",marginTop:9,padding:"12px",background:"#128C7E",color:C.white,border:"none",borderRadius:10,fontWeight:800,fontSize:14,cursor:sharing?"wait":"pointer",opacity:sharing?0.7:1}}>
              {sharing ? "Preparing PDF…" : "💬  Save PDF & open WhatsApp"}
            </button>
          )}
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

// ─────────────────────────────────────────────────────────────────────────────
//  ACADEMIC ANALYTICS — shared by the Performance dashboard and Class Council
//  Scores are out of 20. Averages are weighted by subject coefficient, exactly
//  like the report cards. A student with no marks in the chosen term is left out.
// ─────────────────────────────────────────────────────────────────────────────
const PASS_MARK = 10; // pass mark out of 20 (same as report cards)
const AT_RISK_FAILS = 3; // a student is "at risk" if average < pass mark OR fails this many subjects
const fmt1 = v => (v===null||v===undefined||isNaN(v)) ? "—" : Number(v).toFixed(1);
const fmt2 = v => (v===null||v===undefined||isNaN(v)) ? "—" : Number(v).toFixed(2);
const escHtml = s => String(s ?? "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
const prevTermOf = term => TERMS[TERMS.indexOf(term)-1] || null;

const DECISION_STATUS = {
  pending:     { label:"Pending",     color:C.red  },
  in_progress: { label:"In Progress", color:C.gold },
  completed:   { label:"Completed",   color:C.green },
};
const isOverdue = d => d.deadline && d.status!=="completed" && d.deadline < todayStr();

// One student's result for one term: weighted average + per-subject averages
function studentTermResult(student, marksMap, year, term) {
  const subs = SUBJECTS_BY_FORM[student.form] || []; // (Form 5 rankings here still use the mocks only)
  const seqs = isF5(student.form) ? F5_MOCKS : (TERM_SEQS[term] || []);
  const rows = [];
  let tw = 0, tc = 0;
  subs.forEach(sub => {
    const vals = seqs
      .map(sq => marksMap[`${student.id}-${sub}-${sq}-${year}`]?.score)
      .filter(v => v !== null && v !== undefined)
      .map(Number);
    if (!vals.length) return;
    const avg = vals.reduce((a,b)=>a+b,0) / vals.length;
    const coeff = getCoeff(sub);
    rows.push({ sub, avg, coeff });
    tw += avg*coeff; tc += coeff;
  });
  const sorted = [...rows].sort((a,b) => b.avg - a.avg);
  const n = rows.length >= 2 ? Math.min(3, Math.floor(rows.length/2)) : rows.length;
  return {
    student,
    avg: tc ? tw/tc : null,
    rows,
    failed: rows.filter(r => r.avg < PASS_MARK).length,
    strengths:  sorted.slice(0, n),
    weaknesses: rows.length >= 2 ? sorted.slice(-n).reverse() : [],
  };
}

// Whole-class (or several classes) statistics for one term
function classTermStats(students, marksMap, forms, year, term) {
  const pool = students.filter(s => s.active && s.reg_status==="registered" && forms.includes(s.form));
  const all = pool.map(s => studentTermResult(s, marksMap, year, term));
  const assessed = all.filter(r => r.avg !== null);
  const byForm = {};
  assessed.forEach(r => { (byForm[r.student.form] = byForm[r.student.form] || []).push(r); });
  Object.values(byForm).forEach(list => {
    list.sort((a,b) => b.avg - a.avg);
    list.forEach((r,i) => { r.pos = i+1; r.of = list.length; });
  });
  assessed.sort((a,b) => b.avg - a.avg);

  const n = assessed.length;
  const classAvg = n ? assessed.reduce((a,r)=>a+r.avg,0)/n : null;
  const passRate = n ? assessed.filter(r=>r.avg>=PASS_MARK).length / n * 100 : null;

  const atRisk = assessed
    .filter(r => r.avg < PASS_MARK || r.failed >= AT_RISK_FAILS)
    .map(r => ({ ...r, reason: r.avg < PASS_MARK ? `Average below ${PASS_MARK}` : `Failing ${r.failed} subjects` }))
    .sort((a,b) => a.avg - b.avg);

  const subMap = {};
  assessed.forEach(r => r.rows.forEach(x => {
    const m = subMap[x.sub] = subMap[x.sub] || { sub:x.sub, total:0, count:0, pass:0 };
    m.total += x.avg; m.count++; if (x.avg >= PASS_MARK) m.pass++;
  }));
  const subjects = Object.values(subMap)
    .map(m => ({ sub:m.sub, avg:m.total/m.count, passRate:m.pass/m.count*100, count:m.count }))
    .sort((a,b) => b.avg - a.avg);

  const formStats = forms.map(f => {
    const l = byForm[f] || [];
    return { form:f, n:l.length,
      avg: l.length ? l.reduce((a,r)=>a+r.avg,0)/l.length : null,
      passRate: l.length ? l.filter(r=>r.avg>=PASS_MARK).length/l.length*100 : null };
  });

  return { assessed, n, noMarks: all.length - n, classAvg, passRate, atRisk, subjects, formStats, byForm };
}

// Top 3 / bottom 3 of one class (never overlap in very small classes)
function councilRanking(stats, form) {
  const list = stats.byForm[form] || [];
  const top = list.slice(0, Math.min(3, list.length));
  const bottom = list.slice(top.length).slice(-3).reverse();
  return { list, top, bottom };
}

// Frozen copy of the figures, stored with the meeting so the archive never changes
function buildCouncilSnapshot(stats, form) {
  const { top, bottom } = councilRanking(stats, form);
  const pack = r => ({
    name:r.student.name, id:r.student.id, avg:r.avg, pos:r.pos, of:r.of,
    strengths:r.strengths.map(x=>({sub:x.sub,avg:x.avg})),
    weaknesses:r.weaknesses.map(x=>({sub:x.sub,avg:x.avg})),
  });
  const list = stats.byForm[form] || [];
  return {
    classAvg: list.length ? list.reduce((a,r)=>a+r.avg,0)/list.length : null,
    passRate: list.length ? list.filter(r=>r.avg>=PASS_MARK).length/list.length*100 : null,
    assessed: list.length,
    atRisk: list.filter(r => r.avg < PASS_MARK || r.failed >= AT_RISK_FAILS).length,
    top: top.map(pack), bottom: bottom.map(pack),
    savedOn: todayStr(),
  };
}

// Small horizontal bar used on both pages
const ScoreBar = ({ label, value, max=20, color, note, labelWidth=108 }) => (
  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
    <div style={{width:labelWidth,fontSize:11,color:C.navy,flexShrink:0,lineHeight:1.2}}>{label}</div>
    <div style={{flex:1,background:C.grayBg,borderRadius:4,height:16,overflow:"hidden"}}>
      <div style={{width:`${Math.max(0,Math.min(100,(value/max)*100))}%`,background:color,height:"100%",borderRadius:4,transition:"width .4s"}}/>
    </div>
    <div style={{width:58,fontSize:11,fontWeight:700,color,textAlign:"right",flexShrink:0}}>{note ?? fmt1(value)}</div>
  </div>
);
const avgColor = v => v>=PASS_MARK+4 ? C.green : v>=PASS_MARK ? C.gold : C.red;

// ─── Academic Performance Dashboard ────────────────────────────────────────────
function PerformancePage({ ctx }) {
  const { students, marksMap, teachers, auth } = ctx;
  const isTeacher  = auth.role==="teacher";
  const teacher    = findTeacher(teachers, auth.user.email);
  const availForms = isTeacher ? (teacher?.forms||[]) : FORMS;

  const [f, setF] = useState({ cls: availForms.length===1 ? availForms[0] : "all", year:ACTIVE_YEAR, term:"First Term" });
  const [showAllRisk, setShowAllRisk] = useState(false);
  useEffect(() => { if (f.cls!=="all" && !availForms.includes(f.cls)) setF(x => ({ ...x, cls: availForms.length===1 ? availForms[0] : "all" })); }, [availForms.join("|")]);

  const forms = f.cls==="all" ? availForms : [f.cls];
  const stats = useMemo(() => classTermStats(students, marksMap, forms, f.year, f.term), [students, marksMap, f.cls, f.year, f.term, availForms.join("|")]);
  const termData = useMemo(() => TERMS.map(t => ({ term:t, ...classTermStats(students, marksMap, forms, f.year, t) })), [students, marksMap, f.cls, f.year, availForms.join("|")]);

  if (!availForms.length) return <TeacherSetupNotice teacher={teacher} email={auth.user.email}/>;

  const best  = stats.subjects.slice(0,3);
  const weak  = [...stats.subjects].reverse().slice(0,3).filter(s => !best.find(b=>b.sub===s.sub));
  const prev  = prevTermOf(f.term);
  const prevStats = prev ? termData.find(t=>t.term===prev) : null;
  const moves = prevStats && prevStats.n
    ? stats.subjects.map(s => { const p = prevStats.subjects.find(x=>x.sub===s.sub); return p ? { sub:s.sub, delta:s.avg-p.avg } : null; })
        .filter(Boolean).sort((a,b)=>b.delta-a.delta)
    : [];
  const gains = moves.filter(m=>m.delta>0.05).slice(0,3);
  const drops = [...moves].reverse().filter(m=>m.delta<-0.05).slice(0,3);
  const risk  = showAllRisk ? stats.atRisk : stats.atRisk.slice(0,10);

  return (
    <div>
      <h2 style={{margin:"0 0 12px",fontSize:18,fontWeight:800,color:C.navy}}>📈 Academic Performance</h2>

      <Card style={{marginBottom:11}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:9}}>
          <div><label style={lbl}>Class</label>
            <select style={inp} value={f.cls} onChange={e=>setF(x=>({...x,cls:e.target.value}))}>
              {availForms.length>1 && <option value="all">All classes</option>}
              {availForms.map(x=><option key={x} value={x}>{x}</option>)}
            </select></div>
          <div><label style={lbl}>Academic Year</label>
            <select style={inp} value={f.year} onChange={e=>setF(x=>({...x,year:e.target.value}))}>
              {ACAD_YEARS.map(y=><option key={y}>{y}</option>)}
            </select></div>
        </div>
        <label style={lbl}>Term</label>
        <div style={{display:"flex",gap:6}}>
          {TERMS.map(t=>(
            <button key={t} onClick={()=>setF(x=>({...x,term:t}))} style={{flex:1,padding:"8px 2px",borderRadius:8,fontSize:12,fontWeight:700,cursor:"pointer",border:`1px solid ${f.term===t?C.navy:C.border}`,background:f.term===t?C.navy:C.white,color:f.term===t?C.white:C.gray}}>{t.replace(" Term","")}</button>
          ))}
        </div>
      </Card>

      {!stats.n ? <Card><Empty text={`No marks entered for ${f.cls==="all"?"these classes":f.cls} in ${f.term}, ${f.year}.`}/></Card> : (<>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:11}}>
          {[
            { label:"Class average",     value:`${fmt2(stats.classAvg)}/20`, sub:`${fmt1(stats.classAvg*5)}%`, color:avgColor(stats.classAvg) },
            { label:"Pass rate",         value:`${fmt1(stats.passRate)}%`,   sub:`average ≥ ${PASS_MARK}/20`, color:stats.passRate>=50?C.green:C.red },
            { label:"Students assessed", value:stats.n,                      sub:stats.noMarks?`${stats.noMarks} without marks`:"all have marks", color:C.navy },
            { label:"Students at risk",  value:stats.atRisk.length,          sub:`avg < ${PASS_MARK} or ${AT_RISK_FAILS}+ fails`, color:stats.atRisk.length?C.red:C.green },
          ].map(c=>(
            <div key={c.label} style={{background:C.white,borderRadius:11,padding:13,borderLeft:`4px solid ${c.color}`,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
              <div style={{fontSize:11,color:C.gray}}>{c.label}</div>
              <div style={{fontSize:22,fontWeight:900,color:c.color,margin:"2px 0"}}>{c.value}</div>
              <div style={{fontSize:10,color:C.gray}}>{c.sub}</div>
            </div>
          ))}
        </div>

        {f.cls==="all" && (
          <Card title="Class averages" style={{marginBottom:11}}>
            {stats.formStats.map(x => x.n
              ? <ScoreBar key={x.form} label={`${x.form} (${x.n})`} value={x.avg} color={avgColor(x.avg)} note={`${fmt1(x.avg)} · ${fmt1(x.passRate)}%`} labelWidth={80}/>
              : <div key={x.form} style={{fontSize:11,color:C.gray,marginBottom:6}}>{x.form}: no marks yet</div>)}
            <div style={{fontSize:10,color:C.gray,marginTop:4}}>Bar shows average out of 20 · then pass rate</div>
          </Card>
        )}

        <div style={{display:"grid",gridTemplateColumns:"1fr",gap:11,marginBottom:11}}>
          <Card title="🏆 Best-performing subjects">
            {best.map((s,i)=>(
              <div key={s.sub} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:i<best.length-1?`1px solid ${C.grayBg}`:"none"}}>
                <div><div style={{fontSize:12.5,fontWeight:700,color:C.navy}}>{i+1}. {s.sub}</div><div style={{fontSize:10.5,color:C.gray}}>{fmt1(s.passRate)}% pass · {s.count} students</div></div>
                <Pill color={C.green}>{fmt2(s.avg)}/20</Pill>
              </div>
            ))}
          </Card>
          <Card title="⚠️ Weakest subjects">
            {weak.length ? weak.map((s,i)=>(
              <div key={s.sub} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:i<weak.length-1?`1px solid ${C.grayBg}`:"none"}}>
                <div><div style={{fontSize:12.5,fontWeight:700,color:C.navy}}>{s.sub}</div><div style={{fontSize:10.5,color:C.gray}}>{fmt1(s.passRate)}% pass · {s.count} students</div></div>
                <Pill color={avgColor(s.avg)}>{fmt2(s.avg)}/20</Pill>
              </div>
            )) : <div style={{fontSize:12,color:C.gray}}>Not enough subjects with marks yet.</div>}
          </Card>
        </div>

        <Card title="All subject averages" style={{marginBottom:11}}>
          {stats.subjects.map(s => <ScoreBar key={s.sub} label={s.sub} value={s.avg} color={avgColor(s.avg)} labelWidth={118}/>)}
          <div style={{fontSize:10,color:C.gray,marginTop:4}}>Out of 20 · green = strong, gold = pass, red = below {PASS_MARK}</div>
        </Card>

        <Card title={`Students at risk (${stats.atRisk.length})`} style={{marginBottom:11}}>
          {!stats.atRisk.length && <div style={{fontSize:12,color:C.green}}>✓ No student is currently at risk.</div>}
          {risk.map(r=>(
            <div key={r.student.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,padding:"7px 0",borderBottom:`1px solid ${C.grayBg}`}}>
              <div style={{minWidth:0}}>
                <div style={{fontSize:12.5,fontWeight:700,color:C.navy}}>{r.student.name}</div>
                <div style={{fontSize:10.5,color:C.gray}}>{r.student.form} · {r.student.id} · {r.reason}</div>
                {r.weaknesses.length>0 && <div style={{fontSize:10.5,color:C.red,marginTop:1}}>Weak: {r.weaknesses.map(w=>w.sub).join(", ")}</div>}
              </div>
              <div style={{textAlign:"right",flexShrink:0}}>
                <div style={{fontSize:14,fontWeight:900,color:C.red}}>{fmt2(r.avg)}</div>
                <div style={{fontSize:10,color:C.gray}}>{r.failed} failed</div>
              </div>
            </div>
          ))}
          {stats.atRisk.length>10 && <div style={{marginTop:8}}><SmBtn color={C.navyMid} onClick={()=>setShowAllRisk(v=>!v)}>{showAllRisk?"Show fewer":`Show all ${stats.atRisk.length}`}</SmBtn></div>}
        </Card>

        <Card title="Term comparison">
          {termData.map((t,i)=>{
            const p = i>0 ? termData[i-1] : null;
            const dAvg = p && p.n && t.n ? t.classAvg - p.classAvg : null;
            return t.n ? (
              <div key={t.term} style={{padding:"7px 0",borderBottom:i<2?`1px solid ${C.grayBg}`:"none",background:t.term===f.term?"#f8fafc":"transparent"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                  <div style={{fontSize:12.5,fontWeight:700,color:C.navy}}>{t.term}{t.term===f.term?" (selected)":""}</div>
                  <div style={{fontSize:10.5,color:C.gray}}>{t.n} students</div>
                </div>
                <ScoreBar label="Average /20" value={t.classAvg} color={avgColor(t.classAvg)} note={fmt2(t.classAvg)+(dAvg!==null?` ${dAvg>=0?"▲":"▼"}${Math.abs(dAvg).toFixed(2)}`:"")} labelWidth={78}/>
                <ScoreBar label="Pass rate" value={t.passRate} max={100} color={t.passRate>=50?C.green:C.red} note={`${fmt1(t.passRate)}%`} labelWidth={78}/>
              </div>
            ) : (
              <div key={t.term} style={{padding:"7px 0",borderBottom:i<2?`1px solid ${C.grayBg}`:"none",fontSize:12,color:C.gray}}>{t.term}: no marks yet</div>
            );
          })}
          {prev && (gains.length>0 || drops.length>0) && (
            <div style={{marginTop:10,paddingTop:9,borderTop:`1px dashed ${C.border}`}}>
              <div style={{fontSize:11,fontWeight:700,color:C.navy,marginBottom:5}}>Subject movement vs {prev}</div>
              {gains.map(m=><div key={m.sub} style={{fontSize:11.5,color:C.green,marginBottom:2}}>▲ {m.sub} +{m.delta.toFixed(2)}</div>)}
              {drops.map(m=><div key={m.sub} style={{fontSize:11.5,color:C.red,marginBottom:2}}>▼ {m.sub} {m.delta.toFixed(2)}</div>)}
            </div>
          )}
        </Card>
      </>)}
    </div>
  );
}

// ─── Class Council ─────────────────────────────────────────────────────────────
const RESPONSIBLE_SUGGESTIONS = ["Principal","Vice Principal","Class Master","Discipline Master","Guidance Counsellor","Bursar"];

// Printable council report for the school archive
function printCouncilReport(meeting, decisions, snap, teacherNames=[]) {
  const stuRow = r => `<tr>
    <td style="text-align:center;font-weight:700">${r.pos}/${r.of}</td>
    <td><strong>${escHtml(r.name)}</strong><br><span style="color:#6B7280;font-size:9px">${escHtml(r.id)}</span></td>
    <td style="text-align:center;font-weight:700">${fmt2(r.avg)}</td>
    <td style="color:#15803D">${r.strengths.length ? r.strengths.map(x=>`${escHtml(x.sub)} (${fmt1(x.avg)})`).join(", ") : "—"}</td>
    <td style="color:#B91C1C">${r.weaknesses.length ? r.weaknesses.map(x=>`${escHtml(x.sub)} (${fmt1(x.avg)})`).join(", ") : "—"}</td>
  </tr>`;
  const stuTable = (title, list, color) => `
    <div style="margin:12px 0 4px;font-weight:800;font-size:12px;color:${color}">${title}</div>
    <table class="grid"><thead><tr><th style="width:44px">Position</th><th>Student</th><th style="width:52px">Average /20</th><th>Strengths</th><th>Weaknesses</th></tr></thead>
    <tbody>${list.length ? list.map(stuRow).join("") : `<tr><td colspan="5" style="text-align:center;color:#6B7280">No ranked students</td></tr>`}</tbody></table>`;

  const statusTxt = s => (DECISION_STATUS[s]||DECISION_STATUS.pending).label;
  const statusCol = s => (DECISION_STATUS[s]||DECISION_STATUS.pending).color;
  const decRows = decisions.map((d,i)=>`<tr>
    <td style="text-align:center">${i+1}</td>
    <td>${escHtml(d.decision)}</td>
    <td>${escHtml(d.responsible||"—")}</td>
    <td style="text-align:center">${d.deadline?fmtDate(d.deadline):"—"}</td>
    <td style="text-align:center;font-weight:700;color:${statusCol(d.status)}">${statusTxt(d.status)}${d.status==="completed"&&d.completed_on?`<br><span style="font-weight:400;font-size:9px">${fmtDate(d.completed_on)}</span>`:""}</td>
  </tr>`).join("");

  const html = `<style>
    .rep{font-family:'Segoe UI',Arial,sans-serif;color:#0D2340;font-size:11px;line-height:1.4}
    .rep table.grid{width:100%;border-collapse:collapse}
    .rep table.grid th{background:#0D2340;color:#fff;font-size:9.5px;padding:4px 5px;text-align:left;border:1px solid #0D2340}
    .rep table.grid td{border:1px solid #cbd5e1;padding:4px 5px;vertical-align:top;font-size:10px}
    .rep h3{font-size:12px;margin:14px 0 5px;padding-bottom:3px;border-bottom:2px solid #C9962A}
    .box{display:inline-block;border:1px solid #cbd5e1;border-radius:6px;padding:5px 12px;margin:0 6px 6px 0;text-align:center}
  </style>
  <div class="rep">
    <div style="text-align:center;border-bottom:3px double #0D2340;padding-bottom:8px">
      <img src="${SCHOOL_LOGO}" alt="SBCC" style="width:70px;height:70px;object-fit:contain">
      <div style="font-size:16px;font-weight:900;margin-top:3px">SAKER BAPTIST COMPREHENSIVE COLLEGE BAWE</div>
      <div style="font-size:9px;color:#6B7280">NGEPTANG · NONI · NW REGION</div>
      <div style="display:inline-block;margin-top:7px;background:#0D2340;color:#fff;padding:3px 18px;border-radius:12px;font-weight:800;font-size:11px;letter-spacing:.5px">CLASS COUNCIL REPORT</div>
    </div>

    <table style="width:100%;margin-top:9px;font-size:11px"><tr>
      <td><strong>Class:</strong> ${escHtml(meeting.form)}</td>
      <td><strong>Term:</strong> ${escHtml(meeting.term)}</td>
      <td><strong>Academic Year:</strong> ${escHtml(meeting.acad_year)}</td>
      <td style="text-align:right"><strong>Meeting date:</strong> ${fmtDate(meeting.meeting_date)}</td>
    </tr></table>

    <h3>1. Class performance summary</h3>
    <div>
      <div class="box"><div style="font-size:9px;color:#6B7280">Class average</div><div style="font-size:15px;font-weight:900">${fmt2(snap.classAvg)}/20</div></div>
      <div class="box"><div style="font-size:9px;color:#6B7280">Pass rate</div><div style="font-size:15px;font-weight:900">${fmt1(snap.passRate)}%</div></div>
      <div class="box"><div style="font-size:9px;color:#6B7280">Students ranked</div><div style="font-size:15px;font-weight:900">${snap.assessed}</div></div>
      <div class="box"><div style="font-size:9px;color:#6B7280">Students at risk</div><div style="font-size:15px;font-weight:900;color:#B91C1C">${snap.atRisk}</div></div>
    </div>

    <h3>2. Top and bottom students</h3>
    ${stuTable("🏆 Top three students", snap.top, "#15803D")}
    ${stuTable("⚠ Bottom three students", snap.bottom, "#B91C1C")}

    <h3>3. Meeting record</h3>
    <div style="margin-bottom:5px"><strong>Attendees (${(meeting.attendees||[]).length}):</strong> ${(meeting.attendees||[]).length ? (meeting.attendees||[]).map(escHtml).join(", ") : "—"}</div>
    <div><strong>Discussion:</strong></div>
    <div style="border:1px solid #cbd5e1;border-radius:5px;padding:7px;min-height:50px;white-space:pre-wrap">${escHtml(meeting.discussion||"—")}</div>

    <h3>4. Decisions and follow-up</h3>
    ${decisions.length ? `<table class="grid"><thead><tr><th style="width:24px">#</th><th>Decision</th><th style="width:100px">Responsible</th><th style="width:62px">Deadline</th><th style="width:70px">Status</th></tr></thead><tbody>${decRows}</tbody></table>` : `<div style="color:#6B7280">No decisions recorded.</div>`}

    <table style="width:100%;margin-top:34px;font-size:10px"><tr>
      <td style="width:50%;text-align:center"><div style="border-top:1px solid #0D2340;width:80%;margin:0 auto;padding-top:3px">Class Master / Council Secretary</div></td>
      <td style="width:50%;text-align:center"><div style="border-top:1px solid #0D2340;width:80%;margin:0 auto;padding-top:3px">Principal</div></td>
    </tr></table>
    <div style="margin-top:12px;font-size:8.5px;color:#9ca3af;text-align:center">
      Ranking figures recorded on ${fmtDate(snap.savedOn||todayStr())} · Decision statuses printed on ${fmtDate(todayStr())} · Archive copy
    </div>
  </div>`;
  domPrint("sbc-council-report", html, "A4 portrait", "10mm");
}

const blankDecision = () => ({ key:Math.random().toString(36).slice(2), id:null, decision:"", responsible:"", deadline:"", status:"pending", completed_on:null });

function ClassCouncilPage({ ctx }) {
  const { students, marksMap, teachers, auth, councilMeetings, councilDecisions, councilReady,
          saveCouncilMeeting, deleteCouncilMeeting, updateDecisionStatus } = ctx;
  const isTeacher  = auth.role==="teacher";
  const teacher    = findTeacher(teachers, auth.user.email);
  const availForms = isTeacher ? (teacher?.forms||[]) : FORMS;

  const [sel, setSel]       = useState({ form:availForms[0]||"Form 1", year:ACTIVE_YEAR, term:"First Term" });
  useEffect(() => { if (availForms.length && !availForms.includes(sel.form)) setSel(x => ({ ...x, form:availForms[0] })); }, [availForms.join("|")]);
  const [tab, setTab]       = useState("rank"); // rank | meetings | actions
  const [showFull, setShowFull] = useState(false);
  const [openId, setOpenId] = useState(null);
  const [modal, setModal]   = useState(null);   // null | meeting draft
  const [saving, setSaving] = useState(false);
  const [err, setErr]       = useState("");
  const [actFilter, setActFilter] = useState("open"); // open | all | completed

  const stats   = useMemo(() => classTermStats(students, marksMap, [sel.form], sel.year, sel.term), [students, marksMap, sel.form, sel.year, sel.term]);
  const ranking = councilRanking(stats, sel.form);
  const activeTeachers = teachers.filter(t=>t.active).map(t=>t.name).filter(Boolean);

  const meetingsFor = councilMeetings.filter(m => m.form===sel.form && m.acad_year===sel.year && m.term===sel.term)
                                      .sort((a,b)=> (b.meeting_date||"").localeCompare(a.meeting_date||""));
  const decisionsOf = id => councilDecisions.filter(d => d.meeting_id===id);
  const classMeetings = councilMeetings.filter(m => m.form===sel.form && m.acad_year===sel.year);
  const classDecisions = councilDecisions
    .filter(d => classMeetings.some(m=>m.id===d.meeting_id))
    .map(d => ({ ...d, meeting: classMeetings.find(m=>m.id===d.meeting_id) }))
    .sort((a,b)=> (a.deadline||"9999").localeCompare(b.deadline||"9999"));
  const shownDecisions = classDecisions.filter(d => actFilter==="all" ? true : actFilter==="completed" ? d.status==="completed" : d.status!=="completed");
  const counts = { pending:classDecisions.filter(d=>d.status==="pending").length,
                   in_progress:classDecisions.filter(d=>d.status==="in_progress").length,
                   completed:classDecisions.filter(d=>d.status==="completed").length,
                   overdue:classDecisions.filter(isOverdue).length };

  if (!availForms.length) return <TeacherSetupNotice teacher={teacher} email={auth.user.email}/>;

  function openNew() {
    setErr("");
    setModal({ id:null, meeting_date:todayStr(), attendees:[], other:"", discussion:"", decisions:[blankDecision()], snapshot:null });
  }
  function openEdit(m) {
    setErr("");
    const att = m.attendees||[];
    setModal({
      id:m.id, meeting_date:m.meeting_date, discussion:m.discussion||"", snapshot:m.snapshot||null,
      attendees: att.filter(a=>activeTeachers.includes(a)),
      other: att.filter(a=>!activeTeachers.includes(a)).join(", "),
      decisions: decisionsOf(m.id).map(d=>({ key:String(d.id), id:d.id, decision:d.decision, responsible:d.responsible||"", deadline:d.deadline||"", status:d.status, completed_on:d.completed_on||null })),
    });
  }
  const setM = patch => setModal(m => ({ ...m, ...patch }));
  const setD = (key, patch) => setModal(m => ({ ...m, decisions:m.decisions.map(d=>d.key===key?{...d,...patch}:d) }));
  const toggleAtt = name => setModal(m => ({ ...m, attendees: m.attendees.includes(name) ? m.attendees.filter(a=>a!==name) : [...m.attendees, name] }));

  async function submit() {
    if (!modal.meeting_date) { setErr("Please choose the meeting date."); return; }
    const decisions = modal.decisions.filter(d=>d.decision.trim());
    if (!modal.discussion.trim() && !decisions.length) { setErr("Add the discussion or at least one decision."); return; }
    const attendees = [...modal.attendees, ...modal.other.split(",").map(s=>s.trim()).filter(Boolean)];
    setSaving(true); setErr("");
    try {
      await saveCouncilMeeting({
        id:modal.id, form:sel.form, acad_year:sel.year, term:sel.term,
        meeting_date:modal.meeting_date, attendees, discussion:modal.discussion.trim(),
        snapshot: modal.snapshot || (stats.byForm[sel.form] ? buildCouncilSnapshot(stats, sel.form) : null),
      }, decisions.map(d=>({ ...d, decision:d.decision.trim() })));
      setModal(null);
    } catch(e) { setErr(e.message || "Could not save. Please try again."); }
    setSaving(false);
  }

  async function setStatus(d, status) {
    try { await updateDecisionStatus(d.id, status); } catch(e) { alert(e.message || "Could not update status."); }
  }
  async function removeMeeting(m) {
    if (!window.confirm(`Delete the council meeting of ${fmtDate(m.meeting_date)} and all its decisions?`)) return;
    try { await deleteCouncilMeeting(m.id); } catch(e) { alert(e.message || "Could not delete."); }
  }
  function printMeeting(m) {
    const snap = m.snapshot || buildCouncilSnapshot(stats, sel.form);
    printCouncilReport(m, decisionsOf(m.id), snap);
  }

  const StatusButtons = ({ d }) => (
    <div style={{display:"flex",gap:5,flexWrap:"wrap",marginTop:5}}>
      {Object.entries(DECISION_STATUS).map(([k,v])=>(
        <button key={k} onClick={()=>d.status!==k && setStatus(d,k)} style={{padding:"4px 9px",borderRadius:12,fontSize:11,fontWeight:700,cursor:"pointer",border:`1px solid ${v.color}`,background:d.status===k?v.color:C.white,color:d.status===k?C.white:v.color}}>{v.label}</button>
      ))}
    </div>
  );

  const StudentCard = ({ r, accent }) => (
    <div style={{border:`1px solid ${C.grayLight}`,borderLeft:`4px solid ${accent}`,borderRadius:9,padding:"9px 11px",marginBottom:8}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:9,minWidth:0}}>
          <div style={{width:34,height:34,borderRadius:"50%",background:accent+"20",color:accent,fontWeight:900,fontSize:13,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{r.pos}</div>
          <div style={{minWidth:0}}>
            <div style={{fontSize:13,fontWeight:800,color:C.navy}}>{r.student.name}</div>
            <div style={{fontSize:10.5,color:C.gray}}>{r.student.id} · Position {r.pos}/{r.of}</div>
          </div>
        </div>
        <div style={{textAlign:"right",flexShrink:0}}>
          <div style={{fontSize:17,fontWeight:900,color:avgColor(r.avg)}}>{fmt2(r.avg)}</div>
          <div style={{fontSize:9.5,color:C.gray}}>average /20</div>
        </div>
      </div>
      <div style={{marginTop:7,fontSize:11}}>
        <div style={{color:C.green,marginBottom:2}}><strong>Strengths:</strong> {r.strengths.length ? r.strengths.map(x=>`${x.sub} (${fmt1(x.avg)})`).join(", ") : "—"}</div>
        <div style={{color:C.red}}><strong>Weaknesses:</strong> {r.weaknesses.length ? r.weaknesses.map(x=>`${x.sub} (${fmt1(x.avg)})`).join(", ") : "—"}</div>
      </div>
    </div>
  );

  const tabBtn = (k,label) => (
    <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"9px 2px",fontSize:12.5,fontWeight:800,cursor:"pointer",border:"none",borderBottom:`3px solid ${tab===k?C.gold:"transparent"}`,background:"transparent",color:tab===k?C.navy:C.gray}}>{label}</button>
  );

  return (
    <div>
      <h2 style={{margin:"0 0 12px",fontSize:18,fontWeight:800,color:C.navy}}>🏛️ Class Council</h2>

      {!councilReady && (
        <div style={{background:"#fffbeb",border:"1px solid #fcd34d",borderRadius:10,padding:12,marginBottom:11,fontSize:12,color:"#78350f"}}>
          <strong>One-time setup needed.</strong> The council tables are not in your database yet, so meetings cannot be saved. Run the <code>council_setup.sql</code> script in Supabase (SQL Editor), then reload this page. Rankings below still work.
        </div>
      )}

      <Card style={{marginBottom:11}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:9}}>
          <div><label style={lbl}>Class</label>
            <select style={inp} value={sel.form} onChange={e=>setSel(x=>({...x,form:e.target.value}))}>{availForms.map(x=><option key={x}>{x}</option>)}</select></div>
          <div><label style={lbl}>Academic Year</label>
            <select style={inp} value={sel.year} onChange={e=>setSel(x=>({...x,year:e.target.value}))}>{ACAD_YEARS.map(y=><option key={y}>{y}</option>)}</select></div>
        </div>
        <label style={lbl}>Term</label>
        <div style={{display:"flex",gap:6}}>
          {TERMS.map(t=>(
            <button key={t} onClick={()=>setSel(x=>({...x,term:t}))} style={{flex:1,padding:"8px 2px",borderRadius:8,fontSize:12,fontWeight:700,cursor:"pointer",border:`1px solid ${sel.term===t?C.navy:C.border}`,background:sel.term===t?C.navy:C.white,color:sel.term===t?C.white:C.gray}}>{t.replace(" Term","")}</button>
          ))}
        </div>
      </Card>

      <div style={{display:"flex",background:C.white,borderRadius:"11px 11px 0 0",borderBottom:`1px solid ${C.grayLight}`}}>
        {tabBtn("rank","Rankings")}
        {tabBtn("meetings",`Meetings (${meetingsFor.length})`)}
        {tabBtn("actions",`Actions (${counts.pending+counts.in_progress})`)}
      </div>
      <div style={{background:C.white,borderRadius:"0 0 11px 11px",padding:13,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>

        {/* ── Rankings ── */}
        {tab==="rank" && (!stats.n ? <Empty text={`No marks for ${sel.form} in ${sel.term}, ${sel.year}.`}/> : (<>
          <div style={{display:"flex",gap:8,marginBottom:12}}>
            {[["Class average",`${fmt2(stats.classAvg)}/20`,avgColor(stats.classAvg)],["Pass rate",`${fmt1(stats.passRate)}%`,stats.passRate>=50?C.green:C.red],["Ranked",stats.n,C.navy]].map(([l,v,c])=>(
              <div key={l} style={{flex:1,background:C.grayBg,borderRadius:9,padding:"8px 6px",textAlign:"center"}}>
                <div style={{fontSize:9.5,color:C.gray}}>{l}</div><div style={{fontSize:15,fontWeight:900,color:c}}>{v}</div>
              </div>
            ))}
          </div>
          <div style={{fontSize:12.5,fontWeight:800,color:C.green,margin:"4px 0 7px"}}>🏆 Top three</div>
          {ranking.top.map(r=><StudentCard key={r.student.id} r={r} accent={C.green}/>)}
          <div style={{fontSize:12.5,fontWeight:800,color:C.red,margin:"12px 0 7px"}}>⚠️ Bottom three</div>
          {ranking.bottom.length ? ranking.bottom.map(r=><StudentCard key={r.student.id} r={r} accent={C.red}/>) : <div style={{fontSize:12,color:C.gray}}>Class too small for a separate bottom three.</div>}
          {stats.noMarks>0 && <div style={{fontSize:10.5,color:C.gray,marginTop:6}}>{stats.noMarks} student(s) with no marks this term are not ranked.</div>}
          <div style={{marginTop:10}}><SmBtn color={C.navyMid} onClick={()=>setShowFull(v=>!v)}>{showFull?"Hide full class list":"Show full class list"}</SmBtn></div>
          {showFull && (
            <div style={{marginTop:9}}>
              {ranking.list.map(r=>(
                <div key={r.student.id} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:`1px solid ${C.grayBg}`,fontSize:12}}>
                  <span style={{color:C.navy}}><strong style={{display:"inline-block",width:30}}>{r.pos}.</strong>{r.student.name}</span>
                  <span style={{fontWeight:700,color:avgColor(r.avg)}}>{fmt2(r.avg)}</span>
                </div>
              ))}
            </div>
          )}
        </>))}

        {/* ── Meetings ── */}
        {tab==="meetings" && (<>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <div style={{fontSize:12,color:C.gray}}>{sel.form} · {sel.term} · {sel.year}</div>
            <Btn onClick={openNew} disabled={!councilReady}>+ New meeting</Btn>
          </div>
          {!meetingsFor.length && <Empty text="No council meeting recorded for this class and term yet."/>}
          {meetingsFor.map(m=>{
            const ds = decisionsOf(m.id);
            const done = ds.filter(d=>d.status==="completed").length;
            const open = openId===m.id;
            return (
              <div key={m.id} style={{border:`1px solid ${C.grayLight}`,borderRadius:10,marginBottom:9,overflow:"hidden"}}>
                <div onClick={()=>setOpenId(open?null:m.id)} style={{padding:"10px 12px",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,background:open?"#f8fafc":C.white}}>
                  <div>
                    <div style={{fontSize:13,fontWeight:800,color:C.navy}}>{fmtDate(m.meeting_date)}</div>
                    <div style={{fontSize:10.5,color:C.gray}}>{(m.attendees||[]).length} attendees · {ds.length} decisions{ds.length?` (${done} done)`:""}</div>
                  </div>
                  <span style={{color:C.gray,fontSize:12}}>{open?"▲":"▼"}</span>
                </div>
                {open && (
                  <div style={{padding:"4px 12px 12px"}}>
                    <div style={{fontSize:11,color:C.gray,marginTop:6}}>Attendees</div>
                    <div style={{fontSize:12,color:C.navy,marginBottom:7}}>{(m.attendees||[]).join(", ")||"—"}</div>
                    <div style={{fontSize:11,color:C.gray}}>Discussion</div>
                    <div style={{fontSize:12,color:C.navy,whiteSpace:"pre-wrap",marginBottom:8}}>{m.discussion||"—"}</div>
                    <div style={{fontSize:11,color:C.gray,marginBottom:4}}>Decisions</div>
                    {!ds.length && <div style={{fontSize:12,color:C.gray}}>None recorded.</div>}
                    {ds.map(d=>(
                      <div key={d.id} style={{background:C.grayBg,borderRadius:8,padding:"8px 10px",marginBottom:6}}>
                        <div style={{fontSize:12.5,fontWeight:600,color:C.navy}}>{d.decision}</div>
                        <div style={{fontSize:10.5,color:isOverdue(d)?C.red:C.gray,marginTop:2}}>👤 {d.responsible||"Unassigned"} · ⏰ {d.deadline?fmtDate(d.deadline):"No deadline"}{isOverdue(d)?" · OVERDUE":""}</div>
                        <StatusButtons d={d}/>
                      </div>
                    ))}
                    <div style={{display:"flex",gap:7,marginTop:8,flexWrap:"wrap"}}>
                      <SmBtn color={C.navyMid} onClick={()=>openEdit(m)}>✏️ Edit</SmBtn>
                      <SmBtn color={C.green} onClick={()=>printMeeting(m)}>🖨️ Print report</SmBtn>
                      {!isTeacher && <SmBtn color={C.red} onClick={()=>removeMeeting(m)}>🗑 Delete</SmBtn>}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </>)}

        {/* ── Actions tracker ── */}
        {tab==="actions" && (<>
          <div style={{display:"flex",gap:7,marginBottom:10}}>
            {[["Pending",counts.pending,C.red],["In progress",counts.in_progress,C.gold],["Done",counts.completed,C.green]].map(([l,v,c])=>(
              <div key={l} style={{flex:1,background:c+"15",borderRadius:9,padding:"7px 4px",textAlign:"center"}}>
                <div style={{fontSize:17,fontWeight:900,color:c}}>{v}</div><div style={{fontSize:9.5,color:C.gray}}>{l}</div>
              </div>
            ))}
          </div>
          {counts.overdue>0 && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"6px 10px",fontSize:11.5,color:C.red,marginBottom:9}}>⏰ {counts.overdue} decision(s) past deadline</div>}
          <div style={{display:"flex",gap:6,marginBottom:10}}>
            {[["open","Open"],["completed","Completed"],["all","All"]].map(([k,l])=>(
              <button key={k} onClick={()=>setActFilter(k)} style={{padding:"5px 12px",borderRadius:14,fontSize:11.5,fontWeight:700,cursor:"pointer",border:`1px solid ${actFilter===k?C.navy:C.border}`,background:actFilter===k?C.navy:C.white,color:actFilter===k?C.white:C.gray}}>{l}</button>
            ))}
          </div>
          <div style={{fontSize:10.5,color:C.gray,marginBottom:8}}>All terms of {sel.year} for {sel.form}</div>
          {!shownDecisions.length && <Empty text="No decisions to show."/>}
          {shownDecisions.map(d=>(
            <div key={d.id} style={{border:`1px solid ${isOverdue(d)?"#fca5a5":C.grayLight}`,borderLeft:`4px solid ${DECISION_STATUS[d.status]?.color||C.gray}`,borderRadius:9,padding:"9px 11px",marginBottom:8}}>
              <div style={{fontSize:12.5,fontWeight:600,color:C.navy}}>{d.decision}</div>
              <div style={{fontSize:10.5,color:isOverdue(d)?C.red:C.gray,marginTop:2}}>👤 {d.responsible||"Unassigned"} · ⏰ {d.deadline?fmtDate(d.deadline):"No deadline"}{isOverdue(d)?" · OVERDUE":""}</div>
              <div style={{fontSize:10,color:C.gray}}>From {d.meeting.term} meeting, {fmtDate(d.meeting.meeting_date)}{d.status==="completed"&&d.completed_on?` · done ${fmtDate(d.completed_on)}`:""}</div>
              <StatusButtons d={d}/>
            </div>
          ))}
        </>)}
      </div>

      {/* ── Meeting form ── */}
      {modal && (
        <Modal title={modal.id ? "Edit council meeting" : "New council meeting"} onClose={()=>!saving&&setModal(null)}>
          <div style={{fontSize:11.5,color:C.gray,marginBottom:10}}>{sel.form} · {sel.term} · {sel.year}</div>
          <Fr label="Meeting date"><input style={inp} type="date" value={modal.meeting_date} onChange={e=>setM({meeting_date:e.target.value})}/></Fr>

          <Fr label="Attendees (tap to select)">
            <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:7}}>
              {activeTeachers.map(n=>{ const on=modal.attendees.includes(n); return (
                <button key={n} type="button" onClick={()=>toggleAtt(n)} style={{padding:"5px 10px",borderRadius:14,fontSize:11.5,fontWeight:600,cursor:"pointer",border:`1px solid ${on?C.navy:C.border}`,background:on?C.navy:C.white,color:on?C.white:C.gray}}>{on?"✓ ":""}{n}</button>
              );})}
              {!activeTeachers.length && <span style={{fontSize:11.5,color:C.gray}}>No teachers on file — type names below.</span>}
            </div>
            <input style={inp} placeholder="Others (Principal, parent delegate…) separated by commas" value={modal.other} onChange={e=>setM({other:e.target.value})}/>
          </Fr>

          <Fr label="Discussion">
            <textarea style={{...inp,minHeight:90,resize:"vertical"}} placeholder="Summary of what was discussed…" value={modal.discussion} onChange={e=>setM({discussion:e.target.value})}/>
          </Fr>

          <div style={{fontSize:12,fontWeight:800,color:C.navy,margin:"12px 0 6px"}}>Decisions</div>
          <datalist id="council-people">{[...RESPONSIBLE_SUGGESTIONS,...activeTeachers].map(n=><option key={n} value={n}/>)}</datalist>
          {modal.decisions.map((d,i)=>(
            <div key={d.key} style={{border:`1px solid ${C.grayLight}`,borderRadius:9,padding:10,marginBottom:9,background:"#fafafa"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5}}>
                <span style={{fontSize:11.5,fontWeight:800,color:C.navyMid}}>Decision {i+1}</span>
                <button type="button" onClick={()=>setM({decisions:modal.decisions.filter(x=>x.key!==d.key)})} style={{background:"none",border:"none",color:C.red,fontSize:12,fontWeight:700,cursor:"pointer"}}>Remove</button>
              </div>
              <textarea style={{...inp,minHeight:56,resize:"vertical",marginBottom:7}} placeholder="What was decided?" value={d.decision} onChange={e=>setD(d.key,{decision:e.target.value})}/>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:7,marginBottom:7}}>
                <div><label style={lbl}>Responsible person</label><input style={inp} list="council-people" placeholder="Name or role" value={d.responsible} onChange={e=>setD(d.key,{responsible:e.target.value})}/></div>
                <div><label style={lbl}>Deadline</label><input style={inp} type="date" value={d.deadline} onChange={e=>setD(d.key,{deadline:e.target.value})}/></div>
              </div>
              <label style={lbl}>Status</label>
              <select style={inp} value={d.status} onChange={e=>setD(d.key,{status:e.target.value, completed_on:e.target.value==="completed"?(d.completed_on||todayStr()):null})}>
                {Object.entries(DECISION_STATUS).map(([k,v])=><option key={k} value={k}>{v.label}</option>)}
              </select>
            </div>
          ))}
          <SmBtn color={C.navyMid} onClick={()=>setM({decisions:[...modal.decisions, blankDecision()]})}>+ Add decision</SmBtn>

          {err && <div style={{background:"#fef2f2",border:"1px solid #fca5a5",borderRadius:8,padding:"8px 12px",color:C.red,fontSize:12,margin:"12px 0 0"}}>{err}</div>}
          <div style={{display:"flex",gap:8,marginTop:14}}>
            <Btn onClick={submit} disabled={saving}>{saving?"Saving…":"Save meeting"}</Btn>
            <Btn outline onClick={()=>setModal(null)} disabled={saving}>Cancel</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ─── Fees ──────────────────────────────────────────────────────────────────────
function FeesPage({ ctx }) {
  if (!ctx.auth.canFees) return <Empty text="Fees are managed by the fee administrator only."/>;
  return <FeesPageInner ctx={ctx}/>;
}
function FeesPageInner({ ctx }) {
  const { students, feesMap, saveFee, auth, isOnline, setOfflineQueue, settings } = ctx;
  const [filter,    setFilter]    = useState({ form:"", search:"", status:"" });
  const [rcForm,    setRcForm]    = useState("");      // class for the receipt batch ("" = all classes)
  const [rcCopies,  setRcCopies]  = useState(1);       // 1 copy, or 2 (parent + school)
  const [rcWho,     setRcWho]     = useState("paid"); // which students: paid (paid something) | full (fully paid) | part | all
  const [rcBusy,    setRcBusy]    = useState("");      // "" or "done/total" while a PDF is being made
  const [threshold, setThreshold] = useState(50);
  const [modal,     setModal]     = useState(null);
  const [amount,    setAmount]    = useState("");
  const [saving,    setSaving]    = useState(false);
  const [pendingPayments, setPendingPayments] = useState([]);
  const [loadingPending,  setLoadingPending]  = useState(true);
  const [reviewing,       setReviewing]       = useState(null); // submission being viewed full-size
  const [showStmt,        setShowStmt]        = useState(false);

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

  // ── Fee receipts (4 per A4) ──
  const rcClass  = allActive.filter(s => !rcForm || s.form===rcForm);
  const rcPaidOf = s => feesMap[s.id]?.paid||0;
  const rcMatch  = { all:()=>true, paid:s=>rcPaidOf(s)>0, full:s=>rcPaidOf(s)>=TOTAL_FEE, part:s=>rcPaidOf(s)>0 && rcPaidOf(s)<TOTAL_FEE };
  const rcCount  = k => rcClass.filter(rcMatch[k]).length;
  const rcTag    = { all:"Everyone", paid:"Paid", full:"Fully-paid", part:"Part-paid" }[rcWho];
  const rcList = rcClass
    .filter(rcMatch[rcWho])
    .sort((a,b)=>FORMS.indexOf(a.form)-FORMS.indexOf(b.form) || a.name.localeCompare(b.name));
  const rcItems = rcList.map(s => ({ student:s, paid:feesMap[s.id]?.paid||0 }));
  function runReceiptPrint(items) {
    if (!items.length) { alert("No students to print receipts for."); return; }
    domPrint("sbc-receipts", feeReceiptPrintDoc(feeReceiptExpand(items, rcCopies), FEE_CASHIER, feePlan(settings)), "A4 portrait", "0");
  }
  async function runReceiptPdf(items, fileName) {
    if (!items.length) { alert("No students to make receipts for."); return; }
    setRcBusy("0/"+Math.ceil(feeReceiptExpand(items, rcCopies).length/4));
    try {
      const docs = feeReceiptDocs(feeReceiptExpand(items, rcCopies), FEE_CASHIER, feePlan(settings));
      downloadBlob(await htmlDocsToPdfBlob(docs, (d,t)=>setRcBusy(d+"/"+t), 1.5, 0.85), fileName);
    } catch(e) { alert("Could not create the PDF: "+e.message); }
    setRcBusy("");
  }
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
      exportToExcel(scope.map(feeExportRow), `SBCC-Fees-${filter.form.replace(" ","")}-2026-2027.xlsx`, filter.form);
    } else {
      // No form filter — split into one sheet per class, all in one workbook
      if (!allActive.length) { alert("No students to export."); return; }
      const sheets = FORMS.map(f => ({
        name: f,
        rows: allActive.filter(s=>s.form===f).sort((a,b)=>a.name.localeCompare(b.name)).map(feeExportRow),
      }));
      exportToExcelMultiSheet(sheets, `SBCC-Fees-AllForms-2026-2027.xlsx`);
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
        <img src="${SCHOOL_LOGO}" alt="SBCC" style="width:48px;height:48px;object-fit:contain">
        <h2 style="font-size:15px;font-weight:900;margin:3px 0">SAKER BAPTIST COMPREHENSIVE COLLEGE BAWE</h2>
        <p style="font-size:9px;color:#6B7280;margin:0">NGEPTANG · NONI · NW REGION</p>
        <div style="margin-top:7px;background:#B91C1C;display:inline-block;padding:3px 16px;border-radius:10px;color:#fff;font-size:11px;font-weight:800">SCHOOL FEES DEFAULTER LIST</div>
        <p style="font-size:9.5px;color:#374151;margin-top:5px">Below <strong>${threshold}%</strong> · Total fee: <strong>${TOTAL_FEE.toLocaleString()} FCFA</strong> · Academic Year ${ACTIVE_YEAR} · ${new Date().toLocaleDateString("en-GB")}</p>
      </div>
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:12px;text-align:center">
        <div style="background:#EEF2FF;border-radius:6px;padding:8px;border-left:3px solid #0D2340"><div style="font-size:17px;font-weight:900">${debtors.length}</div><div style="font-size:9px;color:#6B7280">Defaulters</div></div>
        <div style="background:#FEF2F2;border-radius:6px;padding:8px;border-left:3px solid #B91C1C"><div style="font-size:17px;font-weight:900;color:#B91C1C">${totBal.toLocaleString()} F</div><div style="font-size:9px;color:#6B7280">Outstanding</div></div>
        <div style="background:#FFFBEB;border-radius:6px;padding:8px;border-left:3px solid #C9962A"><div style="font-size:17px;font-weight:900;color:#C9962A">${totPd.toLocaleString()} F</div><div style="font-size:9px;color:#6B7280">Partially Paid</div></div>
      </div>
      ${sections}
      <p style="font-size:8px;color:#9CA3AF;text-align:center;margin-top:10px">Confidential — administrative use only · Saker Baptist Comprehensive College Bawe</p>
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
        <button onClick={()=>setShowStmt(true)} style={{width:"100%",padding:"10px",marginTop:8,background:C.navy,color:C.white,border:"none",borderRadius:8,fontWeight:800,fontSize:13,cursor:"pointer"}}>
          🧾 Fee statements and instalments
        </button>
        {showStmt && <FeeStatementsModal ctx={ctx} onClose={()=>setShowStmt(false)}/>}
      </div>

      {/* Fee receipts — 4 per A4 page */}
      <div style={{background:C.white,borderRadius:10,padding:12,marginBottom:11,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{fontWeight:800,color:C.navy,fontSize:13,marginBottom:8}}>🧾 Fee Receipts — 4 per A4 page</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:8}}>
          <select style={{...inp,flex:1,minWidth:130}} value={rcForm} onChange={e=>setRcForm(e.target.value)}>
            <option value="">All classes</option>{FORMS.map(f=><option key={f}>{f}</option>)}
          </select>
          <select style={{...inp,flex:1,minWidth:150}} value={rcCopies} onChange={e=>setRcCopies(Number(e.target.value))}>
            <option value={1}>1 copy per student</option><option value={2}>2 copies (parent + school)</option>
          </select>
        </div>
        <select style={{...inp,marginBottom:8}} value={rcWho} onChange={e=>setRcWho(e.target.value)}>
          <option value="paid">Paid something ({rcCount("paid")})</option>
          <option value="full">Fully paid only ({rcCount("full")})</option>
          <option value="part">Part paid only ({rcCount("part")})</option>
          <option value="all">Everyone ({rcCount("all")})</option>
        </select>
        <div style={{fontSize:12,color:C.navy,marginBottom:8}}>{rcList.length} student(s) → {Math.ceil(rcList.length*rcCopies/4)} A4 page(s)</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <Btn onClick={()=>runReceiptPrint(rcItems)} disabled={!!rcBusy}>🖨 Print</Btn>
          <Btn onClick={()=>runReceiptPdf(rcItems, `Fee-Receipts-${(rcForm||"All-classes").replace(/\s+/g,"")}-${rcTag}.pdf`)} outline disabled={!!rcBusy}>{rcBusy?`Making PDF… ${rcBusy}`:"📄 Download PDF"}</Btn>
        </div>
        <div style={{fontSize:10.5,color:C.gray,marginTop:8}}>Each student's card below also has its own Receipt and PDF buttons (same copies setting).</div>
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
            <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center",marginTop:7}}>
              <SmBtn onClick={()=>{setModal(s);setAmount("");}} color={C.navy}>Record Payment</SmBtn>
              <SmBtn onClick={()=>runReceiptPrint([{student:s,paid}])} color={C.green}>🧾 Receipt</SmBtn>
              <SmBtn onClick={()=>runReceiptPdf([{student:s,paid}], `Fee-Receipt-${(s.name||s.id).replace(/\s+/g,"-")}.pdf`)} color={C.navyMid}>📄 PDF</SmBtn>
              {bal>0 && (() => {
                const num = waNumber(s.phone);
                if (!num) return <span style={{fontSize:11,color:C.gray}}>No parent phone</span>;
                const msg = `Dear ${s.parent||"Parent/Guardian"}, this is a reminder from Saker Baptist Comprehensive College Bawe. ${s.name} (${s.form}) has an outstanding school fees balance of ${bal.toLocaleString()} FCFA (paid ${paid.toLocaleString()} of ${TOTAL_FEE.toLocaleString()} FCFA). Kindly settle it at your earliest convenience. Thank you.`;
                const lnk = (href, color, text) => <a href={href} target="_blank" rel="noreferrer" style={{padding:"4px 10px",background:color+"18",color,border:`1px solid ${color}30`,borderRadius:6,fontWeight:700,fontSize:12,textDecoration:"none",whiteSpace:"nowrap"}}>{text}</a>;
                return <>
                  {lnk(`https://wa.me/${num}?text=${encodeURIComponent(msg)}`, "#128C7E", "💬 WhatsApp reminder")}
                  {lnk(`sms:+${num}?&body=${encodeURIComponent(msg)}`, C.navyMid, "✉️ SMS")}
                </>;
              })()}
            </div>
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
              let receiptFor = null;
              try {
                if (!isOnline) {
                  addToOfflineQueue({ type:"payment", data:{ studentId:modal.id, amount:pay } });
                  setOfflineQueue(loadOfflineQueue());
                  alert(`Payment of ${pay.toLocaleString()} FCFA saved on this device. It will upload automatically once you're back online.`);
                } else {
                  const before = feesMap[modal.id]?.paid||0;
                  await saveFee(modal.id, before+pay);
                  receiptFor = { student:modal, paid:before+pay, received:pay };
                }
                setModal(null);
                if (receiptFor && window.confirm(`Payment of ${pay.toLocaleString()} FCFA recorded. Print a receipt now?`)) runReceiptPrint([receiptFor]);
              }
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
// ─── Fee statements and instalments ────────────────────────────────────────────
// Plan kept in app_settings.fee_instalments → [{amt, due}] (amounts must add up to TOTAL_FEE)
function feePlan(settings) {
  const p = parseJSON(settings?.fee_instalments, null);
  if (Array.isArray(p) && p.length && p.reduce((a,x)=>a+Number(x.amt||0),0)===TOTAL_FEE) return p.map(x => ({ amt:Number(x.amt), due:x.due }));
  const y = Number(ACTIVE_YEAR.slice(0,4));
  return [ { amt:15000, due:`${y}-09-30` }, { amt:10000, due:`${y+1}-01-31` }, { amt:10000, due:`${y+1}-04-30` } ];
}
// Money paid fills the instalments in order
function feeStatus(plan, paid) {
  let rem = paid; const today = localToday();
  const rows = plan.map(p => { const got = Math.min(rem, p.amt); rem -= got; const owing = p.amt - got;
    return { ...p, got, owing, status: owing===0 ? "Paid" : p.due < today ? "Overdue" : got>0 ? "Part paid" : "Upcoming" }; });
  return { rows, paid, balance: Math.max(0, TOTAL_FEE-paid),
    overdue: rows.filter(r=>r.due<today).reduce((a,r)=>a+r.owing,0),
    nextDue: rows.find(r=>r.owing>0 && r.due>=today)?.due || "" };
}
function feeStatementDocs(list, feesMap, plan, paysByStudent) {
  const money = n => Number(n).toLocaleString("en-GB")+" FCFA";
  const one = st => {
    const fs = feeStatus(plan, feesMap[st.id]?.paid||0);
    const col = { Paid:"#15803D", Overdue:"#B91C1C", "Part paid":"#C9962A", Upcoming:"#555" };
    const pays = (paysByStudent[st.id]||[]);
    return `<div class="st">
      <div class="hd"><img src="${SCHOOL_LOGO}" alt="" style="width:15mm;height:15mm;object-fit:contain"><div><div class="sc">Saker Baptist Comprehensive College Bawe</div><div class="tt">Fee statement · ${escHtml(ACTIVE_YEAR)} · printed ${new Date().toLocaleDateString("en-GB")}</div></div></div>
      <table class="m"><tr><td><b>Name:</b> ${escHtml(st.name)}</td><td><b>Matricule:</b> ${escHtml(st.id)}</td></tr><tr><td><b>Class:</b> ${escHtml(st.form)}</td><td><b>Parent/Guardian:</b> ${escHtml(st.parent||"—")}</td></tr></table>
      <table class="t"><thead><tr><th>Instalment</th><th>Due date</th><th>Amount</th><th>Paid</th><th>Owing</th><th>Status</th></tr></thead><tbody>
      ${fs.rows.map((r,i)=>`<tr><td>${i+1}</td><td>${fmtDate(r.due)}</td><td>${r.amt.toLocaleString("en-GB")}</td><td>${r.got.toLocaleString("en-GB")}</td><td>${r.owing.toLocaleString("en-GB")}</td><td style="color:${col[r.status]};font-weight:700">${r.status}</td></tr>`).join("")}
      <tr class="tot"><td colspan="2">Total</td><td>${TOTAL_FEE.toLocaleString("en-GB")}</td><td>${fs.paid.toLocaleString("en-GB")}</td><td>${fs.balance.toLocaleString("en-GB")}</td><td></td></tr></tbody></table>
      <div class="sum">Balance: <b>${money(fs.balance)}</b>${fs.overdue>0?` · <span style="color:#B91C1C"><b>Overdue now: ${money(fs.overdue)}</b></span>`:""}${fs.nextDue?` · Next due: <b>${fmtDate(fs.nextDue)}</b>`:""}</div>
      ${pays.length?`<div class="pl">Mobile Money payments confirmed: ${pays.map(x=>`${fmtDate((x.reviewed_at||x.submitted_at||"").slice(0,10))} — ${Number(x.amount).toLocaleString("en-GB")} FCFA`).join(" · ")}</div>`:""}
    </div>`;
  };
  const pages = [];
  for (let i=0; i<list.length; i+=2) pages.push(list.slice(i,i+2));
  return pages.map(g => `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    body{margin:0;padding:8mm 14mm;font-family:Arial,Helvetica,sans-serif;color:#111;font-size:10pt;background:#fff}
    .st{height:136mm;box-sizing:border-box;border-bottom:0.3mm dashed #999;padding:4mm 0;overflow:hidden}
    .hd{display:flex;gap:4mm;align-items:center;border-bottom:0.5mm solid #C9962A;padding-bottom:2mm;margin-bottom:3mm} .sc{font-weight:700;font-size:12pt;color:#0D2340} .tt{font-size:9pt;color:#555}
    table{width:100%;border-collapse:collapse} .m td{padding:1mm 0;font-size:9.5pt} .t{margin-top:3mm} .t th,.t td{border:1px solid #888;padding:1.6mm 2mm;text-align:center;font-size:9pt} .t th{background:#E8EEF7} .tot td{font-weight:700;background:#F3F4F6}
    .sum{margin-top:3mm;font-size:10pt} .pl{margin-top:2mm;font-size:8pt;color:#555}
  </style></head><body>${g.map(one).join("")}</body></html>`);
}

// ─── Fee receipts: 4 per A4 page (2 × 2), for one student or a whole class ───────────
// item = { student, paid, received?, copy? }   received = amount taken in this payment (optional)
function feeReceiptCss() {
  return `@page{size:A4 portrait;margin:0}
  *{box-sizing:border-box}
  html,body{margin:0;padding:0;background:#fff;font-family:Arial,Helvetica,sans-serif;color:#111;-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .pg{width:210mm;height:297mm;display:grid;grid-template-columns:105mm 105mm;grid-template-rows:148.5mm 148.5mm;page-break-after:always;break-after:page;overflow:hidden}
  .pg:last-child{page-break-after:auto;break-after:auto}
  .rc{padding:7mm 8mm;overflow:hidden;font-size:9pt;line-height:1.35}
  .rc:nth-child(odd){border-right:0.3mm dashed #999}
  .rc:nth-child(-n+2){border-bottom:0.3mm dashed #999}
  .hd{display:flex;gap:3mm;align-items:center;border-bottom:0.5mm solid #C9962A;padding-bottom:2mm}
  .sc{font-weight:700;font-size:9.5pt;color:#0D2340;line-height:1.2} .sub{font-size:7.5pt;color:#555}
  .ttl{text-align:center;margin:3mm 0 1.5mm;font-weight:800;font-size:12pt;letter-spacing:0.5pt;color:#0D2340}
  .cp{text-align:center;font-size:7.5pt;font-weight:700;color:#555;margin-bottom:1.5mm}
  .no{display:flex;justify-content:space-between;font-size:8pt;color:#333;margin-bottom:2mm}
  table{width:100%;border-collapse:collapse}
  .who{display:flex;gap:3mm;align-items:flex-start} .who .m{flex:1;min-width:0}
  .m td{padding:0.9mm 0;font-size:9pt;vertical-align:top} .m td.k{width:23mm;color:#555}
  .ph{width:20mm;height:25mm;object-fit:cover;border:0.3mm solid #0D2340;display:block;flex:none;margin-top:1mm}
  .ph0{width:20mm;height:25mm;border:0.3mm dashed #999;flex:none;margin-top:1mm;display:flex;align-items:center;justify-content:center;font-size:6.5pt;color:#888;text-align:center}
  .rcv{margin:2.5mm 0 1.5mm;border:0.4mm solid #15803D;background:#F0FDF4;border-radius:1.5mm;padding:2mm 3mm;text-align:center;font-size:8.5pt}
  .rcv b{display:block;font-size:14pt;color:#15803D}
  .amt{margin-top:2mm} .amt td{border:0.3mm solid #999;padding:1.5mm 2.5mm;font-size:9.5pt} .amt td.r{text-align:right;font-weight:700}
  .st{text-align:center;font-weight:800;font-size:10pt;margin-top:2.5mm;letter-spacing:0.5pt}
  .nx{text-align:center;font-size:8pt;color:#444;margin-top:1mm}
  .sig{display:flex;justify-content:space-between;gap:6mm;margin-top:4mm;font-size:8pt;color:#333}
  .sig div{flex:1;border-top:0.3mm solid #777;padding-top:1mm;text-align:center}
  .ft{margin-top:2mm;text-align:center;font-size:7pt;color:#777}`;
}
function feeReceiptCell(it, cashier, plan) {
  const money = n => Number(n).toLocaleString("en-GB") + " FCFA";
  const st = it.student, paid = Number(it.paid)||0, bal = Math.max(0, TOTAL_FEE - paid);
  const date = fmtDate(localToday());
  const recNo = "FR-" + localToday().replace(/-/g,"") + "-" + st.id;
  const status = paid>=TOTAL_FEE ? "FULLY PAID" : paid>0 ? "PART PAID" : "NOT PAID";
  const col = paid>=TOTAL_FEE ? "#15803D" : paid>0 ? "#C9962A" : "#B91C1C";
  const fs = feeStatus(plan, paid);
  const nextLine = bal>0
    ? (fs.overdue>0 ? `<div class="nx" style="color:#B91C1C;font-weight:700">Overdue now: ${money(fs.overdue)}</div>` : "")
      + (fs.nextDue ? `<div class="nx">Next instalment due: <b>${fmtDate(fs.nextDue)}</b></div>` : "")
    : "";
  return `<div class="rc">
    <div class="hd"><img src="${SCHOOL_LOGO}" alt="" style="width:13mm;height:13mm;object-fit:contain"><div><div class="sc">Saker Baptist Comprehensive College Bawe</div><div class="sub">Ngeptang · Noni · NW Region</div></div></div>
    <div class="ttl">FEE RECEIPT</div>
    ${it.copy ? `<div class="cp">${escHtml(it.copy)}</div>` : ""}
    <div class="no"><span>Receipt No: <b>${escHtml(recNo)}</b></span><span>Date: <b>${date}</b></span></div>
    <div class="who">
      <table class="m">
        <tr><td class="k">Student</td><td><b>${escHtml(st.name)}</b></td></tr>
        <tr><td class="k">Matricule</td><td>${escHtml(st.id)}</td></tr>
        <tr><td class="k">Class</td><td>${escHtml(st.form)}</td></tr>
        <tr><td class="k">Year</td><td>${escHtml(ACTIVE_YEAR)}</td></tr>
        <tr><td class="k">Parent name</td><td>${escHtml(st.parent||"—")}</td></tr>
      </table>
      ${st.photo_url ? `<img class="ph" src="${st.photo_url}" alt="">` : `<div class="ph0">Photo</div>`}
    </div>
    ${it.received ? `<div class="rcv">Amount received today<b>${money(it.received)}</b></div>` : ""}
    <table class="amt">
      <tr><td>Total school fees</td><td class="r">${money(TOTAL_FEE)}</td></tr>
      <tr><td>Total paid to date</td><td class="r" style="color:#15803D">${money(paid)}</td></tr>
      <tr><td>Balance</td><td class="r" style="color:${bal>0?"#B91C1C":"#15803D"}">${money(bal)}</td></tr>
    </table>
    <div class="st" style="color:${col}">${status}</div>
    ${nextLine}
    <div class="sig"><div>Cashier<br><b>${escHtml(cashier||"")}</b></div><div>Parent / Guardian<br><b>${escHtml(st.parent||"")}</b></div></div>
    <div class="ft">Keep this receipt as proof of payment.</div>
  </div>`;
}
// 1 copy each, or 2 (parent + school) side by side
function feeReceiptExpand(items, copies) {
  if (Number(copies)!==2) return items;
  const out = [];
  items.forEach(it => { out.push({ ...it, copy:"PARENT COPY" }); out.push({ ...it, copy:"SCHOOL COPY" }); });
  return out;
}
function feeReceiptPagesHtml(items, cashier, plan) {
  const pages = [];
  for (let i=0; i<items.length; i+=4) pages.push(items.slice(i,i+4));
  return pages.map(g => `<div class="pg">${g.map(it => feeReceiptCell(it, cashier, plan)).join("")}</div>`);
}
const feeReceiptWrap = body => `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Fee receipts</title><style>${feeReceiptCss()}</style></head><body>${body}</body></html>`;
// one finished page per document → for the PDF maker
const feeReceiptDocs = (items, cashier, plan) => feeReceiptPagesHtml(items, cashier, plan).map(pg => feeReceiptWrap(pg));
// everything in one document → for printing
const feeReceiptPrintDoc = (items, cashier, plan) => feeReceiptWrap(feeReceiptPagesHtml(items, cashier, plan).join(""));

function FeeStatementsModal({ ctx, onClose }) {
  const { students, feesMap, settings, saveSetting } = ctx;
  const [form, setForm] = useState("");
  const [plan, setPlan] = useState(() => feePlan(settings).map(x => ({ amt:String(x.amt), due:x.due })));
  const [busy, setBusy] = useState("");
  const list = students.filter(s => s.active && s.reg_status==="registered" && (!form || s.form===form))
    .sort((a,b)=>FORMS.indexOf(a.form)-FORMS.indexOf(b.form) || a.name.localeCompare(b.name));
  const planNum   = plan.map(x => ({ amt:Number(x.amt)||0, due:x.due }));
  const planTotal = planNum.reduce((a,x)=>a+x.amt,0);
  const planOk    = planTotal===TOTAL_FEE && planNum.every(x=>x.amt>0 && /^\d{4}-\d{2}-\d{2}$/.test(x.due||""));
  const owing = list.map(st => ({ st, fs: feeStatus(planNum, feesMap[st.id]?.paid||0) })).filter(x=>x.fs.overdue>0).sort((a,b)=>b.fs.overdue-a.fs.overdue);

  async function savePlan() {
    if (!planOk) { alert(`Amounts must add up to ${TOTAL_FEE.toLocaleString()} FCFA and every instalment needs a date.`); return; }
    setBusy("plan");
    try { await saveSetting("fee_instalments", JSON.stringify(planNum)); alert("Instalment plan saved."); }
    catch(e) { alert("Error: "+e.message); }
    setBusy("");
  }
  async function statements() {
    if (!planOk) { alert("Fix the instalment plan first."); return; }
    if (!list.length) { alert("No students."); return; }
    setBusy("pdf");
    try {
      const paysBy = {};
      const { data } = await supabase.from("payment_submissions").select("student_id,amount,submitted_at,reviewed_at").eq("status","approved");
      (data||[]).forEach(x => { (paysBy[x.student_id] = paysBy[x.student_id]||[]).push(x); });
      downloadBlob(await htmlDocsToPdfBlob(feeStatementDocs(list, feesMap, planNum, paysBy), null, 1.5, 0.85), `Fee-Statements-${(form||"All-classes").replace(/\s+/g,"")}.pdf`);
    } catch(e) { alert("Could not create the PDF: "+e.message); }
    setBusy("");
  }
  async function overduePdf() {
    if (!owing.length) { alert("Nobody is overdue."); return; }
    setBusy("pdf2");
    try {
      const per = 30, pages = [];
      for (let i=0;i<owing.length;i+=per) pages.push(owing.slice(i,i+per));
      const docs = pages.map((chunk,pi) => `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
        body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:12mm 10mm;color:#111;font-size:9.5pt} h1{font-size:14pt;margin:0;text-align:center} h2{font-size:10.5pt;margin:1mm 0 4mm;text-align:center;font-weight:600;color:#333}
        table{width:100%;border-collapse:collapse} th,td{border:1px solid #999;padding:1.3mm 1.6mm;font-size:8.5pt;text-align:center} th{background:#E8EEF7} td.l{text-align:left} .bad{color:#B91C1C;font-weight:700} .foot{margin-top:4mm;font-size:8pt;color:#555;text-align:center}
      </style></head><body><h1>Saker Baptist Comprehensive College Bawe</h1><h2>Fees overdue on ${fmtDate(localToday())}${form?` · ${escHtml(form)}`:""}</h2>
      <table><thead><tr><th>#</th><th>Name</th><th>Class</th><th>Paid</th><th>Overdue</th><th>Balance</th><th>Next due</th><th>Parent phone</th></tr></thead><tbody>
      ${chunk.map((x,i)=>`<tr><td>${pi*per+i+1}</td><td class="l">${escHtml(x.st.name)}</td><td>${escHtml(x.st.form)}</td><td>${x.fs.paid.toLocaleString("en-GB")}</td><td class="bad">${x.fs.overdue.toLocaleString("en-GB")}</td><td>${x.fs.balance.toLocaleString("en-GB")}</td><td>${fmtDate(x.fs.nextDue)}</td><td>${escHtml(x.st.phone||"")}</td></tr>`).join("")}
      </tbody></table><div class="foot">${owing.length} student(s) · total overdue ${owing.reduce((a,x)=>a+x.fs.overdue,0).toLocaleString("en-GB")} FCFA · Page ${pi+1} of ${pages.length}</div></body></html>`);
      downloadBlob(await htmlDocsToPdfBlob(docs, null, 1.5, 0.85), `Fees-Overdue-${(form||"All-classes").replace(/\s+/g,"")}.pdf`);
    } catch(e) { alert("Could not create the PDF: "+e.message); }
    setBusy("");
  }
  function overdueExcel() {
    if (!owing.length) { alert("Nobody is overdue."); return; }
    exportToExcel(owing.map(x => ({ Matricule:x.st.id, Name:x.st.name, Class:x.st.form, "Parent/Guardian":x.st.parent||"", Phone:x.st.phone||"", "Paid (FCFA)":x.fs.paid, "Overdue (FCFA)":x.fs.overdue, "Balance (FCFA)":x.fs.balance, "Next due":x.fs.nextDue })), `SBCC-Fees-Overdue-${(form||"All").replace(/\s+/g,"")}.xlsx`, "Overdue");
  }
  return (
    <Modal title="🧾 Fee statements" onClose={onClose}>
      <div style={{fontSize:12,fontWeight:800,color:C.navy,marginBottom:4}}>Instalment plan ({TOTAL_FEE.toLocaleString()} FCFA)</div>
      {plan.map((x,i) => (
        <div key={i} style={{display:"grid",gridTemplateColumns:"auto 1fr 1.2fr",gap:6,alignItems:"center",marginBottom:6}}>
          <span style={{fontSize:12,color:C.gray}}>{i+1}.</span>
          <input style={inp} type="number" min={0} value={x.amt} onChange={e=>setPlan(p=>p.map((y,j)=>j===i?{...y,amt:e.target.value}:y))}/>
          <input style={inp} type="date" value={x.due} onChange={e=>setPlan(p=>p.map((y,j)=>j===i?{...y,due:e.target.value}:y))}/>
        </div>
      ))}
      <div style={{fontSize:11,color:planTotal===TOTAL_FEE?C.green:C.red,marginBottom:6}}>Total {planTotal.toLocaleString()} / {TOTAL_FEE.toLocaleString()} FCFA</div>
      <SmBtn onClick={savePlan} color={C.navyMid}>{busy==="plan"?"Saving…":"Save plan"}</SmBtn>
      <div style={{marginTop:12}}><Fr label="Class">
        <select style={inp} value={form} onChange={e=>setForm(e.target.value)}><option value="">All classes</option>{FORMS.map(f=><option key={f}>{f}</option>)}</select>
      </Fr></div>
      <div style={{fontSize:12,color:C.navy,marginBottom:8}}>{list.length} student(s) · {owing.length} overdue today</div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        <Btn onClick={statements} disabled={!!busy}>{busy==="pdf"?"Making PDF…":"📄 Statements PDF"}</Btn>
        <Btn onClick={overduePdf} outline disabled={!!busy}>{busy==="pdf2"?"Making PDF…":"Overdue list PDF"}</Btn>
        <Btn onClick={overdueExcel} outline>📊 Overdue Excel</Btn>
      </div>
      <div style={{fontSize:10.5,color:C.gray,marginTop:10}}>Money paid fills the instalments in order. A statement lists Mobile Money payments approved online; cash payments count in the totals.</div>
    </Modal>
  );
}

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
  const t = findTeacher(teachers, u.email);

  const [showPwForm, setShowPwForm] = useState(false);
  const [pass1, setPass1] = useState("");
  const [pass2, setPass2] = useState("");
  const [pwErr, setPwErr] = useState("");
  const [pwOk,  setPwOk]  = useState("");
  const [pwBusy,setPwBusy]= useState(false);
  const [offTick,setOffTick]= useState(0);
  const [snapAt,setSnapAt]= useState(null);
  useEffect(() => { readSnapshot(u.id).then(x => setSnapAt(x?.savedAt||null)); }, [offTick]);
  const offRec = auth.role==="teacher" ? readOfflineLogins()[emailKey(u.email)] : null;

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

      {auth.role==="teacher" && (
        <div data-tick={offTick} style={{background:C.white,borderRadius:13,padding:18,boxShadow:"0 1px 4px rgba(0,0,0,0.06)",marginBottom:12}}>
          <div style={{fontWeight:800,fontSize:14,color:C.navy,marginBottom:6}}>Offline sign-in on this phone</div>
          {offRec ? (
            <>
              <div style={{fontSize:12,color:C.green,marginBottom:8}}>On. Valid until {fmtDate(new Date(offRec.savedAt+OFFLINE_LOGIN_DAYS*86400000).toISOString().slice(0,10))}. It renews each time you sign in with internet.</div>
              <div style={{fontSize:12,color:snapAt?C.green:C.red,marginBottom:8}}>{snapAt ? `Class lists saved on this phone: ${new Date(snapAt).toLocaleString()}.` : "Class lists are NOT saved on this phone yet. Stay signed in with internet for a minute so they download."}</div>
              <SmBtn onClick={()=>{ if (window.confirm("Remove offline sign-in from this phone?")) { removeOfflineLogin(u.email); setOffTick(n=>n+1); } }} color={C.red}>Remove from this phone</SmBtn>
            </>
          ) : (
            <div style={{fontSize:12,color:C.gray}}>Off. Sign in once with internet and tick Allow offline sign-in on the sign-in screen.</div>
          )}
        </div>
      )}

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