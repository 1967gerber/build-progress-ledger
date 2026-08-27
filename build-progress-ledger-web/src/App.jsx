import React, { useState, useEffect, useMemo, useCallback, createContext, useContext } from "react";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";
import {
  Users, Users2, ClipboardList, TrendingUp, TrendingDown, Minus, Flag, Printer,
  LayoutGrid, ListChecks, Plus, Trash2, ChevronDown, BookOpen, Check, Download, Upload,
} from "lucide-react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";

/* ============================================================
   CHECKPOINT DATA — mirrors the BUILD Progress Monitoring workbook
   ============================================================ */
const BOOK_COLOR = {
  "Book 1": { text: "text-[#1F3864]", bg: "bg-[#1F3864]", ring: "ring-[#1F3864]", light: "bg-[#1F3864]/10", border: "border-[#1F3864]" },
  "Book 2": { text: "text-[#2E7D32]", bg: "bg-[#2E7D32]", ring: "ring-[#2E7D32]", light: "bg-[#2E7D32]/10", border: "border-[#2E7D32]" },
  "Book 3": { text: "text-[#7B241C]", bg: "bg-[#7B241C]", ring: "ring-[#7B241C]", light: "bg-[#7B241C]/10", border: "border-[#7B241C]" },
};

const CHECKPOINTS = [
  { id: "1",     label: "Checkpoint 1",     book: "Book 1", lesson: 5,   maxes: [4, 4, 5, 5] },
  { id: "2",     label: "Checkpoint 2",     book: "Book 1", lesson: 10,  maxes: [5, 5, 5, 5] },
  { id: "3",     label: "Checkpoint 3",     book: "Book 1", lesson: 15,  maxes: [5, 5, 5, 5] },
  { id: "4",     label: "Checkpoint 4",     book: "Book 1", lesson: 20,  maxes: [5, 5, 5, 5] },
  { id: "5",     label: "Checkpoint 5",     book: "Book 1", lesson: 25,  maxes: [5, 5, 5, 5] },
  { id: "5alt",  label: "Checkpoint 5 (Alt)", book: "Book 1", lesson: 25, maxes: [5, 5, 5, 5] },
  { id: "6",     label: "Checkpoint 6",     book: "Book 1", lesson: 30,  maxes: [5, 5, 5, 5] },
  { id: "7",     label: "Checkpoint 7",     book: "Book 2", lesson: 35,  maxes: [5, 5, 5, 5] },
  { id: "8",     label: "Checkpoint 8",     book: "Book 2", lesson: 40,  maxes: [5, 5, 5, 5] },
  { id: "9",     label: "Checkpoint 9",     book: "Book 2", lesson: 45,  maxes: [5, 5, 5, 5] },
  { id: "10",    label: "Checkpoint 10",    book: "Book 2", lesson: 50,  maxes: [5, 5, 5, 5] },
  { id: "11",    label: "Checkpoint 11",    book: "Book 2", lesson: 55,  maxes: [5, 5, 5, 5] },
  { id: "12",    label: "Checkpoint 12",    book: "Book 2", lesson: 60,  maxes: [5, 5, 5, 5] },
  { id: "13",    label: "Checkpoint 13",    book: "Book 3", lesson: 65,  maxes: [5, 5, 5, 5] },
  { id: "14",    label: "Checkpoint 14",    book: "Book 3", lesson: 70,  maxes: [5, 5, 5, 5] },
  { id: "15",    label: "Checkpoint 15",    book: "Book 3", lesson: 75,  maxes: [5, 5, 5, 5] },
  { id: "16",    label: "Checkpoint 16",    book: "Book 3", lesson: 80,  maxes: [5, 5, 5, 5] },
  { id: "17",    label: "Checkpoint 17",    book: "Book 3", lesson: 85,  maxes: [5, 5, 5, 5] },
  { id: "18",    label: "Checkpoint 18",    book: "Book 3", lesson: 90,  maxes: [5, 5, 10, 5] },
  { id: "19",    label: "Checkpoint 19",    book: "Book 3", lesson: 95,  maxes: [5, 5, 10, 5] },
  { id: "20",    label: "Checkpoint 20",    book: "Book 3", lesson: 100, maxes: [5, 5, 5, 5] },
];

const CP_BY_ID_BASE = Object.fromEntries(CHECKPOINTS.map((c) => [c.id, c]));

// Merge the fixed 20 checkpoints with a teacher's own custom ones (e.g. a
// "1b" retest), inserting each custom checkpoint right after its parent so
// dropdowns, charts, and reports all stay in a sensible order.
function mergeCheckpoints(custom) {
  const list = [];
  CHECKPOINTS.forEach((c) => {
    list.push(c);
    (custom || [])
      .filter((cc) => cc.parentId === c.id)
      .forEach((cc) => list.push(cc));
  });
  return {
    list,
    order: list.map((c) => c.id),
    byId: Object.fromEntries(list.map((c) => [c.id, c])),
  };
}

const CheckpointsContext = createContext(null);
const useCheckpoints = () => useContext(CheckpointsContext);

// True whenever the signed-in teacher is browsing another teacher's
// ledger via "All Teachers" — disables every mutating control.
const ReadOnlyContext = createContext(false);
const useReadOnly = () => useContext(ReadOnlyContext);

const AREA_LABELS = [
  "Letter/Pattern Recognition",
  "Letter/Pattern Sound",
  "Instant Words",
  "Phonological Awareness",
];
const AREA_COLORS = ["#1F3864", "#2E7D32", "#B8860B", "#7B241C"];
const MASTERY_LINE_COLOR = "#7A4EAB";

/* ============================================================
   SKILL MASTERY DATA — every letter/pattern and instant word,
   tagged with the checkpoint it's first introduced at. Mirrors
   the Skill Mastery Tracker tab of the Excel workbook exactly.
   ============================================================ */
const LETTER_ITEMS = [
  ["i", 1], ["t", 1], ["p", 1], ["n", 1],
  ["s", 2], ["l", 2], ["d", 2], ["f", 2],
  ["h", 4], ["a", 4], ["g", 4], ["o", 4],
  ["k", 5], ["c", 5], ["m", 5], ["r", 5],
  ["b", 7], ["e", 7], ["y", 7], ["j", 7],
  ["u", 8], ["w", 8], ["v", 8], ["x", 8],
  ["z", 10], ["q", 10],
  ["CC (double consonant = 1 sound)", 10],
  ["sh", 11], ["ch", 11],
  ["a-e = (\u0101)", 13], ["i-e = (\u012b)", 13],
  ["o-e = (\u014d)", 14], ["u-e = (\u016b)", 14],
  ["th", 16], ["ck", 16],
  ["ee", 17], ["oo", 17],
  ["ng", 18],
];
const INSTANT_WORD_ITEMS = [
  ["the", 1], ["of", 1], ["and", 1], ["a", 1], ["to", 1],
  ["in", 2], ["is", 2], ["you", 2], ["that", 2], ["it", 2],
  ["he", 4], ["was", 4], ["for", 4], ["on", 4], ["are", 4],
  ["as", 5], ["with", 5], ["his", 5], ["they", 5], ["I", 5],
  ["at", 7], ["be", 7], ["this", 7], ["have", 7], ["from", 7],
  ["or", 8], ["one", 8], ["had", 8], ["by", 8], ["word", 8],
  ["but", 10], ["not", 10], ["what", 10], ["all", 10], ["were", 10],
  ["we", 11], ["when", 11], ["your", 11], ["can", 11], ["said", 11],
  ["there", 13], ["use", 13], ["an", 13], ["each", 13], ["which", 13],
  ["she", 14], ["do", 14], ["how", 14], ["their", 14], ["if", 14],
  ["will", 16], ["up", 16], ["other", 16], ["about", 16], ["out", 16],
  ["many", 17], ["then", 17], ["them", 17], ["these", 17], ["so", 17],
];

const MASTERY_ITEMS = [
  ...LETTER_ITEMS.flatMap(([item, cp]) => ([
    { key: `letter-recog-${item}`, type: "Letter/Pattern", skill: "Recognition", item, cpId: String(cp) },
    { key: `letter-sound-${item}`, type: "Letter/Pattern", skill: "Sound", item, cpId: String(cp) },
  ])),
  ...INSTANT_WORD_ITEMS.map(([item, cp]) => ({ key: `word-${item}`, type: "Instant Word", skill: "Word", item, cpId: String(cp) })),
];

// First checkpoint (by chart order) this item was marked correct, or null.
function masteredAtId(masteryLog, studentId, itemKey, cpOrder) {
  const hits = masteryLog.filter((m) => m.studentId === studentId && m.itemKey === itemKey && m.result === "\u2713");
  if (hits.length === 0) return null;
  return hits.reduce((earliest, m) => {
    if (!earliest) return m.checkpointId;
    return cpOrder.indexOf(m.checkpointId) < cpOrder.indexOf(earliest) ? m.checkpointId : earliest;
  }, null);
}

// % of items introduced by `checkpointId` that are mastered by then.
function cumulativeMasteryPct(masteryLog, studentId, checkpointId, cpOrder) {
  const targetIdx = cpOrder.indexOf(checkpointId);
  if (targetIdx === -1) return null;
  const introduced = MASTERY_ITEMS.filter((it) => cpOrder.indexOf(it.cpId) <= targetIdx);
  if (introduced.length === 0) return null;
  const masteredCount = introduced.filter((it) => {
    const at = masteredAtId(masteryLog, studentId, it.key, cpOrder);
    return at !== null && cpOrder.indexOf(at) <= targetIdx;
  }).length;
  return (masteredCount / introduced.length) * 100;
}

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
const todayISO = () => new Date().toISOString().slice(0, 10);

const entryPct = (e) => {
  const total = e.maxes.reduce((a, b) => a + b, 0);
  return total ? (e.scores.reduce((a, b) => a + b, 0) / total) * 100 : 0;
};

// Summarize one student's trajectory: latest checkpoint, % , and whether
// they're flagged for flat/declining growth since their prior checkpoint.
function summarizeStudent(entries, studentId, cpOrder) {
  const own = entries
    .filter((e) => e.studentId === studentId)
    .sort((a, b) => cpOrder.indexOf(a.checkpointId) - cpOrder.indexOf(b.checkpointId));

  if (own.length === 0) {
    return { count: 0, latest: null, latestPct: null, prevPct: null, trend: "none", flagged: false };
  }
  const latest = own[own.length - 1];
  const prev = own.length > 1 ? own[own.length - 2] : null;
  const latestPct = entryPct(latest);
  const prevPct = prev ? entryPct(prev) : null;

  let trend = "new";
  let flagged = false;
  if (prev) {
    const delta = latestPct - prevPct;
    if (delta > 1) trend = "up";
    else if (delta < -1) trend = "down";
    else trend = "flat";
    flagged = trend !== "up";
  }
  return { count: own.length, latest, latestPct, prevPct, trend, flagged };
}

/* ============================================================
   EXCEL EXPORT — matches the "Progress Log" tab layout of the
   BUILD Books 1-3 Progress Monitoring workbook, so rows can be
   pasted straight into it if desired.
   ============================================================ */
function exportToExcel(data, cps) {
  const header = [
    "Student Name", "Checkpoint", "Book", "Date",
    "Area 1 Correct", "Area 2 Correct", "Area 3 Correct", "Area 4 Correct",
    "Area 1 Max", "Area 2 Max", "Area 3 Max", "Area 4 Max",
    "Total Correct", "Total Max", "Total % Correct",
  ];

  const studentsById = Object.fromEntries(data.students.map((s) => [s.id, s.name]));

  const rows = [...data.entries]
    .sort((a, b) => {
      const nameA = studentsById[a.studentId] || "";
      const nameB = studentsById[b.studentId] || "";
      if (nameA !== nameB) return nameA.localeCompare(nameB);
      return cps.order.indexOf(a.checkpointId) - cps.order.indexOf(b.checkpointId);
    })
    .map((e) => {
      const cp = cps.byId[e.checkpointId];
      const totalCorrect = e.scores.reduce((a, b) => a + b, 0);
      const totalMax = e.maxes.reduce((a, b) => a + b, 0);
      const pct = totalMax ? Math.round((totalCorrect / totalMax) * 1000) / 10 : 0;
      return [
        studentsById[e.studentId] || "(removed student)",
        cp.label.replace("Checkpoint ", ""),
        cp.book,
        e.date,
        e.scores[0], e.scores[1], e.scores[2], e.scores[3],
        e.maxes[0], e.maxes[1], e.maxes[2], e.maxes[3],
        totalCorrect, totalMax, pct,
      ];
    });

  const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
  ws["!cols"] = [
    { wch: 20 }, { wch: 12 }, { wch: 9 }, { wch: 12 },
    { wch: 13 }, { wch: 13 }, { wch: 13 }, { wch: 13 },
    { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 },
    { wch: 12 }, { wch: 10 }, { wch: 13 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Progress Log");

  // Second sheet: skill mastery log (mirrors the Excel workbook's Skill Mastery Tracker)
  const mHeader = ["Student Name", "Item Type", "Skill", "Item", "Introduced At", "Checkpoint Tested", "Result"];
  const itemByKey = Object.fromEntries(MASTERY_ITEMS.map((it) => [it.key, it]));
  const mRows = [...(data.masteryLog || [])]
    .sort((a, b) => {
      const nameA = studentsById[a.studentId] || "";
      const nameB = studentsById[b.studentId] || "";
      if (nameA !== nameB) return nameA.localeCompare(nameB);
      return cps.order.indexOf(a.checkpointId) - cps.order.indexOf(b.checkpointId);
    })
    .map((m) => {
      const it = itemByKey[m.itemKey] || { type: "?", skill: "?", item: m.itemKey, cpId: "?" };
      const cp = cps.byId[m.checkpointId] || { label: m.checkpointId };
      return [
        studentsById[m.studentId] || "(removed student)",
        it.type,
        it.skill,
        it.item,
        `CP${it.cpId}`,
        cp.label,
        m.result,
      ];
    });
  if (mRows.length > 0) {
    const mws = XLSX.utils.aoa_to_sheet([mHeader, ...mRows]);
    mws["!cols"] = [{ wch: 20 }, { wch: 14 }, { wch: 12 }, { wch: 26 }, { wch: 12 }, { wch: 16 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, mws, "Skill Mastery Log");
  }

  const stamp = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `BUILD_Progress_Log_Export_${stamp}.xlsx`);
}

/* ============================================================
   EXCEL IMPORT — reads back a file from exportToExcel (or a
   compatible Progress Log / Skill Mastery Log workbook) and
   merges it into the current session by matching student names.
   This is the reliability safety net: since persistent storage
   isn't always dependable across separate artifact links/tabs,
   the export/import pair lets a teacher treat the .xlsx file as
   the real backup and reload it into any session.
   ============================================================ */
function importFromExcel(file, currentData, cps, onComplete, onError) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array" });
      const studentsByName = new Map(currentData.students.map((s) => [s.name.trim().toLowerCase(), s]));
      let nextStudents = [...currentData.students];
      const findOrCreateStudent = (name) => {
        const key = String(name).trim().toLowerCase();
        if (studentsByName.has(key)) return studentsByName.get(key).id;
        const created = { id: uid(), name: String(name).trim() };
        nextStudents.push(created);
        studentsByName.set(key, created);
        return created.id;
      };

      let entryCount = 0;
      let nextEntries = [...currentData.entries];
      const plSheet = wb.Sheets["Progress Log"];
      if (plSheet) {
        const rows = XLSX.utils.sheet_to_json(plSheet, { defval: "" });
        rows.forEach((row) => {
          const name = row["Student Name"];
          if (!name || name === "(removed student)") return;
          const cpValue = String(row["Checkpoint"]).trim();
          const cp = cps.list.find((c) => c.label.replace("Checkpoint ", "") === cpValue);
          if (!cp) return;
          const studentId = findOrCreateStudent(name);
          const scores = [row["Area 1 Correct"], row["Area 2 Correct"], row["Area 3 Correct"], row["Area 4 Correct"]].map((v) => Number(v) || 0);
          const maxes = [row["Area 1 Max"], row["Area 2 Max"], row["Area 3 Max"], row["Area 4 Max"]].map((v) => Number(v) || 0);
          nextEntries = nextEntries.filter((en) => !(en.studentId === studentId && en.checkpointId === cp.id));
          nextEntries.push({ id: uid(), studentId, checkpointId: cp.id, date: row["Date"] || todayISO(), scores, maxes });
          entryCount++;
        });
      }

      let masteryCount = 0;
      let nextMasteryLog = [...(currentData.masteryLog || [])];
      const mSheet = wb.Sheets["Skill Mastery Log"];
      if (mSheet) {
        const rows = XLSX.utils.sheet_to_json(mSheet, { defval: "" });
        rows.forEach((row) => {
          const name = row["Student Name"];
          if (!name || name === "(removed student)") return;
          const item = MASTERY_ITEMS.find((it) => it.item === row["Item"] && it.skill === row["Skill"]);
          if (!item) return;
          const cp = cps.list.find((c) => c.label === row["Checkpoint Tested"]);
          if (!cp) return;
          const studentId = findOrCreateStudent(name);
          nextMasteryLog = nextMasteryLog.filter((m) => !(m.studentId === studentId && m.itemKey === item.key && m.checkpointId === cp.id));
          nextMasteryLog.push({ id: uid(), studentId, itemKey: item.key, checkpointId: cp.id, result: row["Result"] });
          masteryCount++;
        });
      }

      if (entryCount === 0 && masteryCount === 0) {
        onError(new Error("No matching rows found — make sure this is a Progress Log export from this tool."));
        return;
      }

      onComplete(
        { ...currentData, students: nextStudents, entries: nextEntries, masteryLog: nextMasteryLog },
        { entryCount, masteryCount, studentCount: nextStudents.length - currentData.students.length }
      );
    } catch (err) {
      onError(err);
    }
  };
  reader.onerror = () => onError(new Error("Could not read the file."));
  reader.readAsArrayBuffer(file);
}

function entriesToRows(students, entries, includeTeacher, teacherName, cps) {
  const studentsById = Object.fromEntries(students.map((s) => [s.id, s.name]));
  return [...entries]
    .sort((a, b) => {
      const nameA = studentsById[a.studentId] || "";
      const nameB = studentsById[b.studentId] || "";
      if (nameA !== nameB) return nameA.localeCompare(nameB);
      return cps.order.indexOf(a.checkpointId) - cps.order.indexOf(b.checkpointId);
    })
    .map((e) => {
      const cp = cps.byId[e.checkpointId] || { label: e.checkpointId, book: "?" };
      const totalCorrect = e.scores.reduce((a, b) => a + b, 0);
      const totalMax = e.maxes.reduce((a, b) => a + b, 0);
      const pct = totalMax ? Math.round((totalCorrect / totalMax) * 1000) / 10 : 0;
      const row = [
        studentsById[e.studentId] || "(removed student)",
        cp.label.replace("Checkpoint ", ""),
        cp.book,
        e.date,
        e.scores[0], e.scores[1], e.scores[2], e.scores[3],
        e.maxes[0], e.maxes[1], e.maxes[2], e.maxes[3],
        totalCorrect, totalMax, pct,
      ];
      return includeTeacher ? [teacherName, ...row] : row;
    });
}

async function exportAllTeachersToExcel(teachers) {
  const header = [
    "Teacher", "Student Name", "Checkpoint", "Book", "Date",
    "Area 1 Correct", "Area 2 Correct", "Area 3 Correct", "Area 4 Correct",
    "Area 1 Max", "Area 2 Max", "Area 3 Max", "Area 4 Max",
    "Total Correct", "Total Max", "Total % Correct",
  ];
  let allRows = [];
  for (const t of teachers) {
    const { data: row } = await supabase.from("teacher_data").select("data").eq("user_id", t.user_id).maybeSingle();
    const d = row?.data || { students: [], entries: [], customCheckpoints: [] };
    const cps = mergeCheckpoints(d.customCheckpoints);
    allRows = allRows.concat(entriesToRows(d.students, d.entries, true, t.teacher_name, cps));
  }

  const ws = XLSX.utils.aoa_to_sheet([header, ...allRows]);
  ws["!cols"] = [
    { wch: 16 }, { wch: 20 }, { wch: 12 }, { wch: 9 }, { wch: 12 },
    { wch: 13 }, { wch: 13 }, { wch: 13 }, { wch: 13 },
    { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 },
    { wch: 12 }, { wch: 10 }, { wch: 13 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "All Teachers - Progress Log");
  const stamp = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `BUILD_All_Teachers_Export_${stamp}.xlsx`);
}

/* ============================================================
   MAIN APP
   ============================================================ */
export default function App() {
  const [session, setSession] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [data, setData] = useState({ students: [], entries: [], customCheckpoints: [], masteryLog: [] });
  const [dataLoaded, setDataLoaded] = useState(false);
  const [viewing, setViewing] = useState(null);   // {userId, name} being viewed read-only, or null = self
  const [tab, setTab] = useState("roster");
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");

  const myUserId = session?.user?.id || null;
  const myDisplayName = session?.user?.user_metadata?.display_name || session?.user?.email || "";
  const activeUserId = viewing ? viewing.userId : myUserId;
  const activeName = viewing ? viewing.name : myDisplayName;

  // ---- restore session on load, and keep it in sync ----
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session: s } }) => {
      setSession(s);
      setAuthLoading(false);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, s) => {
      setSession(s);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  // ---- load the active teacher's data whenever it changes ----
  useEffect(() => {
    if (!activeUserId) { setDataLoaded(false); return; }
    setDataLoaded(false);
    (async () => {
      const { data: row, error: fetchError } = await supabase
        .from("teacher_data")
        .select("data")
        .eq("user_id", activeUserId)
        .maybeSingle();
      if (fetchError) setError("Could not load data — " + fetchError.message);
      const loaded = row?.data || {};
      setData({ students: [], entries: [], customCheckpoints: [], masteryLog: [], ...loaded });
      setDataLoaded(true);
    })();
  }, [activeUserId]);

  const saveInBackground = useCallback((next) => {
    if (!myUserId || viewing) return; // never write while viewing someone else (RLS blocks it anyway)
    supabase
      .from("teacher_data")
      .upsert({ user_id: myUserId, teacher_name: myDisplayName, data: next }, { onConflict: "user_id" })
      .then(({ error: saveError }) => {
        if (saveError) setError("Changes are working, but aren't syncing to the database right now — export to Excel to keep a backup. (" + saveError.message + ")");
        else setError("");
      });
  }, [myUserId, myDisplayName, viewing]);

  const signOut = async () => {
    await supabase.auth.signOut();
    setViewing(null);
    setData({ students: [], entries: [], customCheckpoints: [], masteryLog: [] });
  };
  const startViewing = (t) => { setViewing(t); setTab("roster"); };
  const stopViewing = () => { setViewing(null); setTab("roster"); };

  // ---- students ----
  const addStudent = (name) => {
    const trimmed = name.trim();
    if (!trimmed || viewing) return;
    setData((prev) => {
      const next = { ...prev, students: [...prev.students, { id: uid(), name: trimmed }] };
      saveInBackground(next);
      return next;
    });
  };
  const removeStudent = (id) => {
    if (viewing) return;
    setData((prev) => {
      const next = {
        ...prev,
        students: prev.students.filter((s) => s.id !== id),
        entries: prev.entries.filter((e) => e.studentId !== id),
      };
      saveInBackground(next);
      return next;
    });
  };

  // ---- custom checkpoints (e.g. "Checkpoint 1b" retests) ----
  const addCustomCheckpoint = (cp) => {
    if (viewing) return;
    setData((prev) => {
      const next = { ...prev, customCheckpoints: [...(prev.customCheckpoints || []), cp] };
      saveInBackground(next);
      return next;
    });
  };
  const removeCustomCheckpoint = (id) => {
    if (viewing) return;
    setData((prev) => {
      const next = {
        ...prev,
        customCheckpoints: (prev.customCheckpoints || []).filter((c) => c.id !== id),
        entries: prev.entries.filter((e) => e.checkpointId !== id),
      };
      saveInBackground(next);
      return next;
    });
  };

  // ---- entries ----
  const upsertEntry = (entry) => {
    if (viewing) return;
    setData((prev) => {
      const filtered = prev.entries.filter(
        (e) => !(e.studentId === entry.studentId && e.checkpointId === entry.checkpointId)
      );
      const next = { ...prev, entries: [...filtered, entry] };
      saveInBackground(next);
      return next;
    });
  };
  const deleteEntry = (id) => {
    if (viewing) return;
    setData((prev) => {
      const next = { ...prev, entries: prev.entries.filter((e) => e.id !== id) };
      saveInBackground(next);
      return next;
    });
  };

  // ---- skill mastery log (letter/pattern & instant word results) ----
  const upsertMasteryResult = (studentId, itemKey, checkpointId, result) => {
    if (viewing) return;
    setData((prev) => {
      const filtered = (prev.masteryLog || []).filter(
        (m) => !(m.studentId === studentId && m.itemKey === itemKey && m.checkpointId === checkpointId)
      );
      const next = { ...prev, masteryLog: [...filtered, { id: uid(), studentId, itemKey, checkpointId, result }] };
      saveInBackground(next);
      return next;
    });
  };
  const clearMasteryResult = (studentId, itemKey, checkpointId) => {
    if (viewing) return;
    setData((prev) => {
      const next = {
        ...prev,
        masteryLog: (prev.masteryLog || []).filter(
          (m) => !(m.studentId === studentId && m.itemKey === itemKey && m.checkpointId === checkpointId)
        ),
      };
      saveInBackground(next);
      return next;
    });
  };

  // ---- import a backup .xlsx (safety net) ----
  const importFromFile = (file) => {
    if (viewing) { setError("Switch back to your own ledger before importing."); return; }
    const cpsNow = mergeCheckpoints(data.customCheckpoints);
    importFromExcel(
      file,
      data,
      cpsNow,
      (merged, counts) => {
        setData(merged);
        saveInBackground(merged);
        setError("");
        setNotice(`Imported ${counts.entryCount} checkpoint entr${counts.entryCount === 1 ? "y" : "ies"} and ${counts.masteryCount} mastery result${counts.masteryCount === 1 ? "" : "s"}${counts.studentCount > 0 ? ` (${counts.studentCount} new student${counts.studentCount === 1 ? "" : "s"} added)` : ""}.`);
        setTimeout(() => setNotice(""), 6000);
      },
      (err) => setError(`Import failed — ${err.message}`)
    );
  };

  if (authLoading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#FAF8F3]">
        <p className="font-serif text-[#1F3864] text-lg tracking-wide">Loading…</p>
      </div>
    );
  }

  if (!session) {
    return <AuthGate error={error} setError={setError} />;
  }

  if (!dataLoaded) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#FAF8F3]">
        <p className="font-serif text-[#1F3864] text-lg tracking-wide">Opening ledger…</p>
      </div>
    );
  }

  return <SignedInApp
    teacher={myDisplayName} activeName={activeName} onSwitch={signOut} tab={tab} setTab={setTab}
    data={data} error={error} notice={notice} addStudent={addStudent} removeStudent={removeStudent}
    upsertEntry={upsertEntry} deleteEntry={deleteEntry} upsertMasteryResult={upsertMasteryResult} clearMasteryResult={clearMasteryResult}
    addCustomCheckpoint={addCustomCheckpoint} removeCustomCheckpoint={removeCustomCheckpoint} importFromFile={importFromFile}
    viewing={viewing} onStartViewing={startViewing} onStopViewing={stopViewing}
  />;
}

/* ============================================================
   SIGNED-IN APP SHELL — separate component so the log-check
   hook can be called unconditionally at the top level here,
   keeping its state alive across tab switches.
   ============================================================ */
function SignedInApp({
  teacher, activeName, onSwitch, tab, setTab, data, error, notice, addStudent, removeStudent,
  upsertEntry, deleteEntry, upsertMasteryResult, clearMasteryResult, addCustomCheckpoint, removeCustomCheckpoint,
  importFromFile, viewing, onStartViewing, onStopViewing,
}) {
  const cps = useMemo(() => mergeCheckpoints(data.customCheckpoints), [data.customCheckpoints]);
  const logState = useLogCheckState(data.students, data.entries, upsertEntry, cps);
  const [showAllTeachers, setShowAllTeachers] = useState(false);

  return (
    <CheckpointsContext.Provider value={cps}>
    <ReadOnlyContext.Provider value={!!viewing}>
    <div className="min-h-screen bg-[#FAF8F3] text-[#2A2A28]">
      <TopBar
        teacher={teacher}
        onSwitch={onSwitch}
        tab={tab}
        setTab={setTab}
        onExport={() => exportToExcel(data, cps)}
        onImportFile={importFromFile}
        hasData={data.entries.length > 0}
        onOpenAllTeachers={() => setShowAllTeachers(true)}
        viewing={viewing}
      />
      <main>
        <div className="max-w-5xl mx-auto px-5 sm:px-8 pt-6 pb-16">
          {viewing && (
            <div className="mb-5 flex items-center justify-between gap-3 text-sm bg-[#EFEAD9] border border-[#DAD5C6] rounded px-4 py-2.5 print:hidden">
              <span className="text-[#4a4944]">
                <strong className="text-[#201F1D]">Viewing (read-only)</strong> — {viewing.name}'s ledger
              </span>
              <button onClick={onStopViewing} className="text-[#1F3864] underline underline-offset-2 whitespace-nowrap">
                Return to my ledger
              </button>
            </div>
          )}
          {notice && (
            <div className="mb-4 text-sm text-[#2E7D32] bg-[#E2F0D9] border border-[#B8D8AC] rounded px-3 py-2">
              {notice}
            </div>
          )}
          {error && (
            <div className="mb-4 text-sm text-[#8a6a1f] bg-[#F5EAD1] border border-[#E0C98A] rounded px-3 py-2">
              {error}
            </div>
          )}
          {tab === "roster" && (
            <RosterTab students={data.students} onAdd={addStudent} onRemove={removeStudent} entries={data.entries} />
          )}
          {tab === "log" && (
            <LogTab
              students={data.students}
              log={logState}
              customCheckpoints={data.customCheckpoints || []}
              onAddCustomCheckpoint={addCustomCheckpoint}
              onRemoveCustomCheckpoint={removeCustomCheckpoint}
            />
          )}
          {tab === "growth" && (
            <GrowthTab students={data.students} entries={data.entries} onDeleteEntry={deleteEntry} masteryLog={data.masteryLog || []} />
          )}
          {tab === "overview" && (
            <OverviewTab students={data.students} entries={data.entries} />
          )}
          {tab === "mastery" && (
            <MasteryTab
              students={data.students}
              activeName={activeName}
              masteryLog={data.masteryLog || []}
              onLogResult={upsertMasteryResult}
              onClearResult={clearMasteryResult}
            />
          )}
        </div>
      </main>
      {showAllTeachers && (
        <AllTeachersModal
          onClose={() => setShowAllTeachers(false)}
          onView={(t) => { onStartViewing(t); setShowAllTeachers(false); }}
        />
      )}
    </div>
    </ReadOnlyContext.Provider>
    </CheckpointsContext.Provider>
  );
}

/* ============================================================
   ALL TEACHERS MODAL — fetches the live teacher list itself
   ============================================================ */
function AllTeachersModal({ onClose, onView }) {
  const [teachers, setTeachers] = useState([]);
  const [loadingList, setLoadingList] = useState(true);
  const [exporting, setExporting] = useState(false);

  useEffect(() => {
    (async () => {
      const { data, error } = await supabase.from("teacher_data").select("user_id, teacher_name").order("teacher_name");
      if (!error && data) setTeachers(data);
      setLoadingList(false);
    })();
  }, []);

  const handleExportAll = async () => {
    setExporting(true);
    try { await exportAllTeachersToExcel(teachers); } finally { setExporting(false); }
  };

  return (
    <div className="fixed inset-0 bg-black/40 flex items-start sm:items-center justify-center p-4 z-40" onClick={onClose}>
      <div className="bg-white rounded-lg max-w-md w-full max-h-[80vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <div className="px-5 py-4 border-b border-[#E7E2D6] flex items-center justify-between">
          <h3 className="font-serif text-xl text-[#201F1D]">All Teachers</h3>
          <button onClick={onClose} className="text-[#8a8880] hover:text-[#201F1D] text-xl leading-none">×</button>
        </div>
        <div className="px-5 py-3 text-xs text-[#8a8880] bg-[#FAF8F3] border-b border-[#E7E2D6]">
          Any signed-in teacher can view everyone's data here — but only the owner can edit their own.
        </div>
        {loadingList ? (
          <div className="p-6 text-sm text-[#8a8880] text-center">Loading…</div>
        ) : teachers.length === 0 ? (
          <div className="p-6 text-sm text-[#8a8880] text-center">No teachers found yet.</div>
        ) : (
          <ul className="divide-y divide-[#E7E2D6]">
            {teachers.map((t) => (
              <li key={t.user_id} className="flex items-center justify-between px-5 py-3">
                <span className="text-[15px] text-[#201F1D]">{t.teacher_name}</span>
                <button
                  onClick={() => onView({ userId: t.user_id, name: t.teacher_name })}
                  className="text-sm text-[#1F3864] font-medium hover:underline"
                >
                  View →
                </button>
              </li>
            ))}
          </ul>
        )}
        <div className="px-5 py-4 border-t border-[#E7E2D6]">
          <button
            onClick={handleExportAll}
            disabled={exporting || teachers.length === 0}
            className="w-full flex items-center justify-center gap-2 border border-[#DAD5C6] disabled:opacity-40 rounded px-4 py-2.5 text-sm font-medium text-[#4a4944] hover:border-[#1F3864] hover:text-[#1F3864] transition-colors"
          >
            <Download size={14} /> {exporting ? "Exporting…" : "Export all teachers (.xlsx)"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   AUTH GATE — real email/password accounts via Supabase Auth
   ============================================================ */
// Flip to true to reopen public account creation.
const ALLOW_SIGNUP = false;

function AuthGate({ error, setError }) {
  const [mode, setMode] = useState("signin"); // "signin" | "signup"
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [displayName, setDisplayName] = useState("");
  const [busy, setBusy] = useState(false);
  const [localError, setLocalError] = useState("");
  const [checkEmail, setCheckEmail] = useState(false);

  const handleSubmit = async () => {
    setLocalError("");
    setError("");
    if (!email.trim() || !password) { setLocalError("Enter an email and password."); return; }
    if (mode === "signup" && !displayName.trim()) { setLocalError("Enter your name."); return; }
    setBusy(true);
    if (mode === "signup" && !ALLOW_SIGNUP) { setBusy(false); setLocalError("New account creation is closed."); return; }
    if (mode === "signup") {
      const { data, error: signUpError } = await supabase.auth.signUp({
        email: email.trim(),
        password,
        options: { data: { display_name: displayName.trim() } },
      });
      setBusy(false);
      if (signUpError) { setLocalError(signUpError.message); return; }
      if (data.user && !data.session) {
        setCheckEmail(true);
        return;
      }
      if (data.user) {
        await supabase.from("teacher_data").upsert(
          {
            user_id: data.user.id,
            teacher_name: displayName.trim(),
            data: { students: [], entries: [], customCheckpoints: [], masteryLog: [] },
          },
          { onConflict: "user_id" }
        );
      }
    } else {
      const { error: signInError } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
      setBusy(false);
      if (signInError) { setLocalError(signInError.message); return; }
    }
  };

  if (checkEmail) {
    return (
      <div className="min-h-screen bg-[#FAF8F3] flex items-center justify-center px-6">
        <div className="w-full max-w-sm text-center">
          <BookOpen size={22} strokeWidth={1.75} className="mx-auto mb-3 text-[#1F3864]" />
          <h1 className="font-serif text-2xl text-[#201F1D] mb-3">Check your email</h1>
          <p className="text-sm text-[#6b6a67]">
            We sent a confirmation link to <strong>{email}</strong>. Click it, then come back here and sign in.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#FAF8F3] flex items-center justify-center px-6">
      <div className="w-full max-w-sm">
        <div className="text-center mb-8">
          <div className="inline-flex items-center gap-2 text-[#1F3864] mb-3">
            <BookOpen size={22} strokeWidth={1.75} />
            <span className="text-xs tracking-[0.2em] uppercase font-medium">BUILD</span>
          </div>
          <h1 className="font-serif text-3xl text-[#201F1D] leading-tight">Progress Ledger</h1>
          <p className="text-sm text-[#6b6a67] mt-2">
            {mode === "signin" ? "Sign in to open your ledger." : "Create an account to get started."}
          </p>
        </div>
        <div className="bg-white border border-[#E7E2D6] rounded-lg p-5 shadow-sm space-y-3">
          {mode === "signup" && (
            <div>
              <label className="block text-xs uppercase tracking-wide text-[#8a8880] mb-1.5">Your name</label>
              <input
                value={displayName}
                onChange={(e) => setDisplayName(e.target.value)}
                placeholder="e.g. Ms. Alvarez"
                className="w-full border border-[#DAD5C6] rounded px-3 py-2.5 text-[15px] focus:outline-none focus:ring-2 focus:ring-[#1F3864] focus:border-[#1F3864]"
              />
            </div>
          )}
          <div>
            <label className="block text-xs uppercase tracking-wide text-[#8a8880] mb-1.5">Email</label>
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleSubmit()}
              placeholder="you@school.org"
              className="w-full border border-[#DAD5C6] rounded px-3 py-2.5 text-[15px] focus:outline-none focus:ring-2 focus:ring-[#1F3864] focus:border-[#1F3864]"
            />
          </div>
          <div>
            <label className="block text-xs uppercase tracking-wide text-[#8a8880] mb-1.5">Password</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleSubmit()}
              placeholder={mode === "signup" ? "At least 6 characters" : "\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022"}
              className="w-full border border-[#DAD5C6] rounded px-3 py-2.5 text-[15px] focus:outline-none focus:ring-2 focus:ring-[#1F3864] focus:border-[#1F3864]"
            />
          </div>
          {(localError || error) && <p className="text-xs text-[#7B241C]">{localError || error}</p>}
          <button
            onClick={handleSubmit}
            disabled={busy}
            className="w-full bg-[#1F3864] disabled:bg-[#c8c4b8] text-white rounded py-2.5 text-sm font-medium tracking-wide hover:bg-[#16294c] transition-colors"
          >
            {busy ? "Please wait\u2026" : mode === "signin" ? "Sign in" : "Create account"}
          </button>
        </div>
        <p className="text-xs text-[#a8a59b] text-center mt-5">
          {mode === "signin" ? (
            ALLOW_SIGNUP ? (
              <>New here?{" "}
                <button onClick={() => { setMode("signup"); setLocalError(""); }} className="text-[#1F3864] underline underline-offset-2">
                  Create an account
                </button>
              </>
            ) : (
              <>Accounts are set up by your administrator.</>
            )
          ) : (
            <>Already have an account?{" "}
              <button onClick={() => { setMode("signin"); setLocalError(""); }} className="text-[#1F3864] underline underline-offset-2">
                Sign in
              </button>
            </>
          )}
        </p>
        <p className="text-xs text-[#a8a59b] text-center mt-3 leading-relaxed">
          Any signed-in teacher can view every other teacher's data here (but only edit their own) —
          this is a shared team tool, not a private account.
        </p>
      </div>
    </div>
  );
}

/* ============================================================
   TOP BAR / NAV
   ============================================================ */
function TopBar({ teacher, onSwitch, tab, setTab, onExport, onImportFile, hasData, onOpenAllTeachers, viewing }) {
  const fileInputRef = React.useRef(null);
  const tabs = [
    { id: "roster", label: "Students", icon: Users },
    { id: "log", label: "Log a Check", icon: ClipboardList },
    { id: "mastery", label: "Mastery", icon: ListChecks },
    { id: "growth", label: "Growth", icon: TrendingUp },
    { id: "overview", label: "Overview", icon: LayoutGrid },
  ];
  const handleFileChange = (e) => {
    const file = e.target.files?.[0];
    if (file) onImportFile(file);
    e.target.value = ""; // allow re-selecting the same file later
  };
  return (
    <header className="border-b border-[#E7E2D6] bg-[#FAF8F3]/95 backdrop-blur sticky top-0 z-10 print:hidden">
      <div className="max-w-5xl mx-auto px-5 sm:px-8 pt-5 pb-3 flex items-center justify-between">
        <div className="flex items-center gap-2 text-[#1F3864]">
          <BookOpen size={20} strokeWidth={1.75} />
          <div>
            <div className="font-serif text-lg leading-none text-[#201F1D]">Progress Ledger</div>
            <div className="text-[11px] text-[#8a8880] mt-0.5">Signed in as {teacher}</div>
          </div>
        </div>
        <div className="flex items-center gap-4">
          {!viewing && (
            <button
              onClick={onOpenAllTeachers}
              className="flex items-center gap-1.5 text-xs border border-[#DAD5C6] rounded px-2.5 py-1.5 text-[#4a4944] hover:border-[#1F3864] hover:text-[#1F3864] transition-colors"
            >
              <Users2 size={13} /> All Teachers
            </button>
          )}
          <input ref={fileInputRef} type="file" accept=".xlsx" className="hidden" onChange={handleFileChange} />
          <button
            onClick={() => fileInputRef.current?.click()}
            disabled={!!viewing}
            title={viewing ? "Return to your own ledger first" : "Reload data from a previously exported .xlsx backup"}
            className="flex items-center gap-1.5 text-xs border border-[#DAD5C6] disabled:opacity-40 disabled:cursor-not-allowed rounded px-2.5 py-1.5 text-[#4a4944] hover:border-[#1F3864] hover:text-[#1F3864] transition-colors"
          >
            <Upload size={13} /> Import .xlsx
          </button>
          <button
            onClick={onExport}
            disabled={!hasData}
            title={hasData ? "Download all logged data as an Excel file" : "Log at least one checkpoint first"}
            className="flex items-center gap-1.5 text-xs border border-[#DAD5C6] disabled:opacity-40 disabled:cursor-not-allowed rounded px-2.5 py-1.5 text-[#4a4944] hover:border-[#1F3864] hover:text-[#1F3864] transition-colors"
          >
            <Download size={13} /> Export .xlsx
          </button>
          <button onClick={onSwitch} className="text-xs text-[#8a8880] hover:text-[#1F3864] underline underline-offset-2">
            Sign out
          </button>
        </div>
      </div>
      <nav className="max-w-5xl mx-auto px-5 sm:px-8 flex gap-1 overflow-x-auto">
        {tabs.map((t) => {
          const Icon = t.icon;
          const active = tab === t.id;
          return (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              className={`flex items-center gap-1.5 px-3.5 py-2.5 text-sm border-b-2 transition-colors shrink-0 whitespace-nowrap ${
                active
                  ? "border-[#1F3864] text-[#1F3864] font-medium"
                  : "border-transparent text-[#8a8880] hover:text-[#4a4944]"
              }`}
            >
              <Icon size={15} strokeWidth={1.75} />
              {t.label}
            </button>
          );
        })}
      </nav>
    </header>
  );
}

/* ============================================================
   ROSTER TAB
   ============================================================ */
function RosterTab({ students, onAdd, onRemove, entries }) {
  const readOnly = useReadOnly();
  const [name, setName] = useState("");
  const countFor = (id) => entries.filter((e) => e.studentId === id).length;

  return (
    <div>
      <SectionHeader
        title="Students"
        subtitle={readOnly ? "Viewing this teacher's roster (read-only)." : "Add each student you're tracking. Nothing else needs setup."}
      />
      {!readOnly && (
        <div className="flex gap-2 mb-6">
          <input
            value={name}
            onChange={(e) => setName(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") { onAdd(name); setName(""); } }}
            placeholder="Student name"
            className="flex-1 border border-[#DAD5C6] bg-white rounded px-3 py-2.5 text-[15px] focus:outline-none focus:ring-2 focus:ring-[#1F3864] focus:border-[#1F3864]"
          />
          <button
            onClick={() => { onAdd(name); setName(""); }}
            disabled={!name.trim()}
            className="flex items-center gap-1.5 bg-[#1F3864] disabled:bg-[#c8c4b8] text-white rounded px-4 py-2.5 text-sm font-medium hover:bg-[#16294c] transition-colors"
          >
            <Plus size={16} /> Add
          </button>
        </div>
      )}

      {students.length === 0 ? (
        <EmptyState text="No students yet. Add your first student above to start logging checkpoints." />
      ) : (
        <ul className="divide-y divide-[#E7E2D6] border border-[#E7E2D6] rounded-lg bg-white overflow-hidden">
          {students.map((s) => (
            <li key={s.id} className="flex items-center justify-between px-4 py-3">
              <div>
                <div className="text-[15px] text-[#201F1D]">{s.name}</div>
                <div className="text-xs text-[#8a8880]">{countFor(s.id)} checkpoint{countFor(s.id) === 1 ? "" : "s"} logged</div>
              </div>
              {!readOnly && (
                <button
                  onClick={() => onRemove(s.id)}
                  className="text-[#8a8880] hover:text-[#7B241C] p-1.5 rounded transition-colors"
                  aria-label={`Remove ${s.name}`}
                >
                  <Trash2 size={16} />
                </button>
              )}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

/* ============================================================
   LOG A CHECK — shared state hook (lets the Save button live
   outside the scrollable area, as a real layout element)
   ============================================================ */
function useLogCheckState(students, entries, onSave, cps) {
  const [studentId, setStudentId] = useState(students[0]?.id || "");
  const [checkpointId, setCheckpointId] = useState("1");
  const [date, setDate] = useState(todayISO());
  const [scores, setScores] = useState([0, 0, 0, 0]);
  const [saved, setSaved] = useState(false);

  useEffect(() => {
    if (!studentId && students.length) setStudentId(students[0].id);
  }, [students, studentId]);

  useEffect(() => {
    const existing = entries.find((e) => e.studentId === studentId && e.checkpointId === checkpointId);
    if (existing) {
      setDate(existing.date);
      setScores(existing.scores);
    } else {
      setDate(todayISO());
      setScores([0, 0, 0, 0]);
    }
    setSaved(false);
  }, [studentId, checkpointId, entries]);

  const cp = cps.byId[checkpointId] || cps.list[0];

  const setScore = (i, val) => {
    const clamped = Math.max(0, Math.min(cp.maxes[i], Number(val) || 0));
    const next = [...scores];
    next[i] = clamped;
    setScores(next);
  };

  const totalScore = scores.reduce((a, b) => a + b, 0);
  const totalMax = cp.maxes.reduce((a, b) => a + b, 0);
  const pct = totalMax ? Math.round((totalScore / totalMax) * 100) : 0;

  const handleSave = () => {
    if (!studentId) return;
    onSave({ id: uid(), studentId, checkpointId, date, scores, maxes: cp.maxes });
    setSaved(true);
    setTimeout(() => setSaved(false), 1800);
  };

  return {
    studentId, setStudentId, checkpointId, setCheckpointId, date, setDate,
    scores, setScore, cp, totalScore, totalMax, pct, handleSave, saved,
  };
}

/* ============================================================
   LOG A CHECK TAB
   ============================================================ */
function LogTab({ students, log, customCheckpoints, onAddCustomCheckpoint, onRemoveCustomCheckpoint }) {
  const { studentId, setStudentId, checkpointId, setCheckpointId, date, setDate, scores, setScore, cp, totalScore, totalMax, pct, handleSave, saved } = log;
  const cps = useCheckpoints();
  const readOnly = useReadOnly();
  const [showAddCp, setShowAddCp] = useState(false);
  const colors = BOOK_COLOR[cp.book];

  if (students.length === 0) {
    return (
      <div>
        <SectionHeader title="Log a Check" subtitle="Record a checkpoint's scores for a student." />
        <EmptyState text="Add a student on the Students tab first." />
      </div>
    );
  }

  const SaveButton = () => (
    <button
      type="button"
      onClick={handleSave}
      disabled={readOnly}
      style={{ backgroundColor: readOnly ? "#c8c4b8" : saved ? "#2E7D32" : "#1F3864" }}
      className="inline-flex items-center gap-2 text-white rounded px-6 py-3 text-sm font-semibold transition-colors disabled:cursor-not-allowed"
    >
      {readOnly ? "Read-only" : saved ? <><Check size={16} /> Saved</> : "Save entry"}
    </button>
  );

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-4 mb-2">
        <SectionHeader title="Log a Check" subtitle="Pick a student and checkpoint, enter the scores, and save." />
        <SaveButton />
      </div>

      <div className="grid sm:grid-cols-2 gap-4 mb-2">
        <Field label="Student">
          <Select value={studentId} onChange={setStudentId}>
            {students.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </Select>
        </Field>
        <Field label="Checkpoint">
          <Select value={checkpointId} onChange={setCheckpointId}>
            {cps.list.map((c) => (
              <option key={c.id} value={c.id}>{c.label} — {c.book}, Lesson {c.lesson}{c.custom ? " (custom)" : ""}</option>
            ))}
          </Select>
        </Field>
      </div>

      {!readOnly && (
        <button
          onClick={() => setShowAddCp(true)}
          className="text-xs text-[#1F3864] hover:underline mb-5"
        >
          + Add a custom checkpoint (e.g. a retest for students still mastering a set)
        </button>
      )}

      <Field label="Date">
        <input
          type="date"
          value={date}
          onChange={(e) => setDate(e.target.value)}
          className="border border-[#DAD5C6] bg-white rounded px-3 py-2 text-[15px] focus:outline-none focus:ring-2 focus:ring-[#1F3864] focus:border-[#1F3864]"
        />
      </Field>

      <div className={`mt-5 rounded-lg border ${colors.border} border-opacity-30 overflow-hidden`}>
        <div className={`${colors.light} px-4 py-2 flex items-center justify-between`}>
          <span className={`text-sm font-medium ${colors.text}`}>{cp.label} · {cp.book}</span>
          <span className="text-xs text-[#6b6a67]">Through Lesson {cp.lesson}</span>
        </div>
        <div className="bg-white p-4 space-y-4">
          {AREA_LABELS.map((label, i) => (
            <div key={i} className="flex items-center gap-4">
              <span className="text-sm text-[#4a4944] flex-1">{label}</span>
              <input
                type="number"
                min={0}
                max={cp.maxes[i]}
                value={scores[i]}
                onChange={(e) => setScore(i, e.target.value)}
                disabled={readOnly}
                className="w-16 border border-[#DAD5C6] rounded px-2 py-1.5 text-center text-[15px] tabular-nums focus:outline-none focus:ring-2 focus:ring-[#1F3864] focus:border-[#1F3864] disabled:bg-[#F2F0EA] disabled:text-[#8a8880]"
              />
              <span className="text-xs text-[#8a8880] w-12 tabular-nums">/ {cp.maxes[i]}</span>
            </div>
          ))}
        </div>
        <div className="px-4 py-3 bg-[#FAF8F3] border-t border-[#E7E2D6] flex items-center justify-between">
          <span className="text-sm text-[#4a4944]">Total</span>
          <span className="text-sm tabular-nums font-medium text-[#201F1D]">
            {totalScore} / {totalMax} · {pct}%
          </span>
        </div>
      </div>

      <div className="mt-5">
        <SaveButton />
      </div>

      {showAddCp && (
        <AddCheckpointModal
          currentCheckpointId={checkpointId}
          customCheckpoints={customCheckpoints}
          onAdd={(cp) => { onAddCustomCheckpoint(cp); setCheckpointId(cp.id); setShowAddCp(false); }}
          onRemove={onRemoveCustomCheckpoint}
          onClose={() => setShowAddCp(false)}
        />
      )}
    </div>
  );
}

/* ============================================================
   ADD / MANAGE CUSTOM CHECKPOINT MODAL
   ============================================================ */
function AddCheckpointModal({ currentCheckpointId, customCheckpoints, onAdd, onRemove, onClose }) {
  const [parentId, setParentId] = useState(currentCheckpointId && CP_BY_ID_BASE[currentCheckpointId] ? currentCheckpointId : "1");
  const [suffix, setSuffix] = useState("b");

  const parent = CP_BY_ID_BASE[parentId];
  const previewLabel = parent ? `${parent.label}${suffix}` : "";

  const handleAdd = () => {
    if (!parent || !suffix.trim()) return;
    onAdd({
      id: uid(),
      label: previewLabel,
      book: parent.book,
      lesson: parent.lesson,
      maxes: [...parent.maxes],
      parentId: parent.id,
      custom: true,
    });
  };

  return (
    <div className="fixed inset-0 bg-black/40 flex items-start sm:items-center justify-center p-4 z-40" onClick={onClose}>
      <div className="bg-white rounded-lg max-w-md w-full max-h-[85vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <div className="px-5 py-4 border-b border-[#E7E2D6] flex items-center justify-between">
          <h3 className="font-serif text-xl text-[#201F1D]">Custom Checkpoint</h3>
          <button onClick={onClose} className="text-[#8a8880] hover:text-[#201F1D] text-xl leading-none">×</button>
        </div>

        <div className="p-5 space-y-4">
          <p className="text-sm text-[#6b6a67]">
            Add a supplemental checkpoint tied to an existing one — useful for a retest of students who
            haven't yet mastered that set. It uses the same 4-area scoring as its parent checkpoint.
          </p>
          <Field label="Based on">
            <Select value={parentId} onChange={setParentId}>
              {CHECKPOINTS.map((c) => (
                <option key={c.id} value={c.id}>{c.label} — {c.book}, Lesson {c.lesson}</option>
              ))}
            </Select>
          </Field>
          <Field label="Suffix">
            <input
              value={suffix}
              onChange={(e) => setSuffix(e.target.value)}
              placeholder="b"
              maxLength={12}
              className="w-full border border-[#DAD5C6] rounded px-3 py-2 text-[15px] focus:outline-none focus:ring-2 focus:ring-[#1F3864] focus:border-[#1F3864]"
            />
          </Field>
          {previewLabel && (
            <p className="text-xs text-[#8a8880]">Will be added as: <span className="font-medium text-[#201F1D]">{previewLabel}</span></p>
          )}
          <button
            onClick={handleAdd}
            disabled={!suffix.trim()}
            className="w-full flex items-center justify-center gap-2 bg-[#1F3864] disabled:bg-[#c8c4b8] text-white rounded px-4 py-2.5 text-sm font-medium hover:bg-[#16294c] transition-colors"
          >
            <Plus size={15} /> Add checkpoint
          </button>
        </div>

        {customCheckpoints.length > 0 && (
          <div className="border-t border-[#E7E2D6]">
            <div className="px-5 py-2 text-xs uppercase tracking-wide text-[#8a8880]">Your custom checkpoints</div>
            <ul className="divide-y divide-[#E7E2D6]">
              {customCheckpoints.map((c) => (
                <li key={c.id} className="flex items-center justify-between px-5 py-2.5">
                  <span className="text-sm text-[#201F1D]">{c.label} <span className="text-[#8a8880]">— {c.book}</span></span>
                  <button onClick={() => onRemove(c.id)} className="text-[#8a8880] hover:text-[#7B241C]" aria-label={`Remove ${c.label}`}>
                    <Trash2 size={14} />
                  </button>
                </li>
              ))}
            </ul>
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   PARENT LETTER — a plain, printable progress summary built
   entirely from the Mastery tab data. Opens in a new tab so it
   doesn't collide with the Overview tab's print styles.
   ============================================================ */
function esc(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function groupItems(list) {
  const names = list.filter((x) => x.skill === "Recognition").map((x) => x.item);
  const sounds = list.filter((x) => x.skill === "Sound").map((x) => x.item);
  const words = list.filter((x) => x.skill === "Word").map((x) => x.item);
  const parts = [];
  if (names.length) parts.push(`<li><strong>Letter names:</strong> ${esc(names.join(", "))}</li>`);
  if (sounds.length) parts.push(`<li><strong>Letter sounds:</strong> ${esc(sounds.join(", "))}</li>`);
  if (words.length) parts.push(`<li><strong>Words:</strong> ${esc(words.join(", "))}</li>`);
  return parts.join("");
}

function openParentLetter({ studentName, teacherName, cp, cps, masteryLog, studentId }) {
  const targetIdx = cps.order.indexOf(cp.id);
  const introduced = MASTERY_ITEMS.filter((it) => cps.order.indexOf(it.cpId) <= targetIdx);

  const withStatus = introduced.map((it) => {
    const at = masteredAtId(masteryLog, studentId, it.key, cps.order);
    const atIdx = at === null ? -1 : cps.order.indexOf(at);
    return {
      ...it,
      masteredNow: atIdx !== -1 && atIdx <= targetIdx,
      masteredBefore: atIdx !== -1 && atIdx < targetIdx,
    };
  });

  const masteredCount = withStatus.filter((x) => x.masteredNow).length;
  const priorCount = withStatus.filter((x) => x.masteredBefore).length;
  const newly = withStatus.filter((x) => x.masteredNow && !x.masteredBefore);
  const practicing = withStatus.filter((x) => !x.masteredNow);

  const total = introduced.length;
  const opening =
    targetIdx === 0
      ? `This is our first checkpoint of the year. So far we have taught ${total} letter names, letter sounds, and instant words, and ${esc(studentName)} has mastered ${masteredCount} of them.`
      : `So far we have taught ${total} letter names, letter sounds, and instant words. ${esc(studentName)} has mastered ${masteredCount} of them, up from ${priorCount} at the last checkpoint.`;

  const newlySection = newly.length
    ? `<h3>Newly mastered since the last checkpoint</h3><ul>${groupItems(newly)}</ul>`
    : "";

  const practicingSection = practicing.length
    ? `<h3>Still practicing</h3><ul>${groupItems(practicing)}</ul>
       <p>You can help at home by going over the letters and words in this list together — a few minutes a day makes a real difference.</p>`
    : `<p>${esc(studentName)} has mastered everything taught so far. Wonderful work — keep reading together at home.</p>`;

  const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Progress Update — ${esc(studentName)}</title>
<style>
  body { font-family: Georgia, "Times New Roman", serif; color: #201F1D; max-width: 6.5in; margin: 0.75in auto; line-height: 1.55; }
  h1 { font-size: 20pt; margin: 0 0 2px; }
  h3 { font-size: 12pt; margin: 20px 0 6px; }
  .meta { font-size: 10pt; color: #6b6a67; margin-bottom: 22px; }
  ul { margin: 6px 0 0; padding-left: 22px; }
  li { margin-bottom: 3px; }
  .sig { margin-top: 34px; }
  @media print { .noprint { display: none; } }
  .noprint { margin-bottom: 20px; }
  button { font: inherit; padding: 6px 14px; cursor: pointer; }
</style></head><body>
<div class="noprint"><button onclick="window.print()">Print this letter</button></div>
<h1>Reading Progress Update</h1>
<div class="meta">
  <strong>${esc(studentName)}</strong> &nbsp;·&nbsp; ${esc(cp.label)} — ${esc(cp.book)}, Through Lesson ${esc(cp.lesson)}<br>
  ${new Date().toLocaleDateString()}
</div>
<p>Dear Family,</p>
<p>Here is an update on ${esc(studentName)}'s progress in our BUILD reading program.</p>
<p>${opening}</p>
${newlySection}
${practicingSection}
<div class="sig">Sincerely,<br><br>${esc(teacherName)}</div>
</body></html>`;

  const w = window.open("", "_blank");
  if (!w) { alert("Please allow pop-ups for this site to open the letter."); return; }
  w.document.write(html);
  w.document.close();
}

/* ============================================================
   GROWTH TAB
   ============================================================ */
function GrowthTab({ students, entries, onDeleteEntry, masteryLog }) {
  const cps = useCheckpoints();
  const [studentId, setStudentId] = useState(students[0]?.id || "");
  useEffect(() => {
    if (!studentId && students.length) setStudentId(students[0].id);
  }, [students, studentId]);

  const studentEntries = useMemo(
    () => entries
      .filter((e) => e.studentId === studentId)
      .sort((a, b) => cps.order.indexOf(a.checkpointId) - cps.order.indexOf(b.checkpointId)),
    [entries, studentId, cps]
  );

  const CUMULATIVE_KEY = "Cumulative Mastery %";
  const chartData = studentEntries.map((e) => {
    const cp = cps.byId[e.checkpointId] || { label: e.checkpointId };
    const row = { name: cp.label.replace("Checkpoint ", "CP") };
    AREA_LABELS.forEach((label, i) => {
      row[label] = e.maxes[i] ? Math.round((e.scores[i] / e.maxes[i]) * 100) : 0;
    });
    const cum = cumulativeMasteryPct(masteryLog || [], studentId, e.checkpointId, cps.order);
    row[CUMULATIVE_KEY] = cum === null ? null : Math.round(cum);
    return row;
  });

  if (students.length === 0) {
    return (
      <div>
        <SectionHeader title="Growth" subtitle="See a student's progress across checkpoints." />
        <EmptyState text="Add a student on the Students tab first." />
      </div>
    );
  }

  return (
    <div>
      <SectionHeader title="Growth" subtitle="Percent correct per area across every logged checkpoint." />

      <Field label="Student">
        <Select value={studentId} onChange={setStudentId}>
          {students.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
        </Select>
      </Field>

      {studentEntries.length === 0 ? (
        <div className="mt-5"><EmptyState text="No checkpoints logged for this student yet. Log one on the “Log a Check” tab." /></div>
      ) : (
        <>
          <div className="mt-5 bg-white border border-[#E7E2D6] rounded-lg p-4">
            <ResponsiveContainer width="100%" height={320}>
              <LineChart data={chartData} margin={{ top: 8, right: 16, left: -12, bottom: 8 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E7E2D6" />
                <XAxis dataKey="name" tick={{ fontSize: 12, fill: "#6b6a67" }} />
                <YAxis domain={[0, 100]} tick={{ fontSize: 12, fill: "#6b6a67" }} unit="%" />
                <Tooltip contentStyle={{ fontSize: 13, borderRadius: 6, borderColor: "#E7E2D6" }} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                {AREA_LABELS.map((label, i) => (
                  <Line
                    key={label}
                    type="monotone"
                    dataKey={label}
                    stroke={AREA_COLORS[i]}
                    strokeWidth={1.5}
                    strokeOpacity={0.55}
                    dot={{ r: 3 }}
                    connectNulls
                  />
                ))}
                <Line
                  key={CUMULATIVE_KEY}
                  type="monotone"
                  dataKey={CUMULATIVE_KEY}
                  stroke={MASTERY_LINE_COLOR}
                  strokeWidth={3}
                  dot={{ r: 4 }}
                  connectNulls
                />
              </LineChart>
            </ResponsiveContainer>
          </div>
          <p className="text-xs text-[#8a8880] mt-2 print:hidden">
            The 4 thinner lines are each checkpoint's own snapshot score. The bold purple line is <span className="font-medium text-[#7A4EAB]">Cumulative Mastery %</span> from
            the Mastery tab — % of everything introduced so far that's actually been mastered, carrying forward once a skill sticks.
          </p>

          <div className="mt-6">
            <h3 className="text-xs uppercase tracking-wide text-[#8a8880] mb-2">Logged entries</h3>
            <ul className="divide-y divide-[#E7E2D6] border border-[#E7E2D6] rounded-lg bg-white overflow-hidden">
              {studentEntries.map((e) => {
                const cp = cps.byId[e.checkpointId];
                const colors = BOOK_COLOR[cp.book];
                const total = e.scores.reduce((a, b) => a + b, 0);
                const max = e.maxes.reduce((a, b) => a + b, 0);
                return (
                  <li key={e.id} className="flex items-center justify-between px-4 py-2.5">
                    <div className="flex items-center gap-2.5">
                      <span className={`text-[11px] px-1.5 py-0.5 rounded ${colors.light} ${colors.text} font-medium`}>{cp.book}</span>
                      <span className="text-sm text-[#201F1D]">{cp.label}</span>
                      <span className="text-xs text-[#8a8880]">{e.date}</span>
                    </div>
                    <div className="flex items-center gap-3">
                      <span className="text-sm tabular-nums text-[#4a4944]">{total}/{max} · {max ? Math.round((total/max)*100) : 0}%</span>
                      <button onClick={() => onDeleteEntry(e.id)} className="text-[#8a8880] hover:text-[#7B241C]" aria-label="Delete entry">
                        <Trash2 size={14} />
                      </button>
                    </div>
                  </li>
                );
              })}
            </ul>
          </div>
        </>
      )}
    </div>
  );
}

/* ============================================================
   MASTERY TAB — the item-level, auto-carry-forward tracker.
   Pick a student and a "as of" checkpoint; log a result for any
   item that's been introduced. Once an item is marked correct at
   any checkpoint, it shows as mastered at every checkpoint after
   that automatically — no re-entry needed.
   ============================================================ */
function MasteryTab({ students, activeName, masteryLog, onLogResult, onClearResult }) {
  const cps = useCheckpoints();
  const readOnly = useReadOnly();
  const [studentId, setStudentId] = useState(students[0]?.id || "");
  const [checkpointId, setCheckpointId] = useState("1");
  const [hideMastered, setHideMastered] = useState(false);
  const [typeFilter, setTypeFilter] = useState("all"); // all | Recognition | Sound | Word

  useEffect(() => {
    if (!studentId && students.length) setStudentId(students[0].id);
  }, [students, studentId]);

  const targetIdx = cps.order.indexOf(checkpointId);

  const rows = useMemo(() => {
    return MASTERY_ITEMS
      .filter((it) => cps.order.indexOf(it.cpId) <= targetIdx) // only items introduced by now
      .filter((it) => typeFilter === "all" || it.skill === typeFilter)
      .map((it) => {
        const masteredAt = studentId ? masteredAtId(masteryLog, studentId, it.key, cps.order) : null;
        const masteredByNow = masteredAt !== null && cps.order.indexOf(masteredAt) <= targetIdx;
        const currentResult = studentId
          ? masteryLog.find((m) => m.studentId === studentId && m.itemKey === it.key && m.checkpointId === checkpointId)?.result || null
          : null;
        return { ...it, masteredAt, masteredByNow, currentResult };
      });
  }, [masteryLog, studentId, checkpointId, targetIdx, cps, typeFilter]);

  const visibleRows = hideMastered ? rows.filter((r) => !r.masteredByNow) : rows;
  const masteredCount = rows.filter((r) => r.masteredByNow).length;

  const handleLog = (itemKey, currentResult, value) => {
    if (currentResult === value) {
      onClearResult(studentId, itemKey, checkpointId);
    } else {
      onLogResult(studentId, itemKey, checkpointId, value);
    }
  };

  if (students.length === 0) {
    return (
      <div>
        <SectionHeader title="Mastery" subtitle="Track item-level mastery that carries forward automatically." />
        <EmptyState text="Add a student on the Students tab first." />
      </div>
    );
  }

  return (
    <div>
      <SectionHeader
        title="Mastery"
        subtitle="Log a result for a specific letter name, letter sound, or word. Letters are tracked separately for Name and Sound, since a student can know one before the other. Once correct, an item stays marked mastered at every later checkpoint automatically — click the same button again to undo a mistake, or use \u201cFix at CPx\u201d to jump back to where it was first marked."
      />

      <div className="grid sm:grid-cols-2 gap-4 mb-4">
        <Field label="Student">
          <Select value={studentId} onChange={setStudentId}>
            {students.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </Select>
        </Field>
        <Field label="Logging results as of">
          <Select value={checkpointId} onChange={setCheckpointId}>
            {cps.list.map((c) => (
              <option key={c.id} value={c.id}>{c.label} — {c.book}, Lesson {c.lesson}</option>
            ))}
          </Select>
        </Field>
      </div>

      <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
        <div className="flex items-center gap-4">
          <label className="flex items-center gap-2 text-sm text-[#4a4944] cursor-pointer">
            <input type="checkbox" checked={hideMastered} onChange={(e) => setHideMastered(e.target.checked)} />
            Hide already-mastered items
          </label>
          <Select value={typeFilter} onChange={setTypeFilter}>
            <option value="all">All skills</option>
            <option value="Recognition">Letter/Pattern Recognition only</option>
            <option value="Sound">Letter/Pattern Sound only</option>
            <option value="Word">Instant Word only</option>
          </Select>
        </div>
        <div className="flex items-center gap-3">
          <span className="text-xs text-[#8a8880]">{masteredCount} / {rows.length} mastered by this checkpoint</span>
          <button
            onClick={() =>
              openParentLetter({
                studentName: students.find((s) => s.id === studentId)?.name || "your child",
                teacherName: activeName || "",
                cp: cps.byId[checkpointId],
                cps,
                masteryLog,
                studentId,
              })
            }
            disabled={!studentId}
            className="inline-flex items-center gap-1.5 text-xs border border-[#DAD5C6] rounded px-2.5 py-1.5 text-[#4a4944] hover:bg-[#F5F2EA] disabled:opacity-50 transition-colors"
          >
            <Printer size={13} strokeWidth={1.75} />
            Parent letter
          </button>
        </div>
      </div>

      {visibleRows.length === 0 ? (
        <EmptyState text={hideMastered ? "Everything introduced so far is mastered — nice work." : "No items introduced by this checkpoint yet."} />
      ) : (
        <ul className="divide-y divide-[#E7E2D6] border border-[#E7E2D6] rounded-lg bg-white overflow-hidden">
          {visibleRows.map((r) => (
            <li key={r.key} className="flex items-center justify-between px-4 py-2.5 gap-3">
              <div className="min-w-0 flex items-center gap-2.5">
                <span className={`text-[11px] px-1.5 py-0.5 rounded font-medium whitespace-nowrap ${
                  r.skill === "Recognition" ? "bg-[#1F3864]/10 text-[#1F3864]"
                  : r.skill === "Sound" ? "bg-[#2E7D32]/10 text-[#2E7D32]"
                  : "bg-[#B8860B]/10 text-[#8a6a1f]"
                }`}>
                  {r.skill === "Recognition" ? "Name" : r.skill === "Sound" ? "Sound" : "Word"}
                </span>
                <span className="text-[15px] text-[#201F1D] font-medium">{r.item}</span>
                <span className="text-xs text-[#8a8880]">introduced CP{r.cpId}</span>
              </div>
              {r.masteredByNow ? (
                <div className="flex items-center gap-2 shrink-0">
                  <span className="inline-flex items-center gap-1 text-xs font-medium rounded-full px-2.5 py-1 bg-[#E2F0D9] text-[#2E7D32] whitespace-nowrap">
                    <Check size={12} /> Mastered at CP{r.masteredAt}
                  </span>
                  {!readOnly && (r.masteredAt === checkpointId ? (
                    <div className="flex items-center gap-1.5">
                      <button
                        onClick={() => handleLog(r.key, r.currentResult, "\u2713")}
                        title="Click again to undo"
                        className={`px-2.5 py-1 rounded text-xs font-medium border transition-colors ${
                          r.currentResult === "\u2713" ? "bg-[#2E7D32] text-white border-[#2E7D32]" : "border-[#DAD5C6] text-[#4a4944] hover:border-[#2E7D32] hover:text-[#2E7D32]"
                        }`}
                      >
                        {"\u2713"} Correct
                      </button>
                      <button
                        onClick={() => handleLog(r.key, r.currentResult, "X")}
                        title="Click again to undo"
                        className={`px-2.5 py-1 rounded text-xs font-medium border transition-colors ${
                          r.currentResult === "X" ? "bg-[#7B241C] text-white border-[#7B241C]" : "border-[#DAD5C6] text-[#4a4944] hover:border-[#7B241C] hover:text-[#7B241C]"
                        }`}
                      >
                        X
                      </button>
                    </div>
                  ) : (
                    <button
                      onClick={() => setCheckpointId(r.masteredAt)}
                      className="text-xs text-[#8a8880] underline underline-offset-2 hover:text-[#1F3864] whitespace-nowrap"
                    >
                      Fix at CP{r.masteredAt}
                    </button>
                  ))}
                </div>
              ) : (
                <div className="flex items-center gap-1.5 shrink-0">
                  <button
                    onClick={() => handleLog(r.key, r.currentResult, "\u2713")}
                    disabled={readOnly}
                    title="Click again to undo"
                    className={`px-2.5 py-1 rounded text-xs font-medium border transition-colors disabled:opacity-40 disabled:cursor-not-allowed ${
                      r.currentResult === "\u2713" ? "bg-[#2E7D32] text-white border-[#2E7D32]" : "border-[#DAD5C6] text-[#4a4944] hover:border-[#2E7D32] hover:text-[#2E7D32]"
                    }`}
                  >
                    {"\u2713"} Correct
                  </button>
                  <button
                    onClick={() => handleLog(r.key, r.currentResult, "X")}
                    disabled={readOnly}
                    title="Click again to undo"
                    className={`px-2.5 py-1 rounded text-xs font-medium border transition-colors disabled:opacity-40 disabled:cursor-not-allowed ${
                      r.currentResult === "X" ? "bg-[#7B241C] text-white border-[#7B241C]" : "border-[#DAD5C6] text-[#4a4944] hover:border-[#7B241C] hover:text-[#7B241C]"
                    }`}
                  >
                    X
                  </button>
                </div>
              )}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

/* ============================================================
   OVERVIEW TAB — all students at a glance, growth flags,
   and a printable per-checkpoint summary report.
   ============================================================ */
function TrendBadge({ trend }) {
  const map = {
    up:   { icon: TrendingUp,   label: "Growing",       fg: "#2E7D32", bg: "#E2F0D9" },
    flat: { icon: Minus,        label: "Flat",          fg: "#8a6a1f", bg: "#F5EAD1" },
    down: { icon: TrendingDown, label: "Declining",     fg: "#7B241C", bg: "#F5E3E0" },
    new:  { icon: Flag,         label: "Baseline set",  fg: "#1F3864", bg: "#DCE6F1" },
    none: { icon: Minus,        label: "No data yet",   fg: "#8a8880", bg: "#EFEDE6" },
  };
  const m = map[trend] || map.none;
  const Icon = m.icon;
  return (
    <span className="inline-flex items-center gap-1 text-xs font-medium rounded-full px-2.5 py-1" style={{ color: m.fg, backgroundColor: m.bg }}>
      <Icon size={12} /> {m.label}
    </span>
  );
}

function OverviewTab({ students, entries }) {
  const cps = useCheckpoints();
  const [view, setView] = useState("students"); // "students" | "checkpoint" | "matrix"
  const [onlyFlagged, setOnlyFlagged] = useState(false);
  const [checkpointId, setCheckpointId] = useState("1");

  const summaries = useMemo(
    () => students.map((s) => ({ student: s, summary: summarizeStudent(entries, s.id, cps.order) })),
    [students, entries, cps]
  );
  const flaggedCount = summaries.filter((r) => r.summary.flagged).length;

  const visibleRows = onlyFlagged ? summaries.filter((r) => r.summary.flagged) : summaries;
  const sortedRows = [...visibleRows].sort((a, b) => {
    if (a.summary.flagged !== b.summary.flagged) return a.summary.flagged ? -1 : 1;
    return a.student.name.localeCompare(b.student.name);
  });

  if (students.length === 0) {
    return (
      <div>
        <SectionHeader title="Overview" subtitle="See every student at a glance." />
        <EmptyState text="Add a student on the Students tab first." />
      </div>
    );
  }

  const subtitles = {
    students: "Every student, latest checkpoint, and growth trend.",
    checkpoint: "All students' results for one checkpoint — printable.",
    matrix: "Every student across every checkpoint, in one page — printable.",
  };

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-3 mb-2 print:hidden">
        <SectionHeader title="Overview" subtitle={subtitles[view]} />
      </div>

      <div className="flex flex-wrap gap-1 mb-5 print:hidden">
        <button
          onClick={() => setView("students")}
          className={`px-3.5 py-1.5 text-sm rounded-full border transition-colors ${view === "students" ? "bg-[#1F3864] text-white border-[#1F3864]" : "border-[#DAD5C6] text-[#4a4944] hover:border-[#1F3864]"}`}
        >
          All Students
        </button>
        <button
          onClick={() => setView("checkpoint")}
          className={`px-3.5 py-1.5 text-sm rounded-full border transition-colors ${view === "checkpoint" ? "bg-[#1F3864] text-white border-[#1F3864]" : "border-[#DAD5C6] text-[#4a4944] hover:border-[#1F3864]"}`}
        >
          By Checkpoint
        </button>
        <button
          onClick={() => setView("matrix")}
          className={`px-3.5 py-1.5 text-sm rounded-full border transition-colors ${view === "matrix" ? "bg-[#1F3864] text-white border-[#1F3864]" : "border-[#DAD5C6] text-[#4a4944] hover:border-[#1F3864]"}`}
        >
          All Checkpoints
        </button>
      </div>

      {view === "students" && (
        <>
          <div className="flex items-center justify-between mb-3 print:hidden">
            <label className="flex items-center gap-2 text-sm text-[#4a4944] cursor-pointer">
              <input type="checkbox" checked={onlyFlagged} onChange={(e) => setOnlyFlagged(e.target.checked)} />
              Only show flagged students
            </label>
            {flaggedCount > 0 && (
              <span className="text-xs text-[#7B241C] flex items-center gap-1">
                <Flag size={12} /> {flaggedCount} flagged for flat or declining growth
              </span>
            )}
          </div>

          {sortedRows.length === 0 ? (
            <EmptyState text={onlyFlagged ? "No flagged students right now — nice work." : "No students to show."} />
          ) : (
            <ul className="divide-y divide-[#E7E2D6] border border-[#E7E2D6] rounded-lg bg-white overflow-hidden">
              {sortedRows.map(({ student, summary }) => {
                const cp = summary.latest ? cps.byId[summary.latest.checkpointId] : null;
                const colors = cp ? BOOK_COLOR[cp.book] : null;
                return (
                  <li key={student.id} className="flex items-center justify-between px-4 py-3 gap-3">
                    <div className="min-w-0">
                      <div className="text-[15px] text-[#201F1D]">{student.name}</div>
                      <div className="text-xs text-[#8a8880] mt-0.5">
                        {summary.count === 0
                          ? "No checkpoints logged"
                          : <>
                              {cp && <span className={`${colors.text} font-medium`}>{cp.label}</span>}
                              {" · "}{Math.round(summary.latestPct)}%
                              {summary.prevPct !== null && <> (was {Math.round(summary.prevPct)}%)</>}
                            </>
                        }
                      </div>
                    </div>
                    <TrendBadge trend={summary.trend} />
                  </li>
                );
              })}
            </ul>
          )}
        </>
      )}
      {view === "checkpoint" && (
        <CheckpointReport students={students} entries={entries} checkpointId={checkpointId} setCheckpointId={setCheckpointId} />
      )}
      {view === "matrix" && (
        <AllCheckpointsMatrix students={students} entries={entries} />
      )}
    </div>
  );
}

/* ============================================================
   ALL CHECKPOINTS MATRIX — every student × every logged
   checkpoint, in one printable page.
   ============================================================ */
function pctCellColor(pct) {
  if (pct >= 80) return { bg: "#E2F0D9", fg: "#2E7D32" };
  if (pct >= 60) return { bg: "#F5EAD1", fg: "#8a6a1f" };
  return { bg: "#F5E3E0", fg: "#7B241C" };
}

function AllCheckpointsMatrix({ students, entries }) {
  const cps = useCheckpoints();

  // Only show checkpoint columns that actually have at least one entry, in order.
  const usedCpIds = cps.order.filter((id) => entries.some((e) => e.checkpointId === id));

  if (usedCpIds.length === 0) {
    return <EmptyState text="No checkpoints logged yet." />;
  }

  return (
    <div>
      <div className="flex items-center justify-between mb-3 print:hidden">
        <span className="text-xs text-[#8a8880]">{usedCpIds.length} checkpoint{usedCpIds.length === 1 ? "" : "s"} with data · color = % correct (green ≥80%, amber 60–79%, red &lt;60%)</span>
        <button
          onClick={() => window.print()}
          className="flex items-center gap-1.5 border border-[#DAD5C6] rounded px-3.5 py-2 text-sm text-[#4a4944] hover:border-[#1F3864] hover:text-[#1F3864] transition-colors"
        >
          <Printer size={14} /> Print
        </button>
      </div>

      <div className="hidden print:block mb-4">
        <h2 className="font-serif text-2xl">All Checkpoints — Class Overview</h2>
        <p className="text-sm text-[#6b6a67]">Printed {new Date().toLocaleDateString()}</p>
      </div>

      <div className="overflow-x-auto border border-[#E7E2D6] rounded-lg bg-white">
        <table className="text-sm border-collapse">
          <thead>
            <tr>
              <th className="sticky left-0 bg-[#FAF8F3] px-3 py-2 text-left text-xs uppercase tracking-wide text-[#8a8880] border-b border-[#E7E2D6]">Student</th>
              {usedCpIds.map((id) => {
                const cp = cps.byId[id];
                const colors = BOOK_COLOR[cp.book];
                return (
                  <th key={id} className={`px-3 py-2 text-center text-xs font-medium border-b border-[#E7E2D6] whitespace-nowrap ${colors.text}`}>
                    {cp.label.replace("Checkpoint ", "CP")}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {students.map((s) => (
              <tr key={s.id} className="border-b border-[#E7E2D6] last:border-b-0">
                <td className="sticky left-0 bg-white px-3 py-2 text-[#201F1D] whitespace-nowrap">{s.name}</td>
                {usedCpIds.map((id) => {
                  const entry = entries.find((e) => e.studentId === s.id && e.checkpointId === id);
                  if (!entry) {
                    return <td key={id} className="px-3 py-2 text-center text-[#c8c4b8]">—</td>;
                  }
                  const pct = Math.round(entryPct(entry));
                  const c = pctCellColor(pct);
                  return (
                    <td key={id} className="px-2 py-2 text-center">
                      <span className="inline-block w-12 rounded px-1.5 py-1 text-xs font-medium tabular-nums" style={{ backgroundColor: c.bg, color: c.fg }}>
                        {pct}%
                      </span>
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function CheckpointReport({ students, entries, checkpointId, setCheckpointId }) {
  const cps = useCheckpoints();
  const cp = cps.byId[checkpointId] || cps.list[0];
  const colors = BOOK_COLOR[cp.book];

  const rows = useMemo(() => students.map((s) => {
    const entry = entries.find((e) => e.studentId === s.id && e.checkpointId === checkpointId);
    const allForStudent = entries
      .filter((e) => e.studentId === s.id)
      .sort((a, b) => cps.order.indexOf(a.checkpointId) - cps.order.indexOf(b.checkpointId));
    const idx = entry ? allForStudent.findIndex((e) => e.id === entry.id) : -1;
    const prev = idx > 0 ? allForStudent[idx - 1] : null;
    let trend = "none";
    if (entry) {
      if (!prev) trend = "new";
      else {
        const delta = entryPct(entry) - entryPct(prev);
        trend = delta > 1 ? "up" : delta < -1 ? "down" : "flat";
      }
    }
    return { student: s, entry, pct: entry ? entryPct(entry) : null, trend };
  }), [students, entries, checkpointId, cps]);

  const handleExportCheckpoint = () => {
    const header = ["Student Name", "Checkpoint", "Book", "Date", "Area 1", "Area 2", "Area 3", "Area 4", "Total Correct", "Total Max", "Total %"];
    const dataRows = rows.filter((r) => r.entry).map(({ student, entry }) => {
      const total = entry.scores.reduce((a, b) => a + b, 0);
      const max = entry.maxes.reduce((a, b) => a + b, 0);
      return [student.name, cp.label, cp.book, entry.date, entry.scores[0], entry.scores[1], entry.scores[2], entry.scores[3], total, max, max ? Math.round((total / max) * 1000) / 10 : 0];
    });
    const ws = XLSX.utils.aoa_to_sheet([header, ...dataRows]);
    ws["!cols"] = header.map(() => ({ wch: 14 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, cp.label.slice(0, 28));
    XLSX.writeFile(wb, `BUILD_${cp.label.replace(/\s+/g, "_")}_Report_${todayISO()}.xlsx`);
  };

  return (
    <div>
      <div className="flex flex-wrap items-end gap-3 mb-4 print:hidden">
        <div className="flex-1 min-w-[220px]">
          <Field label="Checkpoint">
            <Select value={checkpointId} onChange={setCheckpointId}>
              {cps.list.map((c) => (
                <option key={c.id} value={c.id}>{c.label} — {c.book}, Lesson {c.lesson}{c.custom ? " (custom)" : ""}</option>
              ))}
            </Select>
          </Field>
        </div>
        <button
          onClick={() => window.print()}
          className="flex items-center gap-1.5 border border-[#DAD5C6] rounded px-3.5 py-2.5 text-sm text-[#4a4944] hover:border-[#1F3864] hover:text-[#1F3864] transition-colors"
        >
          <Printer size={14} /> Print
        </button>
        <button
          onClick={handleExportCheckpoint}
          className="flex items-center gap-1.5 border border-[#DAD5C6] rounded px-3.5 py-2.5 text-sm text-[#4a4944] hover:border-[#1F3864] hover:text-[#1F3864] transition-colors"
        >
          <Download size={14} /> Export .xlsx
        </button>
      </div>

      <div className="hidden print:block mb-4">
        <h2 className="font-serif text-2xl">{cp.label} Report — {cp.book}, Through Lesson {cp.lesson}</h2>
        <p className="text-sm text-[#6b6a67]">Printed {new Date().toLocaleDateString()}</p>
      </div>

      <div className={`rounded-lg border ${colors.border} border-opacity-30 overflow-hidden`}>
        <div className={`${colors.light} px-4 py-2 flex items-center justify-between print:hidden`}>
          <span className={`text-sm font-medium ${colors.text}`}>{cp.label} · {cp.book}</span>
          <span className="text-xs text-[#6b6a67]">Through Lesson {cp.lesson}</span>
        </div>
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-[#FAF8F3] text-left text-xs uppercase tracking-wide text-[#8a8880]">
              <th className="px-4 py-2 font-medium">Student</th>
              <th className="px-4 py-2 font-medium">Date</th>
              <th className="px-4 py-2 font-medium text-right">Score</th>
              <th className="px-4 py-2 font-medium text-right">%</th>
              <th className="px-4 py-2 font-medium text-right print:hidden">Trend</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-[#E7E2D6] bg-white">
            {rows.map(({ student, entry, pct, trend }) => {
              const total = entry ? entry.scores.reduce((a, b) => a + b, 0) : null;
              const max = entry ? entry.maxes.reduce((a, b) => a + b, 0) : null;
              return (
                <tr key={student.id}>
                  <td className="px-4 py-2.5 text-[#201F1D]">{student.name}</td>
                  <td className="px-4 py-2.5 text-[#6b6a67]">{entry ? entry.date : "—"}</td>
                  <td className="px-4 py-2.5 text-right tabular-nums text-[#4a4944]">{entry ? `${total} / ${max}` : "Not logged"}</td>
                  <td className="px-4 py-2.5 text-right tabular-nums font-medium text-[#201F1D]">{entry ? `${Math.round(pct)}%` : "—"}</td>
                  <td className="px-4 py-2.5 text-right print:hidden">{entry && <TrendBadge trend={trend} />}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================================================
   SHARED UI PIECES
   ============================================================ */
function SectionHeader({ title, subtitle }) {
  return (
    <div className="mb-5">
      <h2 className="font-serif text-2xl text-[#201F1D]">{title}</h2>
      <p className="text-sm text-[#8a8880] mt-1">{subtitle}</p>
    </div>
  );
}

function EmptyState({ text }) {
  return (
    <div className="border border-dashed border-[#DAD5C6] rounded-lg py-10 px-6 text-center text-sm text-[#8a8880] bg-white">
      {text}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="block">
      <span className="block text-xs uppercase tracking-wide text-[#8a8880] mb-1.5">{label}</span>
      {children}
    </label>
  );
}

function Select({ value, onChange, children }) {
  return (
    <div className="relative">
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="w-full appearance-none border border-[#DAD5C6] bg-white rounded px-3 py-2.5 pr-9 text-[15px] focus:outline-none focus:ring-2 focus:ring-[#1F3864] focus:border-[#1F3864]"
      >
        {children}
      </select>
      <ChevronDown size={16} className="absolute right-3 top-1/2 -translate-y-1/2 text-[#8a8880] pointer-events-none" />
    </div>
  );
}
