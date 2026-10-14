"use client"
/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unused-vars, no-constant-binary-expression */
import React, { useState, useCallback, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { centralStorageEnabled, centralStorageGet, centralStorageSet, isCentralAdmin } from "../lib/central-storage";
import { clearStageASourceFiles, saveStageASourceFiles } from "../lib/stage-a-source-files";
import { fetchStoredFile } from "../lib/central-storage";
import { cleanStatementText, removeStatementBalances, type StatementCell } from "../lib/account-statement-processor";
import {
  Upload, FileSpreadsheet, Download, ChevronDown, X,
  Check, ChevronLeft, CreditCard,
  Shield, Trash2, Sparkles, Clock, Info, RotateCcw
} from "lucide-react";

// ─── Excel helpers ────────────────────────────────────────────────────────────
function readFileBuf(f: File): Promise<ArrayBuffer> {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(e.target!.result as ArrayBuffer);
    r.onerror = () => rej(new Error("فشل قراءة الملف"));
    r.readAsArrayBuffer(f);
  });
}
function parseSheet(buf: ArrayBuffer) {
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const preview = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: "" });
  const populatedCellCount = (row: unknown[] | undefined) => (row ?? []).filter(value => value !== "" && value != null).length;
  const headerRow = populatedCellCount(preview[0]) > 1
    ? 0
    : preview.findIndex((row, index) => index > 0 && populatedCellCount(row) > 1);
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "", range: headerRow > 0 ? headerRow : 0 });
  const headers = rows.length ? Object.keys(rows[0]) : [];
  const cleaned = removeStatementBalances(headers, rows.map(row => Object.fromEntries(
    Object.entries(row).map(([header, value]) => [header, cleanStatementText(value as StatementCell)])
  )));
  return cleaned;
}
function fmtDate(v: unknown): string {
  if (!v) return "";
  if (v instanceof Date) return v.toLocaleDateString("ar-SA");
  const d = new Date(String(v));
  return isNaN(d.getTime()) ? String(v) : d.toLocaleDateString("ar-SA");
}
function toNum(v: unknown): number {
  if (v === "" || v == null) return 0;
  if (typeof v === "number") return v;
  return parseFloat(String(v).replace(/,/g, "")) || 0;
}
function fmtNum(n: number) {
  return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ─── تخزين دائم (IndexedDB) ────────────────────────────────────────────────────
const IDB_NAME = "recon_store";
const IDB_STORE = "kv";
let idbReady: Promise<IDBDatabase> | null = null;

function openIDB(): Promise<IDBDatabase> {
  if (idbReady) return idbReady;
  idbReady = new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return idbReady;
}

async function idbGet<T>(key: string): Promise<T | undefined> {
  const db = await openIDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readonly");
    const req = tx.objectStore(IDB_STORE).get(key);
    req.onsuccess = () => resolve(req.result as T | undefined);
    req.onerror = () => reject(req.error);
  });
}

async function idbSet(key: string, value: unknown): Promise<void> {
  const db = await openIDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, "readwrite");
    tx.objectStore(IDB_STORE).put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function storageGet<T>(key: string, fallback: T): Promise<T> {
  if (centralStorageEnabled) {
    const remote = await centralStorageGet<T>(key);
    if (remote !== undefined) return remote;
    if (isCentralAdmin()) {
      try {
        const local = await idbGet<T>(key);
        if (local !== undefined) {
          await centralStorageSet(key, local);
          return local;
        }
      } catch (error) {
        console.error(`تعذرت ترقية بيانات «${key}» إلى قاعدة البيانات المركزية:`, error);
        throw error;
      }
    }
    return fallback;
  }
  try {
    const w = window as any;
    if (w?.storage?.get) {
      const res = await w.storage.get(key);
      return res ? (JSON.parse(res.value) as T) : fallback;
    }
  } catch { /* */ }
  try {
    const val = await idbGet<T>(key);
    return val !== undefined ? val : fallback;
  } catch {
    try {
      const raw = localStorage.getItem(key);
      return raw ? (JSON.parse(raw) as T) : fallback;
    } catch { return fallback; }
  }
}

async function storageSet(key: string, value: unknown): Promise<void> {
  if (centralStorageEnabled) {
    try {
      await centralStorageSet(key, value);
    } catch (error) {
      console.error(`تعذر حفظ «${key}» في قاعدة البيانات المركزية:`, error);
    }
    return;
  }
  try {
    const w = window as any;
    if (w?.storage?.set) { await w.storage.set(key, JSON.stringify(value)); return; }
  } catch (e) { console.error("storage.set فشل:", e); }
  try {
    await idbSet(key, value);
    return;
  } catch (e) {
    console.error("IndexedDB فشل، محاولة localStorage:", e);
  }
  try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) { console.error("فشل الحفظ:", e); }
}

// ─── Name matching (نفس خوارزمية المنصة الأولى) ──────────────────────────────
const AR_EN: Record<string, string> = {
  "ا":"a","أ":"a","إ":"a","آ":"a","ى":"a","ة":"h","ب":"b","ت":"t","ث":"th",
  "ج":"j","ح":"h","خ":"kh","د":"d","ذ":"th","ر":"r","ز":"z","س":"s","ش":"sh",
  "ص":"s","ض":"d","ط":"t","ظ":"z","ع":"a","غ":"g","ف":"f","ق":"q","ك":"k",
  "ل":"l","م":"m","ن":"n","ه":"h","و":"w","ي":"i","ئ":"i","ء":"","ّ":"",
  "َ":"","ُ":"","ِ":"","ْ":"","ً":"","ٌ":"","ٍ":"",
};
function normName(s: string) {
  const c = s.toLowerCase()
    .replace(/اجمد/g,"احمد")
    .replace(/تحويل\s+الكتروني\s+موبايل:/g,"")
    .replace(/الدفع\s+لصديق\s+من/g,"")
    .replace(/تحويل\s+اي-براق/g,"")
    .replace(/wallet/g,"").replace(/خبز/g,"").trim();
  let t = c.split("").map(ch => AR_EN[ch] ?? ch).join("");
  t = t.replace(/gh/g,"g").replace(/ou/g,"o").replace(/oo/g,"o").replace(/w/g,"o")
       .replace(/ee/g,"i").replace(/y/g,"i").replace(/h$/g,"a")
       .replace(/[^a-z0-9 ]/g," ").replace(/\s+/g," ").trim();
  return t;
}
function levenshtein(a: string, b: string): number {
  const dp = Array.from({ length: a.length+1 }, (_, i) =>
    Array.from({ length: b.length+1 }, (_, j) => i===0?j:j===0?i:0));
  for (let i=1;i<=a.length;i++)
    for (let j=1;j<=b.length;j++)
      dp[i][j]=a[i-1]===b[j-1]?dp[i-1][j-1]:1+Math.min(dp[i-1][j-1],dp[i][j-1],dp[i-1][j]);
  return dp[a.length][b.length];
}
function nameSim(a: string, b: string): number {
  const na=normName(a), nb=normName(b);
  if (!na||!nb) return 0;
  if (na===nb) return 1;
  return 1-levenshtein(na,nb)/Math.max(na.length,nb.length);
}
function wordTypoMatch(a: string, b: string): boolean {
  if (!a || !b) return false;
  if (a === b) return true;
  const maxLen = Math.max(a.length, b.length);
  const d = levenshtein(a, b);
  const threshold = maxLen <= 4 ? 1 : maxLen <= 7 ? 2 : Math.ceil(maxLen * 0.3);
  return d <= threshold;
}
function nameTokens(s: string): string[] {
  return normName(s).split(" ").filter(t => t.length > 1 && t !== "al");
}
function advancedMatchCheck(cashierName: string, bankDesc: string): {
  isMatch: boolean; isApprox: boolean; matchType: string;
} {
  const cT = nameTokens(cashierName);
  const bT = nameTokens(bankDesc);
  if (!cT.length || !bT.length) return { isMatch: false, isApprox: false, matchType: "none" };
  const squished = (arr: string[]) => arr.slice().sort().join("");
  if (squished(cT) === squished(bT)) return { isMatch: true, isApprox: false, matchType: "exact" };
  if (cT.length >= 2 && bT.length >= 2) {
    if (wordTypoMatch(cT[0], bT[0]) && wordTypoMatch(cT[1], bT[1])) {
      return { isMatch: true, isApprox: !(cT[0]===bT[0] && cT[1]===bT[1]), matchType: "firstSecond" };
    }
  }
  if (cT.length >= 4 && bT.length >= 4) {
    if (wordTypoMatch(cT[3], bT[3])) return { isMatch: true, isApprox: cT[3]!==bT[3], matchType: "fourthName" };
  }
  if (cT.length === 1 || bT.length === 1) {
    const shortSide = cT.length === 1 ? cT : bT;
    const longSide = cT.length === 1 ? bT : cT;
    if (wordTypoMatch(shortSide[0], longSide[0])) return { isMatch: true, isApprox: shortSide[0]!==longSide[0], matchType: "typo" };
  }
  for (const ct of cT) {
    for (const bt of bT) {
      if (wordTypoMatch(ct, bt)) return { isMatch: true, isApprox: true, matchType: "typo" };
    }
  }
  return { isMatch: false, isApprox: false, matchType: "none" };
}

// ─── Visa number detection ────────────────────────────────────────────────────
const VISA_NUM_RE = new RegExp("[/-]\\s*(\\d{4})(?!\\d)");
function extractVisaNumber(rawName: string): string | null {
  const m = rawName.match(VISA_NUM_RE);
  return m ? m[1] : null;
}

function isMerchantPurchaseBankDescription(value: string): boolean {
  const description = value.normalize("NFKC").replace(/[\u064B-\u065F\u0670\u0640]/g, "").toLowerCase();
  return (/مشتريات/.test(description) && /عمول/.test(description) && /تجار|بطاق/.test(description)) ||
    /merchant.*(?:purchase|commission)|(?:purchase|commission).*merchant/.test(description);
}

// ─── Auto-detect columns ──────────────────────────────────────────────────────
function autoDetect(headers: string[], hints: string[]): string {
  const scores=headers.map(h=>({h,s:hints.reduce((a,hint)=>a+(h.toLowerCase().includes(hint.toLowerCase())?1:0),0)})).sort((a,b)=>b.s-a.s);
  return scores[0]?.s>0?scores[0].h:"";
}
const HINTS = {
  date:   ["date","تاريخ","value","posting"],
  desc:   ["description","narrative","detail","إيضاح","بيان","وصف","narr","payee","particular"],
  debit:  ["debit","مدين","سحب","paid","withdrawal","dr","مدفوع","خروج","مبالغ مدفوعة"],
  credit: ["credit","دائن","إيداع","received","deposit","cr","دخول","مستلم","مبالغ مستلمة"],
  name:   ["name","اسم","customer","client","employee","موظف","عميل","الزبون","بيان","وصف","narrative","detail","إيضاح"],
  accountType: ["account type","account_type","نوع الحساب","نوع حساب","نوع","حساب","account"],
  authNum: ["تفويض","authorization","auth","رقم","number","ref","مرجع"],
};
const CARD_PAID_HINTS = ["مدفوع بالكارد", "مدفوع بالكرت", "مدفوع كارد", "دفع بالكارد", "دفع بالكرت", "card paid", "card payment", "card amount"];

function detectCardPaidColumn(headers: string[]): string {
  const exact = headers.find(header => CARD_PAID_HINTS.some(hint => header.trim().toLowerCase() === hint.toLowerCase()));
  return exact || autoDetect(headers, CARD_PAID_HINTS);
}

function detectAuthorizationColumn(headers: string[]): string {
  return autoDetect(headers, ["رقم التفويض", "authorization", "auth number", "auth no", "auth code", "approval code", "رقم الموافقة"]);
}

// ─── Types ────────────────────────────────────────────────────────────────────
interface BankRow2 {
  id: number;
  date: string;
  description: string;
  rawAmount: number;
  type: "مدفوع" | "مستلم";
  accountType: string;
  isVisaBank: boolean; // مشتريات عمولة تجار
  orig: Record<string, unknown>;
}

interface AzaRow {
  id: number;
  rawName: string;
  name: string;
  amount: number;
  type: "مدفوع" | "مستلم";
  accountType: string; // محفظة / تطبيق / فيزا
  date: string;
  visaNumber: string | null;
  orig: Record<string, unknown>;
}

interface SonyVisaRow {
  id: number;
  rawName: string;
  name: string;
  amount: number;
  visaNumber: string | null;
  authNumber: string | null; // رقم التفويض
  date: string;
  orig: Record<string, unknown>;
}

interface MatchedVisa {
  id: string;
  sonyId: number;
  azaId: number;
  sonyName: string;
  azaName: string;
  visaNumber: string;
  authNumber: string | null;
  amount: number;
  sonyRawName: string;
  azaRawName: string;
  matchedAt: string;
  note?: string;
}

interface PendingVisaBank {
  id: string;
  bankId: number;
  bankDesc: string;
  authNumber: string | null;
  amount: number;
  date: string;
  movedAt: string;
}

type R2Tab = "overview" | "bank" | "aza" | "sony" | "auth" | "matched" | "pendingBank" | "unmatched";

// ─── DropZone ─────────────────────────────────────────────────────────────────
function DropZone({ file, onFile, onFiles, multiple = false, onClear, rowCount }: { file:File|null; onFile:(f:File)=>void; onFiles?:(files:File[])=>void; multiple?:boolean; onClear:()=>void; rowCount?:number }) {
  const handleDrop=useCallback((e:React.DragEvent)=>{e.preventDefault();const files=Array.from(e.dataTransfer.files);if(files.length){if(onFiles)onFiles(files);else onFile(files[0]);}},[onFile,onFiles]);
  if (file) return (
    <div className="flex items-center justify-between gap-3 p-3.5 bg-blue-500/5 border border-blue-500/25 rounded-lg">
      <div className="flex items-center gap-2.5">
        <FileSpreadsheet className="w-4 h-4 text-blue-600 shrink-0"/>
        <div><div className="text-sm font-medium">{file.name}</div><div className="text-xs text-muted-foreground">{(file.size/1024).toFixed(1)} KB{rowCount !== undefined ? ` · ${rowCount.toLocaleString("ar")} صف` : ""}</div></div>
      </div>
      <button onClick={onClear} className="p-1 hover:bg-muted rounded transition-colors text-muted-foreground"><X className="w-4 h-4"/></button>
    </div>
  );
  return (
    <label onDrop={handleDrop} onDragOver={e=>e.preventDefault()}
      className="flex flex-col items-center gap-2 p-6 border-2 border-dashed border-border rounded-lg cursor-pointer hover:border-blue-500/50 hover:bg-muted/10 transition-all">
      <input type="file" multiple={multiple} accept=".xlsx,.xls,.csv" className="hidden" onChange={e=>{const files=Array.from(e.target.files ?? []);if(files.length){if(onFiles)onFiles(files);else onFile(files[0]);}e.currentTarget.value="";}}/>
      <Upload className="w-6 h-6 text-muted-foreground"/>
      <span className="text-sm">اسحب الملف أو انقر هنا</span>
      <span className="text-xs text-muted-foreground">xlsx · xls · csv</span>
    </label>
  );
}

function Sel({ label, headers, value, onChange }: { label:string; headers:string[]; value:string; onChange:(v:string)=>void; }) {
  return (
    <div className="space-y-1">
      <label className="text-xs font-medium text-muted-foreground">{label}</label>
      <div className="relative">
        <select value={value} onChange={e=>onChange(e.target.value)}
          className="w-full appearance-none bg-background border border-border rounded-md px-3 py-2 text-sm pr-7 focus:outline-none">
          <option value="">— اختر العمود —</option>
          {headers.map(h=><option key={h} value={h}>{h}</option>)}
        </select>
        <ChevronDown className="absolute left-2 top-1/2 -translate-y-1/2 w-3 h-3 text-muted-foreground pointer-events-none"/>
      </div>
    </div>
  );
}

export interface StageAResult {
  id: number;
  source: string;
  name: string;
  originalName: string;
  amount: number;
  databaseName: string;
  databaseAmount: number;
  invoiceTotal: number;
  cashAmount: number;
  checkAmount: number;
  debtAmount: number;
  discountAmount: number;
  type: "مدفوع" | "مستلم";
  userId: string;
  invoiceDate: string;
  registrationTime: string;
  customerNumber: string;
  invoiceNumber: string;
  accountType: string;
  invoice: Record<string, unknown>;
  matched: boolean;
  issue: string;
  category: string;
  categoryIssue: string;
}

export interface StageASummary {
  databaseInvoiceCount: number;
  databaseInvoiceTotal: number;
  databaseCashTotal: number;
  databaseDebtTotal: number;
  databaseCardTotal: number;
  databaseDiscountTotal?: number;
  databaseInvoices?: StageADatabaseInvoice[];
}

export interface StageADatabaseInvoice {
  invoiceNumber: string;
  invoiceDate: string;
  registrationTime: string;
  name: string;
  customerNumber: string;
  userId: string;
  invoiceTotal: number;
  cashAmount: number;
  checkAmount: number;
  cardAmount: number;
  debtAmount: number;
  discountAmount: number;
}

interface StageAInvoiceFileSnapshot {
  fileName: string;
  fileCount: number;
  headers: string[];
  rows: Record<string, unknown>[];
  map: { name: string; paidAmount: string; receivedAmount: string; accountType: string; date: string; registrationTime: string; reference: string; userId: string };
  swap: boolean;
}

function cleanStageAName(value: string): string {
  return String(cleanStatementText(value) ?? "").trim();
}

function stageAEnglishDate(value: unknown): string {
  const text = String(value ?? "").normalize("NFKC")
    .replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, digit => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)))
    .replace(/[\u200e\u200f\u202a-\u202e]/g, "").trim();
  const parts = text.match(/^(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})/);
  if (!parts) return text;
  const [, first, second, third] = parts;
  if (first.length === 4) return `${third.padStart(2, "0")}/${second.padStart(2, "0")}/${first}`;
  return `${first.padStart(2, "0")}/${second.padStart(2, "0")}/${third}`;
}

function stageANameKey(value: string): string {
  return normName(cleanStageAName(value)).split(" ").filter(token => /^[a-z]+$/.test(token)).sort().join(" ");
}

function classifyStageAInvoice(originalName: string, source: string): { category: string; categoryIssue: string } {
  const normalized = originalName.toLowerCase().replace(/[ـ_-]/g, " ").replace(/\s+/g, " ").trim();
  const isJawwalSource = /جوال\s*بي+|jawwal\s*pay|jawwalpay/i.test(source);
  const jawwalPay = isJawwalSource || /جوال\s*بي+|jawwal\s*pay|jawwalpay/i.test(normalized);
  if (jawwalPay) return { category: "جوال بي", categoryIssue: "" };
  const hasVisaNumber = /(?:^|[^\d])\d{4}(?:$|[^\d])/.test(originalName);
  const walletWords = ["محفظة", "wallet"];
  const isWallet = walletWords.some(word => normalized.includes(word));
  if (hasVisaNumber) {
    return { category: "فيزا", categoryIssue: isWallet || source === "محفظة تجارية" ? "مسجلة في ملف المحفظة، لكن رقم 4 أرقام يدل على فيزا" : "" };
  }
  if (/(محمود|شخصية)/.test(normalized) && isWallet) return { category: "محفظة شخصية", categoryIssue: "" };
  if (isWallet) return { category: "بال بي", categoryIssue: "" };
  if (/محفظة|wallet|بال\s*بي|pal\s*pay|palpay/i.test(source)) return { category: "بال بي", categoryIssue: "" };
  if (/بنك|bank/i.test(source)) return { category: source, categoryIssue: "" };
  if (source) return { category: source, categoryIssue: "" };
  return { category: "محفظة", categoryIssue: "لم يظهر رقم فيزا أو اسم محفظة واضح؛ تم تصنيفها كمحفظة" };
}

function stageAVisaNumber(value: string): string {
  return value.match(/(?:^|\D)(\d{4})(?:\D|$)/)?.[1] || "";
}

async function parseInvoiceFiles(files: File[], baseHeaders: string[] = []) {
  if (!files.length) throw new Error("اختر ملفاً واحداً على الأقل.");
  const buffers = await Promise.all(files.map(readFileBuf));
  const worker = new Worker(new URL("../workers/statement-parser.worker.ts", import.meta.url), { type: "module" });
  try {
    const parsed = await new Promise<{ headers: string[]; rows: Record<string, unknown>[]; fileName: string }>((resolve, reject) => {
      worker.onmessage = (event: MessageEvent<
        | { ok: true; headers: string[]; rows: Record<string, unknown>[]; fileName: string }
        | { ok: false; message: string }
      >) => {
        if (event.data.ok) resolve(event.data);
        else reject(new Error(event.data.message));
      };
      worker.onerror = event => reject(new Error(event.message || "تعذّر تشغيل معالج الملفات."));
      worker.onmessageerror = () => reject(new Error("تعذّر استلام نتيجة معالجة الملف."));
      worker.postMessage({
        role: "invoice",
        files: files.map((file, index) => ({ fileName: file.name, buffer: buffers[index] })),
        baseHeaders,
      }, buffers);
    });
    return parsed;
  } finally {
    worker.terminate();
  }
}

function detectInvoiceFileColumns(headers: string[]) {
  return {
    name: headers.find(header => ["البيان", "بيان"].includes(normName(header))) ||
      autoDetect(headers, ["إسم الزبون", "اسم الزبون", "اسم العميل", "customer name", "customer", "client", "name"]),
    paidAmount: detectCardPaidColumn(headers) || autoDetect(headers, ["paid", "مدفوع", "المبلغ المدفوع"]) || autoDetect(headers, HINTS.credit),
    receivedAmount: autoDetect(headers, ["received", "مستلم", "المبلغ المستلم"]) || autoDetect(headers, HINTS.debit),
  };
}

type ProjectSourceFile = { stage: "stageA" | "stageB" | "stageC"; source: string; file: File };
type SavedProjectSourceFile = { id: string; stage: ProjectSourceFile["stage"]; source: string; fileName: string; mimeType: string };

function ProjectSourceFileLauncher({ files, onOpen }: { files: SavedProjectSourceFile[]; onOpen: (file: ProjectSourceFile) => void }) {
  const [opening, setOpening] = useState<string | null>(null);
  const [error, setError] = useState("");
  if (!files.length) return null;
  return <section className="mt-4 rounded-xl border border-indigo-200 bg-indigo-50 p-3">
    <h3 className="text-sm font-bold text-indigo-900">ملفات المشروع في هذه المنصة</h3>
    <div className="mt-2 flex flex-wrap gap-2">
      {files.map(file => <button key={file.id} type="button" disabled={opening !== null} onClick={async () => {
        setOpening(file.id);
        setError("");
        try {
          const openedFile = await fetchStoredFile(file.id, file.fileName, file.mimeType);
          onOpen({ stage: file.stage, source: file.source, file: openedFile });
        } catch (cause) {
          setError(cause instanceof Error ? cause.message : "تعذر فتح ملف المشروع.");
        } finally {
          setOpening(null);
        }
      }} className="rounded-lg border border-indigo-200 bg-white px-3 py-2 text-xs font-medium text-indigo-800 hover:bg-indigo-100 disabled:opacity-50">
        {opening === file.id ? "جارٍ الفتح..." : `فتح ${file.fileName}`}
      </button>)}
    </div>
    {error && <p role="alert" className="mt-2 text-xs text-red-700">{error}</p>}
  </section>;
}

function StageTransferDialog({ dialog, onCancel, onTransferToB, onTransferToC }: { dialog: { invoices: StageAResult[]; from: "B" | "C" } | null; onCancel: () => void; onTransferToB: (invoices: StageAResult[]) => void; onTransferToC: (invoices: StageAResult[]) => void }) {
  if (!dialog) return null;
  const destination = dialog.from === "B" ? "C" : "B";
  const transfer = dialog.from === "B" ? onTransferToC : onTransferToB;
  return <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-3" role="dialog" aria-modal="true" aria-label="اختيار منصة الترحيل">
    <section className="w-full max-w-md rounded-2xl bg-white p-5 shadow-2xl">
      <h2 className="text-lg font-bold">اختيار المنصة</h2>
      <p className="mt-2 text-sm text-slate-600">ترحيل {dialog.invoices.length} فاتورة من {dialog.from} إلى المنصة الأخرى؟</p>
      <div className="mt-5 flex gap-2">
        <button type="button" onClick={() => transfer(dialog.invoices)} className="flex-1 rounded-lg bg-blue-600 px-4 py-2.5 text-sm font-semibold text-white">منصة {destination}</button>
        <button type="button" onClick={onCancel} className="rounded-lg border px-4 py-2.5 text-sm">إلغاء</button>
      </div>
    </section>
  </div>;
}

function StageAPlatform({ onBack, onTransferToB, onTransferVisaToC, onSafeActivities, initialFile }: { onBack: () => void; onTransferToB: (results: StageAResult[], summary: StageASummary) => void; onTransferVisaToC: (results: StageAResult[]) => void; onSafeActivities?: (activities: string[]) => void; initialFile?: ProjectSourceFile }) {
  const [databaseFile, setDatabaseFile] = useState<File | null>(null);
  const [databaseFileName, setDatabaseFileName] = useState("");
  const [databaseHeaders, setDatabaseHeaders] = useState<string[]>([]);
  const [databaseRows, setDatabaseRows] = useState<Record<string, unknown>[]>([]);
  const [databaseMap, setDatabaseMap] = useState({ name: "", amount: "", invoiceTotal: "", cashAmount: "", checkAmount: "", debtAmount: "", userId: "", date: "", registrationTime: "", customerNumber: "", invoiceNumber: "" });
  const [removeRepeatedHeaders, setRemoveRepeatedHeaders] = useState(true);
  const invoiceMapDefaults = { name: "", paidAmount: "", receivedAmount: "", accountType: "", date: "", registrationTime: "", reference: "", userId: "" };
  const [invoiceFiles, setInvoiceFiles] = useState<Record<string, File | null>>({ bank: null, wallet: null, visa: null, jawwal: null });
  const [invoiceFileNames, setInvoiceFileNames] = useState<Record<string, string>>({ bank: "", wallet: "", visa: "", jawwal: "" });
  const [invoiceFileCounts, setInvoiceFileCounts] = useState<Record<string, number>>({ bank: 0, wallet: 0, visa: 0, jawwal: 0 });
  const [invoiceFileSnapshots, setInvoiceFileSnapshots] = useState<Record<string, StageAInvoiceFileSnapshot[]>>({});
  const [showInvoiceSnapshots, setShowInvoiceSnapshots] = useState<Record<string, boolean>>({});
  const [invoiceProcessing, setInvoiceProcessing] = useState<Record<string, boolean>>({});
  const [invoiceHeaders, setInvoiceHeaders] = useState<Record<string, string[]>>({ bank: [], wallet: [], visa: [], jawwal: [] });
  const [invoiceRows, setInvoiceRows] = useState<Record<string, Record<string, unknown>[]>>({ bank: [], wallet: [], visa: [], jawwal: [] });
  const [invoiceMaps, setInvoiceMaps] = useState<Record<string, typeof invoiceMapDefaults>>({
    bank: { ...invoiceMapDefaults }, wallet: { ...invoiceMapDefaults }, visa: { ...invoiceMapDefaults }, jawwal: { ...invoiceMapDefaults }
  });
  const [invoiceSwaps, setInvoiceSwaps] = useState<Record<string, boolean>>({ bank: true, wallet: true, visa: true, jawwal: true });
  const [invoiceRemoveRepeatedHeaders, setInvoiceRemoveRepeatedHeaders] = useState<Record<string, boolean>>({ bank: true, wallet: true, visa: true, jawwal: true });
  const [customInvoicePlatforms, setCustomInvoicePlatforms] = useState<Array<{ id: string; name: string }>>([]);
  const [newInvoicePlatformName, setNewInvoicePlatformName] = useState("");
  const [results, setResults] = useState<StageAResult[]>([]);
  const [databaseSummary, setDatabaseSummary] = useState<StageASummary>({ databaseInvoiceCount: 0, databaseInvoiceTotal: 0, databaseCashTotal: 0, databaseDebtTotal: 0, databaseCardTotal: 0 });
  const [resultFilter, setResultFilter] = useState<"all" | "matched" | "unmatched" | "wallet" | "جوال بي" | "visa">("all");
  const [selectedResultIds, setSelectedResultIds] = useState<Set<string>>(new Set());
  const [resultNameSearch, setResultNameSearch] = useState("");
  const [resultAmountSearch, setResultAmountSearch] = useState("");
  const [resultMinAmount, setResultMinAmount] = useState("");
  const [resultMaxAmount, setResultMaxAmount] = useState("");
  const [stageASourceFilter, setStageASourceFilter] = useState("all");
  const [matchedSearch, setMatchedSearch] = useState("");
  const [selectedMatchedIds, setSelectedMatchedIds] = useState<Set<string>>(new Set());
  const [replacementSearch, setReplacementSearch] = useState("");
  const [stageAPage, setStageAPage] = useState(1);
  const [error, setError] = useState<string | null>(null);
  const [showDeleteMenu, setShowDeleteMenu] = useState(false);
  const [showDatabaseInvoices, setShowDatabaseInvoices] = useState(false);
  const [databaseInvoicePage, setDatabaseInvoicePage] = useState(1);
  const [transferDestinationDialog, setTransferDestinationDialog] = useState<{ label: string; results: StageAResult[] } | null>(null);
  const stageADataLoaded = useRef(false);

  useEffect(() => {
    storageGet<any>("stage_a_session", null).then(session => {
      if (session) {
        setResults(session.results ?? []);
        setDatabaseSummary(session.summary ?? { databaseInvoiceCount: 0, databaseInvoiceTotal: 0, databaseCashTotal: 0, databaseDebtTotal: 0, databaseCardTotal: 0 });
        setDatabaseFileName(session.databaseFileName ?? "");
        setDatabaseHeaders(session.databaseHeaders ?? []);
        setDatabaseRows(session.databaseRows ?? []);
        setDatabaseMap(prev => ({ ...prev, ...(session.databaseMap ?? {}) }));
        setRemoveRepeatedHeaders(session.removeRepeatedHeaders ?? true);
        setInvoiceFileNames(session.invoiceFileNames ?? {});
        setInvoiceFileCounts(session.invoiceFileCounts ?? {});
        setInvoiceFileSnapshots(session.invoiceFileSnapshots ?? {});
        setInvoiceHeaders(prev => ({ ...prev, ...(session.invoiceHeaders ?? {}) }));
        setInvoiceRows(prev => ({ ...prev, ...(session.invoiceRows ?? {}) }));
        const restoredInvoiceMaps = { ...Object.fromEntries(Object.keys(session.invoiceHeaders ?? {}).map(source => [source, { ...invoiceMapDefaults, ...(session.invoiceMaps?.[source] ?? {}) }])), ...(session.invoiceMaps ?? {}) } as Record<string, typeof invoiceMapDefaults>;
        (Object.entries(session.invoiceHeaders ?? {}) as Array<[string, string[]]>).forEach(([source, headers]) => {
          if (session.invoiceSwaps?.[source] === true || !headers.length) return;
          const detected = detectInvoiceFileColumns(headers);
          restoredInvoiceMaps[source] = {
            ...restoredInvoiceMaps[source],
            name: restoredInvoiceMaps[source]?.name || detected.name,
            paidAmount: detected.paidAmount,
            receivedAmount: detected.receivedAmount,
          };
        });
        setInvoiceMaps(prev => ({ ...prev, ...restoredInvoiceMaps }));
        const savedInvoiceSwaps = session.invoiceSwaps ?? {};
        setInvoiceSwaps(prev => ({
          ...prev,
          ...savedInvoiceSwaps,
          bank: savedInvoiceSwaps.bank ?? true,
          wallet: savedInvoiceSwaps.wallet ?? true,
          visa: savedInvoiceSwaps.visa ?? true,
          jawwal: savedInvoiceSwaps.jawwal ?? true,
        }));
        setInvoiceRemoveRepeatedHeaders(prev => ({ ...prev, ...(session.invoiceRemoveRepeatedHeaders ?? {}) }));
        setCustomInvoicePlatforms(session.customInvoicePlatforms ?? []);
        setNewInvoicePlatformName("");
      }
      stageADataLoaded.current = true;
    });
  }, []);
  useEffect(() => {
    if (!stageADataLoaded.current) return;
    const timer = window.setTimeout(() => {
      void storageSet("stage_a_session", {
        results, summary: databaseSummary, databaseFileName, databaseHeaders, databaseRows, databaseMap, removeRepeatedHeaders,
        invoiceFileNames, invoiceFileCounts, invoiceFileSnapshots, invoiceHeaders, invoiceRows, invoiceMaps, invoiceSwaps, invoiceRemoveRepeatedHeaders, customInvoicePlatforms,
      });
    }, 500);
    return () => window.clearTimeout(timer);
  }, [results, databaseSummary, databaseFileName, databaseHeaders, databaseRows, databaseMap, removeRepeatedHeaders,
    invoiceFileNames, invoiceFileCounts, invoiceFileSnapshots, invoiceHeaders, invoiceRows, invoiceMaps, invoiceSwaps, invoiceRemoveRepeatedHeaders, customInvoicePlatforms]);

  const loadDatabase = async (file: File) => {
    setDatabaseFile(file); setDatabaseFileName(file.name); setError(null);
    try {
      await saveStageASourceFiles("stageA", "database", [file]);
      const parsed = parseSheet(await readFileBuf(file));
      setDatabaseHeaders(parsed.headers); setDatabaseRows(parsed.rows);
      setDatabaseMap({
        name: autoDetect(parsed.headers, ["اسم الزبون", "اسم العميل", "customer", "client", "name", "اسم"]),
        amount: autoDetect(parsed.headers, ["مدفوع بالكارد", "مدفوع بالكرت", "card paid", "card payment"]) || autoDetect(parsed.headers, HINTS.debit) || autoDetect(parsed.headers, ["amount", "مبلغ"]),
        invoiceTotal: autoDetect(parsed.headers, ["مبلغ الفاتورة", "اجمالي الفاتورة", "إجمالي الفاتورة", "invoice total", "total"]),
        cashAmount: autoDetect(parsed.headers, ["مدفوع نقدا", "مدفوع نقداً", "نقدا", "cash paid", "cash"]),
        checkAmount: autoDetect(parsed.headers, ["مدفوع بشيك", "قيمة الشيك", "الشيك", "check paid", "check"]),
        debtAmount: autoDetect(parsed.headers, ["مبلغ الحساب", "المتبقي", "الدين", "الرصيد المستحق", "balance due", "due"]),
        userId: autoDetect(parsed.headers, ["user id", "userid", "user number", "رقم المستخدم", "معرف المستخدم", "رقم الكاشير", "المستخدم"]),
        date: autoDetect(parsed.headers, ["التاريخ", "تاريخ الفاتورة", "invoice date", "date"]),
        registrationTime: autoDetect(parsed.headers, ["الساعة", "وقت التسجيل", "وقت الفاتورة", "time", "timestamp"]),
        customerNumber: autoDetect(parsed.headers, ["رقم الزبون", "رقم العميل", "customer number", "customer id", "client id"]),
        invoiceNumber: autoDetect(parsed.headers, ["رقم الفاتورة", "رقم الفاتوره", "invoice number", "invoice no", "invoice id"])
      });
    } catch (e) { setError((e as Error).message); }
  };

  const loadInvoices = async (source: string, files: File[], append = false) => {
    if (!files.length) return;
    setInvoiceFiles(previous => ({ ...previous, [source]: files[0] }));
    setInvoiceProcessing(previous => ({ ...previous, [source]: true }));
    setError(`جارٍ تحليل ${files.length} ملف${files.length > 1 ? "ات" : ""} واستخراج الحركات وتنظيف الأرصدة والترويسات...`);
    try {
      await saveStageASourceFiles("stageA", source, files, append);
      const previousHeaders = append ? invoiceHeaders[source] ?? [] : [];
      const previousRows = append ? invoiceRows[source] ?? [] : [];
      const parsed = await parseInvoiceFiles(files, previousHeaders);
      if (!append && previousRows.length && invoiceFileNames[source]) {
        const snapshot: StageAInvoiceFileSnapshot = {
          fileName: invoiceFileNames[source],
          fileCount: invoiceFileCounts[source] ?? 1,
          headers: previousHeaders,
          rows: previousRows,
          map: invoiceMaps[source] ?? { ...invoiceMapDefaults },
          swap: invoiceSwaps[source] ?? true,
        };
        setInvoiceFileSnapshots(previous => ({ ...previous, [source]: [snapshot] }));
      }
      const fileName = append && invoiceFileNames[source] ? `${invoiceFileNames[source]} + ${parsed.fileName}` : parsed.fileName;
      setInvoiceFileNames(previous => ({ ...previous, [source]: fileName }));
      setInvoiceFileCounts(previous => ({ ...previous, [source]: (append ? previous[source] ?? 0 : 0) + files.length }));
      setInvoiceHeaders(previous => ({ ...previous, [source]: parsed.headers }));
      setInvoiceRows(previous => ({ ...previous, [source]: [...previousRows, ...parsed.rows] }));
      const invoiceColumns = detectInvoiceFileColumns(parsed.headers);
      setInvoiceMaps(previous => ({ ...previous, [source]: {
        ...previous[source],
        ...invoiceColumns,
        paidAmount: invoiceColumns.paidAmount || autoDetect(parsed.headers, ["amount", "مبلغ"]),
        accountType: autoDetect(parsed.headers, HINTS.accountType) || previous[source]?.accountType || "",
        userId: autoDetect(parsed.headers, ["user id", "userid", "user number", "رقم المستخدم", "معرف المستخدم", "رقم الكاشير", "المستخدم"]) || previous[source]?.userId || "",
        date: autoDetect(parsed.headers, HINTS.date) || previous[source]?.date || "",
        registrationTime: autoDetect(parsed.headers, ["الساعة", "وقت التسجيل", "وقت الفاتورة", "time", "timestamp"]) || previous[source]?.registrationTime || "",
        reference: autoDetect(parsed.headers, HINTS.authNum) || previous[source]?.reference || "",
      } }));
      setInvoiceSwaps(previous => ({ ...previous, [source]: true }));
      setError(`تمت معالجة ${files.length} ملف${files.length > 1 ? "ات" : ""}: اكتشاف صف العناوين وتنظيف الترويسات والأرصدة وتجهيز ${previousRows.length + parsed.rows.length} صفاً.`);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setInvoiceProcessing(previous => ({ ...previous, [source]: false }));
    }
  };

  const processDatabaseHeaders = () => {
    if (!removeRepeatedHeaders) return;
    const headerKeys = databaseHeaders.map(header => normName(header));
    const cleanedRows = databaseRows.filter(row => !databaseHeaders.every((header, index) =>
      normName(String(row[header] ?? "")) === headerKeys[index]
    ));
    setDatabaseRows(cleanedRows);
    setError(cleanedRows.length === databaseRows.length ? "لم يتم العثور على ترويسات مكررة لإزالتها." : `تمت إزالة ${databaseRows.length - cleanedRows.length} صفوف ترويسة مكررة.`);
  };

  const processInvoiceHeaders = (source: string) => {
    const headers = invoiceHeaders[source] ?? [];
    const sourceRows = invoiceRows[source] ?? [];
    const headerKeys = headers.map(header => normName(header));
    const withoutRepeatedHeaders = invoiceRemoveRepeatedHeaders[source] === false
      ? sourceRows
      : sourceRows.filter(row => !headers.every((header, index) => normName(String(row[header] ?? "")) === headerKeys[index]));
    const cleaned = removeStatementBalances(headers, withoutRepeatedHeaders.map(row => Object.fromEntries(
      Object.entries(row).map(([header, value]) => [header, cleanStatementText(value as StatementCell)])
    )));
    setInvoiceHeaders(previous => ({ ...previous, [source]: cleaned.headers }));
    setInvoiceRows(previous => ({ ...previous, [source]: cleaned.rows }));
    setError(`اكتملت معالجة ${invoiceFileNames[source] || source}: إزالة ${sourceRows.length - withoutRepeatedHeaders.length} ترويسة مكررة و${withoutRepeatedHeaders.length - cleaned.rows.length} صف رصيد/غير حركي، واستبقاء ${cleaned.rows.length} حركة.`);
  };
  const exportInvoiceSheet = (source: string) => {
    const headers = invoiceHeaders[source] ?? [];
    const rows = invoiceRows[source] ?? [];
    if (!headers.length || !rows.length) {
      setError("لا توجد حركات معالجة لتصديرها.");
      return;
    }
    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows.map(row => headers.map(header => row[header] ?? ""))]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "كشف منظف");
    XLSX.writeFile(workbook, `${(invoiceFileNames[source] || source).replace(/\.[^.]+/g, "").replace(/[\\/]+/g, "-")}-منظف.xlsx`);
  };
  useEffect(() => {
    if (!initialFile || initialFile.stage !== "stageA") return;
    if (initialFile.source === "database") void loadDatabase(initialFile.file);
    else void loadInvoices(initialFile.source, [initialFile.file]);
  }, [initialFile]);
  const restoreInvoiceSnapshot = (source: string, snapshot: StageAInvoiceFileSnapshot) => {
    setInvoiceFiles(previous => ({ ...previous, [source]: null }));
    setInvoiceFileNames(previous => ({ ...previous, [source]: snapshot.fileName }));
    setInvoiceFileCounts(previous => ({ ...previous, [source]: snapshot.fileCount }));
    setInvoiceHeaders(previous => ({ ...previous, [source]: snapshot.headers }));
    setInvoiceRows(previous => ({ ...previous, [source]: snapshot.rows }));
    setInvoiceMaps(previous => ({ ...previous, [source]: snapshot.map }));
    setInvoiceSwaps(previous => ({ ...previous, [source]: snapshot.swap }));
    setShowInvoiceSnapshots(previous => ({ ...previous, [source]: false }));
    setError(`تم استرجاع ${snapshot.fileName} (${snapshot.rows.length} حركة).`);
  };

  const clearStageASession = () => {
    if (!window.confirm("مسح جميع ملفات ونتائج المرحلة A؟ المشاريع المحفوظة لن تتأثر.")) return;
    setDatabaseFile(null);
    setDatabaseFileName("");
    setDatabaseHeaders([]);
    setDatabaseRows([]);
    setDatabaseMap({ name: "", amount: "", invoiceTotal: "", cashAmount: "", checkAmount: "", debtAmount: "", userId: "", date: "", registrationTime: "", customerNumber: "", invoiceNumber: "" });
    setInvoiceFiles({ bank: null, wallet: null, visa: null, jawwal: null });
    setInvoiceFileNames({ bank: "", wallet: "", visa: "", jawwal: "" });
    setInvoiceFileCounts({ bank: 0, wallet: 0, visa: 0, jawwal: 0 });
    setInvoiceFileSnapshots({});
    setShowInvoiceSnapshots({});
    setInvoiceHeaders({ bank: [], wallet: [], visa: [], jawwal: [] });
    setInvoiceRows({ bank: [], wallet: [], visa: [], jawwal: [] });
    setInvoiceMaps({ bank: { ...invoiceMapDefaults }, wallet: { ...invoiceMapDefaults }, visa: { ...invoiceMapDefaults }, jawwal: { ...invoiceMapDefaults } });
    setInvoiceSwaps({ bank: true, wallet: true, visa: true, jawwal: true });
    setInvoiceRemoveRepeatedHeaders({ bank: true, wallet: true, visa: true, jawwal: true });
    setCustomInvoicePlatforms([]);
    setResults([]);
    setDatabaseSummary({ databaseInvoiceCount: 0, databaseInvoiceTotal: 0, databaseCashTotal: 0, databaseDebtTotal: 0, databaseCardTotal: 0, databaseDiscountTotal: 0, databaseInvoices: [] });
    setSelectedResultIds(new Set());
    setError(null);
    void storageSet("stage_a_session", null);
    void clearStageASourceFiles("stageA").catch(cause => console.error("تعذر حذف ملفات الجلسة المؤقتة للمرحلة A:", cause));
  };

  const runStageA = () => {
    if (!databaseRows.length || !databaseMap.name || !databaseMap.amount) { setError("ارفع ملف قاعدة البيانات وحدد عمود الاسم والمبلغ أولاً"); return; }
    const available = databaseRows.map((row, index) => ({ row, index, used: false }));
    const parseStageAAmount = (value: unknown) => Number.parseFloat(String(value ?? "").replace(/[٫٬]/g, match => match === "٫" ? "." : "").replace(/,/g, "")) || 0;
    const output: StageAResult[] = [];
    const databaseInvoices: StageADatabaseInvoice[] = databaseRows.map(row => {
      const invoiceTotal = parseStageAAmount(row[databaseMap.invoiceTotal]);
      const cashAmount = parseStageAAmount(row[databaseMap.cashAmount]);
      const checkAmount = parseStageAAmount(row[databaseMap.checkAmount]);
      const cardAmount = parseStageAAmount(row[databaseMap.amount]);
      const debtAmount = parseStageAAmount(row[databaseMap.debtAmount]);
      const discountAmount = Math.max(0, Math.round((invoiceTotal - cashAmount - checkAmount - cardAmount - debtAmount) * 100) / 100);
      return {
        invoiceNumber: String(row[databaseMap.invoiceNumber] ?? ""),
        invoiceDate: stageAEnglishDate(row[databaseMap.date]),
        registrationTime: String(row[databaseMap.registrationTime] ?? ""),
        name: String(row[databaseMap.name] ?? ""),
        customerNumber: String(row[databaseMap.customerNumber] ?? ""),
        userId: String(row[databaseMap.userId] ?? ""),
        invoiceTotal, cashAmount, checkAmount, cardAmount, debtAmount, discountAmount,
      };
    });
    const sources = [
      { key: "bank", label: "بنك فلسطين" },
      { key: "wallet", label: "بال بي" },
      { key: "visa", label: "فيزا" },
      { key: "jawwal", label: "جوال بي" },
      ...customInvoicePlatforms.map(platform => ({ key: platform.id, label: platform.name })),
    ];
    sources.forEach(({ key: source, label: sourceLabel }) => {
      const map = invoiceMaps[source];
      (invoiceRows[source] || []).forEach((invoice, index) => {
        const originalName = String(map.name ? invoice[map.name] : "").trim();
        const name = cleanStageAName(originalName);
        const paidAmount = parseStageAAmount(map.paidAmount ? invoice[map.paidAmount] : 0);
        const receivedAmount = parseStageAAmount(map.receivedAmount ? invoice[map.receivedAmount] : 0);
        const amount = invoiceSwaps[source] ? receivedAmount || paidAmount : paidAmount || receivedAmount;
        const isReceived = invoiceSwaps[source] ? receivedAmount > 0 || paidAmount === 0 : paidAmount > 0 || receivedAmount === 0;
        if (!name && !amount) return;
        const sameName = available.filter(item => !item.used && stageANameKey(String(item.row[databaseMap.name])) === stageANameKey(name));
        const match = sameName.find(item => Math.abs(parseStageAAmount(item.row[databaseMap.amount]) - amount) <= 0.01);
        if (match) match.used = true;
        // لا نعرض سجل قاعدة البيانات كأنه تطابق إلا إذا تطابق الاسم والمبلغ فعلاً.
        // المبالغ الأخرى لنفس الاسم تظهر في رسالة المشكلة فقط.
        const databaseCandidate = match;
        const issue = match ? "" : sameName.length ? `الاسم موجود، لكن المبلغ مختلف (المطلوب: ${fmtNum(amount)}؛ المتاح: ${sameName.map(item => fmtNum(parseStageAAmount(item.row[databaseMap.amount]))).join("، ")})` : "الاسم غير موجود في قاعدة البيانات";
        const classification = classifyStageAInvoice(originalName, sourceLabel);
        output.push({
          id: index, source: sourceLabel, name, originalName, amount,
          databaseName: databaseCandidate ? String(databaseCandidate.row[databaseMap.name] ?? "") : "",
          databaseAmount: databaseCandidate ? parseStageAAmount(databaseCandidate.row[databaseMap.amount]) : 0,
          invoiceTotal: databaseCandidate ? parseStageAAmount(databaseCandidate.row[databaseMap.invoiceTotal]) : 0,
          cashAmount: databaseCandidate ? parseStageAAmount(databaseCandidate.row[databaseMap.cashAmount]) : 0,
          checkAmount: databaseCandidate ? parseStageAAmount(databaseCandidate.row[databaseMap.checkAmount]) : 0,
          debtAmount: databaseCandidate ? parseStageAAmount(databaseCandidate.row[databaseMap.debtAmount]) : 0,
          discountAmount: databaseCandidate ? Math.max(0, Math.round((
            parseStageAAmount(databaseCandidate.row[databaseMap.invoiceTotal]) -
            parseStageAAmount(databaseCandidate.row[databaseMap.cashAmount]) -
            parseStageAAmount(databaseCandidate.row[databaseMap.checkAmount]) -
            parseStageAAmount(databaseCandidate.row[databaseMap.amount]) -
            parseStageAAmount(databaseCandidate.row[databaseMap.debtAmount])
          ) * 100) / 100) : 0,
          type: isReceived ? "مستلم" : "مدفوع",
          userId: databaseCandidate ? String(databaseCandidate.row[databaseMap.userId] ?? "") : String(map.userId ? invoice[map.userId] ?? "" : ""),
          invoiceDate: databaseCandidate ? stageAEnglishDate(databaseCandidate.row[databaseMap.date]) : stageAEnglishDate(map.date ? invoice[map.date] : ""),
          registrationTime: databaseCandidate ? String(databaseCandidate.row[databaseMap.registrationTime] ?? "") : String(map.registrationTime ? invoice[map.registrationTime] ?? "" : ""),
          customerNumber: databaseCandidate ? String(databaseCandidate.row[databaseMap.customerNumber] ?? "") : "",
          invoiceNumber: databaseCandidate ? String(databaseCandidate.row[databaseMap.invoiceNumber] ?? "") : String(map.reference ? invoice[map.reference] ?? "" : ""),
          accountType: String(map.accountType ? invoice[map.accountType] : "").trim(), invoice, matched: !!match, issue, ...classification
        });
      });
    });
    const summary: StageASummary = databaseInvoices.reduce<StageASummary>((current, invoice) => ({
      databaseInvoiceCount: current.databaseInvoiceCount + 1,
      databaseInvoiceTotal: current.databaseInvoiceTotal + invoice.invoiceTotal,
      databaseCashTotal: current.databaseCashTotal + invoice.cashAmount,
      databaseDebtTotal: current.databaseDebtTotal + invoice.debtAmount,
      databaseCardTotal: current.databaseCardTotal + invoice.cardAmount,
      databaseDiscountTotal: (current.databaseDiscountTotal ?? 0) + invoice.discountAmount,
    }), { databaseInvoiceCount: 0, databaseInvoiceTotal: 0, databaseCashTotal: 0, databaseDebtTotal: 0, databaseCardTotal: 0, databaseDiscountTotal: 0 });
    summary.databaseInvoices = databaseInvoices;
    setDatabaseSummary(summary);
    setResults(output);
    void storageSet("stage_a_session", {
      results: output, summary, databaseFileName, databaseHeaders, databaseRows, databaseMap, removeRepeatedHeaders,
      invoiceFileNames, invoiceHeaders, invoiceRows, invoiceMaps, invoiceSwaps, invoiceRemoveRepeatedHeaders, customInvoicePlatforms,
    });
    setError(null);
  };

  const exportStageA = () => {
    if (!results.length) return;
    const workbook = XLSX.utils.book_new();
    const rows = results.map(result => [result.source, result.category, result.originalName, result.databaseName, result.amount, result.type, result.databaseAmount || "", result.invoiceTotal || "", result.cashAmount || "", result.checkAmount || "", result.debtAmount || "", result.discountAmount || "", result.userId, result.invoiceDate, result.registrationTime, result.customerNumber, result.invoiceNumber, result.accountType, result.matched ? "مطابق" : "غير مطابق", result.categoryIssue || result.issue]);
    const worksheet = XLSX.utils.aoa_to_sheet([["المنصة", "التصنيف", "بيان الفاتورة", "اسم الزبون في قاعدة البيانات", "مبلغ الحركة", "النوع", "مدفوع بالتطبيق/الكارد", "إجمالي الفاتورة", "مدفوع نقداً", "مدفوع بشيك", "المتبقي على الزبون", "الخصم", "رقم المستخدم", "تاريخ الفاتورة", "الساعة", "رقم الزبون", "رقم الفاتورة", "نوع الحساب", "حالة المطابقة", "المشكلة / التصحيح المطلوب"], ...rows]);
    worksheet["!cols"] = [{ wch: 18 }, { wch: 14 }, { wch: 32 }, { wch: 28 }, { wch: 14 }, { wch: 12 }, { wch: 18 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 18 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 18 }, { wch: 14 }, { wch: 48 }];
    XLSX.utils.book_append_sheet(workbook, worksheet, "نتائج المرحلة A");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ["ملخص اليوم من قاعدة البيانات", "القيمة"],
      ["عدد الفواتير", databaseSummary.databaseInvoiceCount],
      ["إجمالي الفواتير", databaseSummary.databaseInvoiceTotal],
      ["المبلغ النقدي", databaseSummary.databaseCashTotal],
      ["المدفوع بالشيك", (databaseSummary.databaseInvoices ?? []).reduce((sum, invoice) => sum + invoice.checkAmount, 0)],
      ["مجموع الخصومات", databaseSummary.databaseDiscountTotal ?? 0],
      ["المتبقي على الزبائن", databaseSummary.databaseDebtTotal],
      ["المدفوع بالتطبيقات/البطاقات", databaseSummary.databaseCardTotal],
    ]), "ملخص اليوم");
    const databaseInvoiceRows = databaseSummary.databaseInvoices ?? [];
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ["رقم الفاتورة", "التاريخ", "الوقت", "اسم الزبون", "رقم الزبون", "رقم المستخدم", "إجمالي الفاتورة", "نقدي", "شيك", "تطبيق/بطاقة", "دين", "الخصم"],
      ...databaseInvoiceRows.map(invoice => [invoice.invoiceNumber, invoice.invoiceDate, invoice.registrationTime, invoice.name, invoice.customerNumber, invoice.userId, invoice.invoiceTotal, invoice.cashAmount, invoice.checkAmount, invoice.cardAmount, invoice.debtAmount, invoice.discountAmount]),
    ]), "كل فواتير قاعدة البيانات");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ["رقم الفاتورة", "التاريخ", "الوقت", "اسم الزبون", "رقم الزبون", "رقم المستخدم", "إجمالي الفاتورة", "نقدي", "شيك", "تطبيق/بطاقة", "دين", "الخصم"],
      ...databaseInvoiceRows.filter(invoice => invoice.discountAmount > 0).map(invoice => [invoice.invoiceNumber, invoice.invoiceDate, invoice.registrationTime, invoice.name, invoice.customerNumber, invoice.userId, invoice.invoiceTotal, invoice.cashAmount, invoice.checkAmount, invoice.cardAmount, invoice.debtAmount, invoice.discountAmount]),
    ]), "فواتير فيها خصم");
    XLSX.writeFile(workbook, "نتائج_المرحلة_A_ومشاكل_اخرى.xlsx");
  };

  const resultKey = (result: StageAResult) => `${result.source}-${result.id}`;
  const toggleResultSelection = (result: StageAResult) => setSelectedResultIds(previous => {
    const next = new Set(previous); const key = resultKey(result);
    if (next.has(key)) next.delete(key); else next.add(key);
    return next;
  });
  const updateResultStatus = (result: StageAResult, matched: boolean) => {
    setResults(previous => previous.map(item => item === result ? { ...item, matched, issue: matched ? "تم اعتمادها يدوياً" : "تم رفضها يدوياً" } : item));
    if (matched) onSafeActivities?.([`${result.originalName} ↔ ${result.databaseName} · ${fmtNum(result.amount)} ₪`]);
  };
  const rejectSelectedResults = () => {
    setResults(previous => previous.map(item => selectedResultIds.has(resultKey(item)) ? { ...item, matched: false, issue: "تم رفضها يدوياً" } : item));
    setSelectedResultIds(new Set());
  };

  const allActionResults = results.filter(result => {
    const nameQuery = resultNameSearch.trim().toLowerCase();
    const amountQuery = resultAmountSearch.trim();
    return (!nameQuery || `${result.originalName} ${result.databaseName} ${JSON.stringify(result.invoice)}`.toLowerCase().includes(nameQuery)) &&
      (stageASourceFilter === "all" || result.source === stageASourceFilter || result.category === stageASourceFilter) &&
      (resultFilter === "all" || (resultFilter === "matched" ? result.matched : resultFilter === "unmatched" ? !result.matched : resultFilter === "wallet" ? /محفظة|بال بي/.test(result.category) : result.category === resultFilter)) &&
      (!resultMinAmount || result.amount >= Number(resultMinAmount)) &&
      (!resultMaxAmount || result.amount <= Number(resultMaxAmount)) &&
      (!amountQuery || String(result.amount).includes(amountQuery) || String(result.databaseAmount).includes(amountQuery));
  });
  const pagedActionResults = allActionResults.slice((Math.min(stageAPage, Math.max(1, Math.ceil(allActionResults.length / 50))) - 1) * 50, Math.min(stageAPage, Math.max(1, Math.ceil(allActionResults.length / 50))) * 50);
    const stageAFilterBar = results.length > 0 && <div className="flex flex-row-reverse flex-wrap items-center gap-3 rounded-xl border border-slate-200 bg-white p-3"><input value={resultNameSearch} onChange={event => { setResultNameSearch(event.target.value); setStageAPage(1); }} placeholder="بحث برقم الفاتورة أو العميل أو المرجع..." className="order-1 min-w-[240px] flex-1 rounded-lg border border-slate-200 px-4 py-3 text-sm outline-none focus:border-blue-400 focus:ring-2 focus:ring-blue-100" /><select value={stageASourceFilter} onChange={event => { setStageASourceFilter(event.target.value); setStageAPage(1); }} className="order-2 rounded-lg border border-slate-200 bg-white px-4 py-3 text-sm outline-none"><option value="all">كل الوسائل</option><option value="بنك فلسطين">بنك فلسطين</option><option value="بال بي">بال بي</option><option value="محفظة شخصية">محفظة شخصية</option><option value="جوال بي">جوال بي</option><option value="فيزا">فيزا</option>{customInvoicePlatforms.map(platform => <option key={platform.id} value={platform.name}>{platform.name}</option>)}</select><details className="order-3 relative"><summary className="cursor-pointer list-none rounded-lg border px-3 py-2 text-sm">{resultMinAmount || resultMaxAmount ? "نطاق المبلغ · مفعّل" : "نطاق المبلغ"}</summary><div className="absolute left-0 z-30 mt-1 flex w-64 gap-2 rounded-lg border bg-white p-2 shadow-lg"><input type="number" min="0" value={resultMinAmount} onChange={event => { setResultMinAmount(event.target.value); setStageAPage(1); }} placeholder="من مبلغ" className="w-1/2 rounded border px-2 py-1.5 text-xs" /><input type="number" min="0" value={resultMaxAmount} onChange={event => { setResultMaxAmount(event.target.value); setStageAPage(1); }} placeholder="إلى مبلغ" className="w-1/2 rounded border px-2 py-1.5 text-xs" /></div></details><button onClick={() => setSelectedResultIds(new Set(visibleActionResults.map(result => resultKey(result))))} className="order-4 rounded-lg border border-slate-200 bg-white px-5 py-3 text-sm font-medium hover:bg-slate-50">تحديد الصفحة</button></div>;
    const inlineResultActions = results.length > 0 && <div className="space-y-2 rounded-xl border border-slate-200 bg-white p-3"><div className="flex items-center justify-between"><strong className="text-xs">خيارات الفواتير</strong><span className="text-xs text-slate-500">حدد الفاتورة لإظهار خياراتها</span></div>{pagedActionResults.map(result => <div key={`inline-${resultKey(result)}`} className="flex flex-wrap items-center gap-2 rounded-lg border border-slate-100 bg-slate-50 p-2"><input type="checkbox" checked={selectedResultIds.has(resultKey(result))} onChange={() => toggleResultSelection(result)} /><span className="min-w-48 flex-1 text-xs font-medium">{result.originalName}<span className="mr-2 text-slate-500">{fmtNum(result.amount)}</span></span>{selectedResultIds.has(resultKey(result)) && <><button onClick={() => updateResultStatus(result, true)} className="rounded border border-green-200 bg-green-50 px-2 py-1 text-[10px] text-green-700">اعتماد</button><button onClick={() => updateResultStatus(result, false)} className="rounded border border-red-200 bg-red-50 px-2 py-1 text-[10px] text-red-700">رفض</button><select onChange={event => replaceMatchedDatabase(result, event.target.value)} defaultValue="-1" className="max-w-40 rounded border bg-white px-2 py-1 text-[10px]"><option value="-1">تغيير إلى شخص آخر</option>{databaseRows.map((row, index) => <option key={index} value={index}>{String(row[databaseMap.name] ?? "")} · {String(row[databaseMap.amount] ?? "")}</option>)}</select><button onClick={() => { if (window.confirm("حذف الفاتورة وسجلها من النتائج؟")) deleteMatchedPair(result); }} className="rounded border border-slate-300 bg-slate-100 px-2 py-1 text-[10px] text-slate-700">حذف الاثنين</button></>}</div>)}</div>;
  const stageAPageSize = 20;
  const totalStageAPages = Math.max(1, Math.ceil(allActionResults.length / stageAPageSize));
  const currentStageAPage = Math.min(stageAPage, totalStageAPages);
  const visibleActionResults = allActionResults.slice((currentStageAPage - 1) * stageAPageSize, currentStageAPage * stageAPageSize);
  const paginationControls = results.length > stageAPageSize && <div className="flex items-center justify-center gap-2 rounded-xl border border-slate-200 bg-white p-3 text-xs"><button onClick={() => setStageAPage(page => Math.max(1, page - 1))} disabled={currentStageAPage === 1} className="rounded border px-3 py-1.5 disabled:opacity-40">السابق</button><span className="text-slate-500">صفحة {currentStageAPage} من {totalStageAPages} · عرض {visibleActionResults.length} من {allActionResults.length}</span><button onClick={() => setStageAPage(page => Math.min(totalStageAPages, page + 1))} disabled={currentStageAPage === totalStageAPages} className="rounded border px-3 py-1.5 disabled:opacity-40">التالي</button></div>;
  const visibleMatchedResults = results.filter(result => {
    if (!result.matched) return false;
    const query = matchedSearch.trim().toLowerCase();
    if (!query) return true;
    return `${result.originalName} ${result.databaseName}`.toLowerCase().includes(query) || String(result.amount).includes(query) || String(result.databaseAmount).includes(query);
  });
  const rejectMatched = (result: StageAResult) => updateResultStatus(result, false);
  const deleteMatchedPair = (result: StageAResult) => setResults(previous => previous.filter(item => item !== result));
  const replaceMatchedDatabase = (result: StageAResult, value: string) => {
    const row = databaseRows[Number(value)];
    if (!row) return;
    const invoiceTotal = Number.parseFloat(String(row[databaseMap.invoiceTotal] ?? "").replace(/,/g, "")) || 0;
    const cashAmount = Number.parseFloat(String(row[databaseMap.cashAmount] ?? "").replace(/,/g, "")) || 0;
    const checkAmount = Number.parseFloat(String(row[databaseMap.checkAmount] ?? "").replace(/,/g, "")) || 0;
    const cardAmount = Number.parseFloat(String(row[databaseMap.amount] ?? "").replace(/,/g, "")) || 0;
    const debtAmount = Number.parseFloat(String(row[databaseMap.debtAmount] ?? "").replace(/,/g, "")) || 0;
    setResults(previous => previous.map(item => item === result ? {
      ...item,
      invoiceTotal, cashAmount, checkAmount, debtAmount,
      discountAmount: Math.max(0, Math.round((invoiceTotal - cashAmount - checkAmount - cardAmount - debtAmount) * 100) / 100),
      databaseName: String(row[databaseMap.name] ?? ""),
      databaseAmount: Number.parseFloat(String(row[databaseMap.amount] ?? "").replace(/,/g, "")) || 0,
      userId: String(row[databaseMap.userId] ?? ""),
      invoiceDate: stageAEnglishDate(row[databaseMap.date]),
      registrationTime: String(row[databaseMap.registrationTime] ?? ""),
      customerNumber: String(row[databaseMap.customerNumber] ?? ""),
      invoiceNumber: String(row[databaseMap.invoiceNumber] ?? ""),
      matched: true, issue: "تم تعديل سجل قاعدة البيانات يدوياً"
    } : item));
  };

  const replacementRows = databaseRows.filter(row => {
    const query = replacementSearch.trim().toLowerCase();
    return !query || `${String(row[databaseMap.name] ?? "")} ${String(row[databaseMap.amount] ?? "")}`.toLowerCase().includes(query);
  });

  const matchedActionsPanel = results.length > 0 && <div className="rounded-xl border border-green-200 bg-green-50/40 p-3"><div className="mb-2 flex flex-wrap items-center gap-2"><strong className="text-xs text-green-800">المطابقات وإجراءاتها</strong><input value={matchedSearch} onChange={event => setMatchedSearch(event.target.value)} placeholder="بحث المطابقات بالاسم أو المبلغ" className="min-w-56 flex-1 rounded border bg-white px-2 py-1.5 text-[11px]" /><span className="text-xs text-green-700">{visibleMatchedResults.length} مطابقة</span></div><div className="space-y-2">{visibleMatchedResults.map(result => <div key={`matched-${resultKey(result)}`} className="flex flex-wrap items-center gap-2 rounded-lg border border-green-100 bg-white p-2"><input type="checkbox" checked={selectedMatchedIds.has(resultKey(result))} onChange={() => setSelectedMatchedIds(previous => { const next = new Set(previous); const key = resultKey(result); if (next.has(key)) next.delete(key); else next.add(key); return next; })} /><span className="min-w-44 flex-1 text-xs font-medium">{result.originalName}<span className="mr-2 text-slate-500">{fmtNum(result.amount)}</span></span><span className="text-[11px] text-green-700">↔ {result.databaseName} ({fmtNum(result.databaseAmount)})</span><select value={databaseRows.findIndex(row => String(row[databaseMap.name] ?? "") === result.databaseName && Number.parseFloat(String(row[databaseMap.amount] ?? "").replace(/,/g, "")) === result.databaseAmount)} onChange={event => replaceMatchedDatabase(result, event.target.value)} className="max-w-36 rounded border px-2 py-1 text-[10px]"><option value="-1">اختيار شخص آخر</option>{databaseRows.map((row, index) => <option key={index} value={index}>{String(row[databaseMap.name] ?? "")} · {String(row[databaseMap.amount] ?? "")}</option>)}</select><button onClick={() => rejectMatched(result)} className="rounded border border-red-200 bg-red-50 px-2 py-1 text-[10px] text-red-700">✕ رفض</button><button onClick={() => deleteMatchedPair(result)} className="rounded border border-slate-300 bg-slate-100 px-2 py-1 text-[10px] text-slate-700">حذف الاثنين</button></div>)}</div></div>;

  const selectionPanel = results.length > 0 && <div className="rounded-xl border border-slate-200 bg-white p-3"><div className="mb-2 flex flex-wrap items-center gap-2"><strong className="text-xs">إجراءات الفواتير</strong><span className="text-xs text-slate-500">محدد: {selectedResultIds.size}</span><input value={resultNameSearch} onChange={event => setResultNameSearch(event.target.value)} placeholder="بحث بالاسم أو البيان" className="min-w-48 flex-1 rounded border px-2 py-1.5 text-[11px]" /><input value={resultAmountSearch} onChange={event => setResultAmountSearch(event.target.value)} placeholder="بحث بالمبلغ" inputMode="decimal" className="w-28 rounded border px-2 py-1.5 text-[11px]" /></div><div className="space-y-2">{visibleActionResults.map(result => <div key={resultKey(result)} className="flex flex-wrap items-center gap-2 rounded-lg border border-slate-100 bg-slate-50 p-2"><input type="checkbox" checked={selectedResultIds.has(resultKey(result))} onChange={() => toggleResultSelection(result)} /><span className="min-w-48 flex-1 text-xs font-medium">{result.originalName}<span className="mr-2 text-slate-500">{fmtNum(result.amount)}</span></span><button onClick={() => updateResultStatus(result, true)} className="rounded border border-green-200 bg-green-50 px-2 py-1 text-[10px] text-green-700">اعتماد</button><button onClick={() => updateResultStatus(result, false)} className="rounded border border-red-200 bg-red-50 px-2 py-1 text-[10px] text-red-700">رفض</button><button onClick={() => { onTransferToB([result], databaseSummary); setSelectedResultIds(previous => { const next = new Set(previous); next.delete(resultKey(result)); return next; }); }} className="rounded bg-emerald-600 px-2 py-1 text-[10px] text-white">نقل B</button><button onClick={() => { onTransferVisaToC([result]); setSelectedResultIds(previous => { const next = new Set(previous); next.delete(resultKey(result)); return next; }); }} className="rounded bg-amber-600 px-2 py-1 text-[10px] text-white">نقل C</button></div>)}</div></div>;

  const renderInvoice = (source: string, label: string) => (
    <div className="space-y-3 rounded-xl border border-slate-200 bg-white p-4 shadow-[0_4px_12px_rgba(30,64,175,0.08)]">
      <h3 className="text-sm font-bold text-slate-800">{label}</h3>
      {!!invoiceHeaders[source].length && (
        <label className="flex w-fit cursor-pointer items-center gap-2 text-xs text-slate-600">
          <input type="checkbox" checked={invoiceSwaps[source]} onChange={event => setInvoiceSwaps(prev => ({ ...prev, [source]: event.target.checked }))} className="rounded" />
          {invoiceSwaps[source] ? "المستلم مدين والمدفوع دائن" : "المدفوع مدين والمستلم دائن"}
        </label>
      )}
      <DropZone file={invoiceFiles[source] ?? null} multiple rowCount={invoiceRows[source]?.length ?? 0} onFile={file => void loadInvoices(source, [file])} onFiles={files => void loadInvoices(source, files)} onClear={() => { setInvoiceFiles(prev => ({ ...prev, [source]: null })); setInvoiceFileNames(prev => ({ ...prev, [source]: "" })); setInvoiceFileCounts(prev => ({ ...prev, [source]: 0 })); setInvoiceHeaders(prev => ({ ...prev, [source]: [] })); setInvoiceRows(prev => ({ ...prev, [source]: [] })); }} />
      {!!invoiceFileNames[source] && <p className="text-xs text-emerald-700">{invoiceFiles[source] ? `${invoiceFileCounts[source] || 1} ملف${(invoiceFileCounts[source] || 1) > 1 ? "ات مدمجة" : ""}: ` : "بيانات الملفات المستعادة: "}{invoiceFileNames[source]}</p>}
      {!!invoiceHeaders[source].length && <>
        <div className="flex flex-wrap items-center gap-3">
          <label className="flex items-center gap-2 text-xs text-slate-600"><input type="checkbox" checked={invoiceRemoveRepeatedHeaders[source] ?? true} onChange={event => setInvoiceRemoveRepeatedHeaders(prev => ({ ...prev, [source]: event.target.checked }))} />إزالة الترويسات المكررة</label>
          <button type="button" onClick={() => processInvoiceHeaders(source)} disabled={invoiceProcessing[source]} className="rounded-lg border border-blue-200 bg-blue-50 px-3 py-2 text-xs font-semibold text-blue-800 hover:bg-blue-100 disabled:opacity-50">معالجة الملف وإزالة الترويسات والأرصدة</button>
          <label className={`cursor-pointer rounded-lg border px-3 py-2 text-xs font-medium hover:bg-slate-50 ${invoiceProcessing[source] ? "pointer-events-none opacity-50" : ""}`}>إضافة ملف للدمج<input type="file" multiple accept=".xlsx,.xls,.csv" disabled={invoiceProcessing[source]} className="hidden" onChange={event => { const files = Array.from(event.currentTarget.files ?? []); if (files.length) void loadInvoices(source, files, true); event.currentTarget.value = ""; }} /></label>
          <button type="button" onClick={() => exportInvoiceSheet(source)} className="rounded-lg bg-emerald-50 px-3 py-2 text-xs font-medium text-emerald-700 hover:bg-emerald-100"><Download className="ml-1 inline h-3.5 w-3.5" />تنزيل Excel منظف</button>
          {!!invoiceFileSnapshots[source]?.length && <div className="relative">
            <button type="button" onClick={() => setShowInvoiceSnapshots(previous => ({ ...previous, [source]: !previous[source] }))} className="rounded-lg border bg-slate-50 px-3 py-2 text-xs font-medium text-slate-700 hover:bg-slate-100"><RotateCcw className="ml-1 inline h-3.5 w-3.5" />استرجاع ملف سابق ({invoiceFileSnapshots[source].length})</button>
            {showInvoiceSnapshots[source] && <>
              <button type="button" aria-label="إغلاق قائمة الملفات السابقة" className="fixed inset-0 z-10 cursor-default" onClick={() => setShowInvoiceSnapshots(previous => ({ ...previous, [source]: false }))} />
              <div className="absolute left-0 z-20 mt-1 max-h-64 w-72 overflow-y-auto rounded-lg border bg-white py-1 shadow-lg">
                {invoiceFileSnapshots[source].map((snapshot, index) => <button key={`${snapshot.fileName}-${index}`} type="button" onClick={() => restoreInvoiceSnapshot(source, snapshot)} className="w-full px-3 py-2 text-right text-xs hover:bg-blue-50">
                  <span className="block truncate font-medium">{snapshot.fileName}</span><span className="mt-1 block text-[10px] text-slate-500">{snapshot.fileCount} ملف · {snapshot.rows.length} حركة</span>
                </button>)}
              </div>
            </>}
          </div>}
        </div>
        <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
          <Sel label="التاريخ" headers={invoiceHeaders[source]} value={invoiceMaps[source].date} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], date: value } }))} />
          <Sel label="الساعة" headers={invoiceHeaders[source]} value={invoiceMaps[source].registrationTime} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], registrationTime: value } }))} />
          <Sel label="البيان / اسم الزبون" headers={invoiceHeaders[source]} value={invoiceMaps[source].name} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], name: value } }))} />
          <Sel label="المبالغ المدفوعة (Credit)" headers={invoiceHeaders[source]} value={invoiceMaps[source].paidAmount} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], paidAmount: value } }))} />
          <Sel label="المبالغ المستلمة (Debit)" headers={invoiceHeaders[source]} value={invoiceMaps[source].receivedAmount} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], receivedAmount: value } }))} />
          <Sel label="نوع الحساب" headers={invoiceHeaders[source]} value={invoiceMaps[source].accountType} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], accountType: value } }))} />
          <Sel label="رقم المستخدم" headers={invoiceHeaders[source]} value={invoiceMaps[source].userId} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], userId: value } }))} />
          <Sel label="المرجع / رقم الفاتورة" headers={invoiceHeaders[source]} value={invoiceMaps[source].reference} onChange={value => setInvoiceMaps(prev => ({ ...prev, [source]: { ...prev[source], reference: value } }))} />
        </div>
      </>}
    </div>
  );

  const requestStageATransfer = (mode: "matched" | "matched-and-selected" | "all") => {
    const candidates = results.filter(result => mode === "all"
      || (mode === "matched" && result.matched)
      || (mode === "matched-and-selected" && (result.matched || selectedResultIds.has(resultKey(result)))));
    setTransferDestinationDialog({
      label: mode === "all" ? "كل نتائج منصة A" : mode === "matched" ? "المطابقات" : "المطابقات والفواتير المحددة",
      results: candidates,
    });
  };
  const databaseInvoiceRows = databaseSummary.databaseInvoices ?? [];
  const databaseInvoicePageSize = 50;
  const databaseInvoicePageCount = Math.max(1, Math.ceil(databaseInvoiceRows.length / databaseInvoicePageSize));
  const visibleDatabaseInvoices = databaseInvoiceRows.slice(
    (Math.min(databaseInvoicePage, databaseInvoicePageCount) - 1) * databaseInvoicePageSize,
    Math.min(databaseInvoicePage, databaseInvoicePageCount) * databaseInvoicePageSize
  );
  const clearAllStageASelection = () => setSelectedResultIds(new Set());
  const selectUnmatchedStageAResults = () => setSelectedResultIds(new Set(results.filter(result => !result.matched).map(result => resultKey(result))));
  const addCustomInvoicePlatform = () => {
    const name = newInvoicePlatformName.trim();
    if (!name) return;
    const id = `custom-${Date.now()}`;
    setCustomInvoicePlatforms(platforms => [...platforms, { id, name }]);
    setInvoiceFiles(files => ({ ...files, [id]: null }));
    setInvoiceFileCounts(counts => ({ ...counts, [id]: 0 }));
    setInvoiceFileSnapshots(snapshots => { const next = { ...snapshots }; delete next[id]; return next; });
    setShowInvoiceSnapshots(open => { const next = { ...open }; delete next[id]; return next; });
    setInvoiceHeaders(headers => ({ ...headers, [id]: [] }));
    setInvoiceRows(rows => ({ ...rows, [id]: [] }));
    setInvoiceMaps(maps => ({ ...maps, [id]: { ...invoiceMapDefaults } }));
    setInvoiceSwaps(swaps => ({ ...swaps, [id]: true }));
    setNewInvoicePlatformName("");
  };

  return <div dir="rtl" className="min-h-screen bg-slate-50 p-6 text-slate-900"><div className="mx-auto max-w-6xl space-y-5">
    <div className="flex flex-wrap items-center justify-between gap-3"><button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-blue-600"><ChevronLeft className="h-4 w-4" />العودة للمنصة الرئيسية</button><div className="flex items-center gap-3"><span className="flex h-12 w-12 items-center justify-center rounded-xl bg-blue-600 text-2xl font-black text-white shadow-sm">A</span><div><h1 className="text-2xl font-black text-blue-700">فرز الفواتير حسب قاعدة البيانات</h1><p className="mt-1 text-sm text-slate-500">مطابقة الاسم والمبلغ مع قاعدة البيانات، ثم استخراج المطابقات ومشاكل أخرى.</p></div></div><div className="flex gap-2"><div className="relative"><button onClick={() => setShowDeleteMenu(open => !open)} className="flex items-center gap-1 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700"><Trash2 className="h-4 w-4" />أدوات الحذف<ChevronDown className="h-3 w-3" /></button>{showDeleteMenu && <><button type="button" aria-label="إغلاق القائمة" className="fixed inset-0 z-10 cursor-default" onClick={() => setShowDeleteMenu(false)} /><div className="absolute left-0 z-20 mt-1 w-56 rounded-lg border bg-white p-1 shadow-lg"><button onClick={() => { setShowDeleteMenu(false); clearStageASession(); }} className="w-full rounded px-3 py-2 text-right text-xs text-red-700 hover:bg-red-50">مسح ملفات ونتائج المرحلة A</button></div></>}</div><button onClick={runStageA} className="flex items-center gap-2 rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700"><Sparkles className="h-4 w-4" />تشغيل الفرز</button>{results.length > 0 && <button onClick={exportStageA} className="flex items-center gap-2 rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white hover:bg-green-700"><Download className="h-4 w-4" />تصدير Excel</button>}</div></div>
    <div className="rounded-xl border border-blue-200 bg-blue-50 p-4 text-sm text-blue-900">تتم المطابقة بالاسم والمبلغ معاً. كل صف في قاعدة البيانات يُستهلك مرة واحدة، لذلك يمكن معالجة فاتورتين للاسم نفسه بمبلغين مختلفين بشكل صحيح. تُزال عبارة <strong>بطاقة ائتمان</strong> والفاصل <strong>/</strong> من بداية البيان، مع إبقاء بقية الاسم والرموز والأرقام.</div>
    {error && <div className="rounded-lg bg-red-100 p-3 text-sm text-red-800">{error}</div>}{stageAFilterBar}
    <div className="space-y-3 rounded-2xl border border-blue-100 bg-white p-5 shadow-sm">
      <div className="flex items-center justify-between"><h2 className="text-base font-bold text-slate-800">ملف قاعدة البيانات</h2><span className="rounded-full bg-blue-50 px-3 py-1 text-xs font-semibold text-blue-700">المصدر الأساسي للفرز</span></div>
      <DropZone file={databaseFile} rowCount={databaseRows.length} onFile={loadDatabase} onClear={() => { setDatabaseFile(null); setDatabaseFileName(""); setDatabaseHeaders([]); setDatabaseRows([]); }} />
      {!databaseFile && databaseFileName && <p className="text-xs text-emerald-700">بيانات الملف المستعادة: {databaseFileName}</p>}
      {!!databaseHeaders.length && <>
        <div className="flex flex-wrap items-center gap-3">
          <label className="flex items-center gap-2 text-xs text-slate-700"><input type="checkbox" checked={removeRepeatedHeaders} onChange={event => setRemoveRepeatedHeaders(event.target.checked)} />إزالة صفوف الترويسة المكررة</label>
          <button type="button" onClick={processDatabaseHeaders} className="rounded-lg border border-blue-200 bg-blue-50 px-3 py-2 text-xs font-semibold text-blue-800 hover:bg-blue-100">معالجة ملف قاعدة البيانات</button>
        </div>
        <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
          <Sel label="اسم الزبون" headers={databaseHeaders} value={databaseMap.name} onChange={value => setDatabaseMap(prev => ({ ...prev, name: value }))} />
          <Sel label="مدفوع بالكارد / تطبيق" headers={databaseHeaders} value={databaseMap.amount} onChange={value => setDatabaseMap(prev => ({ ...prev, amount: value }))} />
          <Sel label="إجمالي مبلغ الفاتورة" headers={databaseHeaders} value={databaseMap.invoiceTotal} onChange={value => setDatabaseMap(prev => ({ ...prev, invoiceTotal: value }))} />
          <Sel label="مدفوع نقداً" headers={databaseHeaders} value={databaseMap.cashAmount} onChange={value => setDatabaseMap(prev => ({ ...prev, cashAmount: value }))} />
          <Sel label="مدفوع بشيك" headers={databaseHeaders} value={databaseMap.checkAmount} onChange={value => setDatabaseMap(prev => ({ ...prev, checkAmount: value }))} />
          <Sel label="المتبقي على الزبون / الدين" headers={databaseHeaders} value={databaseMap.debtAmount} onChange={value => setDatabaseMap(prev => ({ ...prev, debtAmount: value }))} />
          <Sel label="المستخدم / رقم المستخدم" headers={databaseHeaders} value={databaseMap.userId} onChange={value => setDatabaseMap(prev => ({ ...prev, userId: value }))} />
          <Sel label="تاريخ الفاتورة" headers={databaseHeaders} value={databaseMap.date} onChange={value => setDatabaseMap(prev => ({ ...prev, date: value }))} />
          <Sel label="الساعة" headers={databaseHeaders} value={databaseMap.registrationTime} onChange={value => setDatabaseMap(prev => ({ ...prev, registrationTime: value }))} />
          <Sel label="رقم الزبون" headers={databaseHeaders} value={databaseMap.customerNumber} onChange={value => setDatabaseMap(prev => ({ ...prev, customerNumber: value }))} />
          <Sel label="رقم الفاتورة" headers={databaseHeaders} value={databaseMap.invoiceNumber} onChange={value => setDatabaseMap(prev => ({ ...prev, invoiceNumber: value }))} />
        </div>
      </>}
    </div>
    <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">{renderInvoice("bank", "ملف فواتير بنك فلسطين")}{renderInvoice("wallet", "ملف فواتير بال بي")}{renderInvoice("jawwal", "ملف فواتير جوال بي")}{renderInvoice("visa", "ملف فواتير الفيزا")}{customInvoicePlatforms.map(platform => <div key={platform.id}>{renderInvoice(platform.id, `فواتير ${platform.name}`)}<button type="button" onClick={() => setCustomInvoicePlatforms(items => items.filter(item => item.id !== platform.id))} className="mt-1 rounded px-2 py-1 text-xs text-red-700 hover:bg-red-50">حذف المنصة</button></div>)}</div>
    <div className="flex flex-wrap items-center gap-2 rounded-xl border border-slate-200 bg-white p-3"><input value={newInvoicePlatformName} onChange={event => setNewInvoicePlatformName(event.target.value)} placeholder="اسم منصة إضافية" className="min-w-48 flex-1 rounded-lg border px-3 py-2 text-sm" /><button type="button" onClick={addCustomInvoicePlatform} disabled={!newInvoicePlatformName.trim()} className="rounded-lg bg-slate-800 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-900 disabled:opacity-50">إضافة منصة</button></div>
    {paginationControls}
    {!!databaseSummary.databaseInvoiceCount && <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-6">
      {([
        ["عدد الفواتير", databaseSummary.databaseInvoiceCount],
        ["إجمالي الفواتير", databaseSummary.databaseInvoiceTotal],
        ["النقدي", databaseSummary.databaseCashTotal],
        ["الشيك", databaseSummary.databaseInvoices?.reduce((sum, invoice) => sum + invoice.checkAmount, 0) ?? 0],
        ["المتبقي على الزبائن", databaseSummary.databaseDebtTotal],
        ["مدفوع بالكارد / التطبيقات", databaseSummary.databaseCardTotal],
      ] as const).map(([label, value]) => label === "عدد الفواتير"
        ? <button key={label} type="button" onClick={() => { setDatabaseInvoicePage(1); setShowDatabaseInvoices(true); }} className="rounded-xl border border-blue-200 bg-white p-3 text-right hover:border-blue-400 hover:bg-blue-50"><div className="text-xs text-slate-500">{label} · عرض التفاصيل</div><div className="mt-1 font-mono text-sm font-bold">{value}</div></button>
        : <div key={label} className="rounded-xl border border-slate-200 bg-white p-3"><div className="text-xs text-slate-500">{label}</div><div className="mt-1 font-mono text-sm font-bold">{fmtNum(value)}</div></div>)}
    </div>}
    {!!results.length && <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
      <div className="flex flex-wrap items-center justify-between gap-3 border-b bg-slate-50 px-4 py-3 text-sm">
        <div className="flex flex-wrap items-center gap-1"><strong className="ml-3">نتائج الفرز</strong>{([["all", "الكل"], ["matched", "المطابق"], ["unmatched", "غير المطابق"], ["wallet", "محفظة"], ["جوال بي", "جوال بي"], ["visa", "فيزا"]] as const).map(([value, label]) => <button key={value} onClick={() => setResultFilter(value)} className={`rounded-lg px-2.5 py-1.5 text-xs ${resultFilter === value ? "bg-blue-600 text-white" : "bg-white text-slate-600 hover:bg-blue-50"}`}>{label}</button>)}</div>
        <div className="flex flex-wrap items-center gap-2"><span className="text-xs text-slate-500">مطابق: {results.filter(result => result.matched).length} · غير مطابق: {results.filter(result => !result.matched).length} · المحدد: {selectedResultIds.size}</span>
          <button onClick={selectUnmatchedStageAResults} disabled={!results.some(result => !result.matched)} className="rounded-lg border border-amber-300 bg-amber-50 px-3 py-1.5 text-xs font-medium text-amber-800 disabled:opacity-40">تحديد كل غير المطابق</button>
          <button onClick={() => requestStageATransfer("matched")} disabled={!results.some(result => result.matched)} className="rounded-lg bg-emerald-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-emerald-700 disabled:opacity-40">ترحيل المطابق ({results.filter(result => result.matched).length})</button>
          <button onClick={() => requestStageATransfer("matched-and-selected")} disabled={!results.some(result => result.matched || selectedResultIds.has(resultKey(result)))} className="rounded-lg bg-blue-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-blue-700 disabled:opacity-40">ترحيل المطابق والمحدد</button>
          <button onClick={() => requestStageATransfer("all")} className="rounded-lg bg-amber-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-amber-700">ترحيل الكل ({results.length})</button>
          {selectedResultIds.size > 0 && <button onClick={clearAllStageASelection} className="rounded-lg border px-3 py-1.5 text-xs">إلغاء التحديد</button>}
        </div>
      </div>
      <div className="overflow-x-auto"><table className="w-full text-sm"><thead className="bg-slate-100 text-xs"><tr><th className="p-3 text-right"><input aria-label="تحديد كل النتائج الظاهرة" type="checkbox" checked={visibleActionResults.length > 0 && visibleActionResults.every(result => selectedResultIds.has(resultKey(result)))} onChange={event => setSelectedResultIds(previous => { const next = new Set(previous); visibleActionResults.forEach(result => event.target.checked ? next.add(resultKey(result)) : next.delete(resultKey(result))); return next; })} /></th>{["المنصة", "التصنيف", "بيان الفاتورة", "اسم الزبون من قاعدة البيانات", "النوع", "مبلغ التطبيق", "إجمالي الفاتورة", "نقدي", "شيك", "دين", "الخصم", "رقم المستخدم", "تاريخ الفاتورة", "الساعة", "رقم الفاتورة", "الحالة", "ملاحظة"].map(column => <th key={column} className="p-3 text-right">{column}</th>)}</tr></thead>
        <tbody>{visibleActionResults.map(result => <tr key={`${result.source}-${result.id}`} className={`border-t ${result.matched ? "bg-green-50/30" : "bg-red-50/30"}`}>
          <td className="p-3"><input aria-label={`تحديد فاتورة ${result.originalName}`} type="checkbox" checked={selectedResultIds.has(resultKey(result))} onChange={() => toggleResultSelection(result)} /></td><td className="p-3">{result.source}</td><td className="p-3 font-semibold">{result.category}</td><td className="p-3 font-medium">{result.originalName}</td><td className="p-3">{result.databaseName || "—"}</td>
            <td className="p-3">{result.type}</td><td className="p-3 font-mono">{fmtNum(result.amount)}</td><td className="p-3 font-mono">{fmtNum(result.invoiceTotal)}</td><td className="p-3 font-mono">{fmtNum(result.cashAmount)}</td><td className="p-3 font-mono">{fmtNum(result.checkAmount)}</td><td className="p-3 font-mono">{fmtNum(result.debtAmount)}</td><td className="p-3 font-mono font-semibold text-amber-700">{fmtNum(result.discountAmount)}</td>
          <td className="p-3">{result.userId || "—"}</td><td className="p-3">{result.invoiceDate || "—"}</td><td className="p-3">{result.registrationTime || "—"}</td><td className="p-3">{result.invoiceNumber || "—"}</td>
          <td className={`p-3 font-semibold ${result.matched ? "text-green-700" : "text-red-700"}`}>{result.matched ? "مطابق" : "غير مطابق"}</td><td className="p-3 text-xs text-orange-700">{result.categoryIssue || result.issue || "—"}</td>
        </tr>)}</tbody></table></div>
    </div>}
    {transferDestinationDialog && <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-3" role="dialog" aria-modal="true" aria-label="اختيار وجهة الترحيل">
      <section className="w-full max-w-md rounded-2xl bg-white p-5 shadow-2xl">
        <h2 className="text-lg font-bold">اختيار وجهة الترحيل</h2>
        <p className="mt-2 text-sm text-slate-600">سيتم ترحيل {transferDestinationDialog.results.length} فاتورة من {transferDestinationDialog.label}. اختر المنصة:</p>
        <div className="mt-5 flex flex-wrap gap-2">
          <button type="button" onClick={() => { onTransferToB(transferDestinationDialog.results, databaseSummary); clearAllStageASelection(); setTransferDestinationDialog(null); }} className="flex-1 rounded-lg bg-emerald-600 px-4 py-2.5 text-sm font-semibold text-white hover:bg-emerald-700">منصة B</button>
          <button type="button" onClick={() => { onTransferVisaToC(transferDestinationDialog.results); clearAllStageASelection(); setTransferDestinationDialog(null); }} className="flex-1 rounded-lg bg-amber-600 px-4 py-2.5 text-sm font-semibold text-white hover:bg-amber-700">منصة C</button>
          <button type="button" onClick={() => setTransferDestinationDialog(null)} className="rounded-lg border px-4 py-2.5 text-sm">إلغاء</button>
        </div>
      </section>
    </div>}
    {showDatabaseInvoices && <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-3" role="dialog" aria-modal="true" aria-label="تفاصيل فواتير قاعدة البيانات">
        <section className="flex max-h-[92vh] w-full max-w-7xl flex-col overflow-hidden rounded-2xl bg-white shadow-2xl">
          <header className="flex items-center justify-between gap-3 border-b p-4">
            <div><h2 className="text-lg font-black">فواتير قاعدة البيانات</h2><p className="mt-1 text-xs text-slate-500">{databaseInvoiceRows.length} فاتورة · الإجمالي {fmtNum(databaseSummary.databaseInvoiceTotal)} ₪ · الخصم {fmtNum(databaseSummary.databaseDiscountTotal ?? 0)} ₪</p></div>
            <button type="button" onClick={() => setShowDatabaseInvoices(false)} aria-label="إغلاق التفاصيل" className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5" /></button>
          </header>
          <div className="grid gap-2 border-b bg-slate-50 p-3 sm:grid-cols-3 lg:grid-cols-6">
            {[["الإجمالي", databaseSummary.databaseInvoiceTotal], ["النقدي", databaseSummary.databaseCashTotal], ["الشيك", databaseInvoiceRows.reduce((sum, invoice) => sum + invoice.checkAmount, 0)], ["الكارد / التطبيقات", databaseSummary.databaseCardTotal], ["المتبقي", databaseSummary.databaseDebtTotal], ["الخصم", databaseSummary.databaseDiscountTotal ?? 0]].map(([label, value]) => <div key={String(label)} className="rounded-lg border bg-white p-2"><div className="text-[10px] text-slate-500">{label}</div><div className="font-mono text-xs font-bold">{fmtNum(Number(value))} ₪</div></div>)}
          </div>
          <div className="min-h-0 flex-1 overflow-auto">
            <table className="w-full min-w-[1100px] text-xs">
              <thead className="sticky top-0 bg-slate-100"><tr>{["رقم الفاتورة", "التاريخ", "الوقت", "اسم الزبون", "رقم الزبون", "رقم المستخدم", "الإجمالي", "نقدي", "شيك", "مدفوع بالكارد / تطبيق", "المتبقي", "الخصم"].map(label => <th key={label} className="p-2 text-right">{label}</th>)}</tr></thead>
              <tbody>{visibleDatabaseInvoices.map((invoice, index) => <tr key={`${invoice.invoiceNumber}-${index}`} className="border-t">
                <td className="p-2 font-mono">{invoice.invoiceNumber || "—"}</td><td className="p-2">{invoice.invoiceDate || "—"}</td><td className="p-2">{invoice.registrationTime || "—"}</td><td className="p-2">{invoice.name || "—"}</td><td className="p-2">{invoice.customerNumber || "—"}</td><td className="p-2">{invoice.userId || "—"}</td>
                <td className="p-2 font-mono">{fmtNum(invoice.invoiceTotal)}</td><td className="p-2 font-mono">{fmtNum(invoice.cashAmount)}</td><td className="p-2 font-mono">{fmtNum(invoice.checkAmount)}</td><td className="p-2 font-mono">{fmtNum(invoice.cardAmount)}</td><td className="p-2 font-mono">{fmtNum(invoice.debtAmount)}</td><td className={`p-2 font-mono ${invoice.discountAmount > 0 ? "font-bold text-amber-700" : ""}`}>{fmtNum(invoice.discountAmount)}</td>
              </tr>)}</tbody>
            </table>
          </div>
          <footer className="flex items-center justify-center gap-3 border-t p-3 text-xs"><button type="button" onClick={() => setDatabaseInvoicePage(page => Math.max(1, page - 1))} disabled={databaseInvoicePage <= 1} className="rounded border px-3 py-1.5 disabled:opacity-40">السابق</button><span>صفحة {Math.min(databaseInvoicePage, databaseInvoicePageCount)} من {databaseInvoicePageCount}</span><button type="button" onClick={() => setDatabaseInvoicePage(page => Math.min(databaseInvoicePageCount, page + 1))} disabled={databaseInvoicePage >= databaseInvoicePageCount} className="rounded border px-3 py-1.5 disabled:opacity-40">التالي</button></footer>
        </section>
    </div>}
  </div></div>;
}

interface StageBInvoice {
  id: string;
  name: string;
  originalName: string;
  amount: number;
  paidAmount: number;
  accountType: string;
  userId: string;
  registrationTime: string;
  customerNumber: string;
  invoiceNumber: string;
  date: string;
  source: string;
  raw: Record<string, unknown>;
  stageAInvoice?: StageAResult;
}

interface StageBTransfer {
  id: string;
  name: string;
  description: string;
  amount: number;
  receivedAmount: number;
  accountType: string;
  type: string;
  date: string;
  raw: Record<string, unknown>;
}

interface StageBExtraTransferFile {
  fileName: string;
  headers: string[];
  rows: Record<string, unknown>[];
  map: { name: string; amount: string; paidAmount: string; receivedAmount: string; accountType: string; userId: string; date: string };
  swap: boolean;
}

interface StageBMatch {
  id: string;
  invoice: StageBInvoice;
  transfer?: StageBTransfer;
  score: number;
  status: "pending" | "confirmed" | "rejected" | "held" | "heldInvoice" | "heldTransfer" | "visa" | "difference";
  reason: string;
}

function stageBNameParts(value: string): string[][] {
  const cleaned = value
    .replace(/pacs008|wallet|ct|ils|usd|eur/gi, " ")
    .replace(/\d+/g, " ")
    .replace(/[|_*:;,()[\]{}]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned.split("/").map(part => normName(part).split(" ").filter(Boolean)).filter(Boolean);
}

function stageBNameKey(value: string): string {
  const parts = stageBNameParts(value);
  const names = parts[0] || [];
  return names.length >= 2 ? `${names[0]} ${names[names.length - 1]}` : names.join(" ");
}

function stageBNameMatch(invoiceName: string, transferName: string, aliases: Record<string, string>): number {
  const invoiceKey = stageBNameKey(invoiceName);
  const aliasKey = aliases[invoiceKey];
  const segments = stageBNameParts(transferName);
  if (aliasKey && segments.some(segment => stageBNameKey(segment.join(" ")) === aliasKey)) return 1;
  const invoiceParts = stageBNameParts(invoiceName);
  return segments.reduce((best, segment) => {
    const candidate = segment.length >= 2 ? `${segment[0]} ${segment[segment.length - 1]}` : segment.join(" ");
    return Math.max(best, nameSim(invoiceParts[0]?.join(" ") || "", candidate));
  }, 0);
}

function stageBAccountType(invoice: StageBInvoice): string {
  const value = invoice.accountType || invoice.source;
  if (/بنك\s*فلسطين|bank\s*of\s*palestine/i.test(value)) return "بنك فلسطين";
  if (/محفظة\s*تجارية|بال\s*بي|pal\s*pay|palpay|wallet/i.test(value)) return "بال بي";
  if (/جوال\s*بي|jawwal/i.test(value)) return "جوال بي";
  return value.trim();
}

function stageBTransferAccountType(transfer: StageBTransfer): string {
  return transfer.accountType || transfer.type;
}

function StageBPlatform({ invoices, onBack, onTransferVisaToC, onSafeActivities, onOpenInvoiceReference, initialFile }: { invoices: StageAResult[]; onBack: () => void; onTransferVisaToC: (results: StageAResult[]) => void; onSafeActivities?: (activities: string[]) => void; onOpenInvoiceReference?: (invoice: StageAResult) => boolean; initialFile?: ProjectSourceFile }) {
  const emptyMap = useMemo(() => ({ name: "", amount: "", paidAmount: "", receivedAmount: "", accountType: "", userId: "", date: "" }), []);
  const [files, setFiles] = useState<Record<string, File | null>>({ invoiceBank: null, transferBank: null, invoiceWallet: null, transferWallet: null, transferJawwal: null });
  const [headers, setHeaders] = useState<Record<string, string[]>>({ invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [], transferJawwal: [] });
  const [rows, setRows] = useState<Record<string, Record<string, unknown>[]>>({ invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [], transferJawwal: [] });
  const [maps, setMaps] = useState<Record<string, typeof emptyMap>>({ invoiceBank: emptyMap, transferBank: emptyMap, invoiceWallet: emptyMap, transferWallet: emptyMap, transferJawwal: emptyMap });
  const [extraTransfers, setExtraTransfers] = useState<Record<string, StageBExtraTransferFile>>({});
  const [matches, setMatches] = useState<StageBMatch[]>([]);
  const [aliases, setAliases] = useState<Record<string, string>>({});
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [invoiceSearch, setInvoiceSearch] = useState("");
  const [transferSearch, setTransferSearch] = useState("");
  const [drawerMatchId, setDrawerMatchId] = useState<string | null>(null);
  const [drawerInvoiceSearch, setDrawerInvoiceSearch] = useState("");
  const [drawerTransferSearch, setDrawerTransferSearch] = useState("");
  const [drawerMinAmount, setDrawerMinAmount] = useState("");
  const [drawerMaxAmount, setDrawerMaxAmount] = useState("");
  const [drawerMode, setDrawerMode] = useState<"invoice" | "transfer">("transfer");
  const [stageBView, setStageBView] = useState<"pending" | "confirmed" | "manual">("pending");
  const rowRefs = useRef<Record<string, HTMLTableRowElement | null>>({});
  const stageBCheckboxRefs = useRef<Record<string, HTMLInputElement | null>>({});
  const [activeStageBMatchId, setActiveStageBMatchId] = useState<string | null>(null);
  const [detailMatchId, setDetailMatchId] = useState<string | null>(null);
  const [detailMode, setDetailMode] = useState<"invoice" | "transfer">("invoice");
  const undoStack = useRef<StageBMatch[][]>([]);
  const [error, setError] = useState<string | null>(null);
  const [stageBPage, setStageBPage] = useState(1);
  const [showSettings, setShowSettings] = useState(false);
  const [amountTolerancePercent, setAmountTolerancePercent] = useState(0.5);
  const [removeRepeatedInvoiceHeaders, setRemoveRepeatedInvoiceHeaders] = useState(true);
  const [stageBAccountFilter, setStageBAccountFilter] = useState("all");
  type StageBSort = "none" | "name" | "amountAsc" | "amountDesc";
  const [invoiceSort, setInvoiceSort] = useState<StageBSort>("none");
  const [transferSort, setTransferSort] = useState<StageBSort>("none");
  const [stageBSwaps, setStageBSwaps] = useState<Record<string, boolean>>({ invoiceBank: true, transferBank: false, invoiceWallet: true, transferWallet: false, transferJawwal: true });
  const sessionFileRef = useRef<HTMLInputElement>(null);
  const hasTransferredInvoices = invoices.length > 0;

  useEffect(() => {
    storageGet<Record<string, string>>("stage_b_aliases", {}).then(setAliases);
    storageGet<any>("stage_b_session", null).then(session => {
      if (!session) return;
      setHeaders({ invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [], transferJawwal: [], ...(session.headers ?? {}) });
      setRows({ invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [], transferJawwal: [], ...(session.rows ?? {}) });
      const restoredMaps = { invoiceBank: emptyMap, transferBank: emptyMap, invoiceWallet: emptyMap, transferWallet: emptyMap, transferJawwal: emptyMap, ...(session.maps ?? {}) };
      (["invoiceBank", "invoiceWallet"] as const).forEach(source => {
        const fileHeaders = session.headers?.[source] ?? [];
        if (session.stageBSwaps?.[source] === true || !fileHeaders.length) return;
        const detected = detectInvoiceFileColumns(fileHeaders);
        restoredMaps[source] = { ...restoredMaps[source], ...detected };
      });
      setMaps(restoredMaps);
      setExtraTransfers(session.extraTransfers ?? {});
      const restoredSwaps = { invoiceBank: true, transferBank: false, invoiceWallet: true, transferWallet: false, transferJawwal: true, ...(session.stageBSwaps ?? {}) };
      restoredSwaps.invoiceBank = session.stageBSwaps?.invoiceBank ?? true;
      restoredSwaps.invoiceWallet = session.stageBSwaps?.invoiceWallet ?? true;
      setStageBSwaps(restoredSwaps);
      setMatches(session.matches || []);
      setAmountTolerancePercent(session.amountTolerancePercent ?? 0.5);
    });
  }, [emptyMap]);

  const parseAmount = (value: unknown) => Number.parseFloat(String(value ?? "").replace(/[٫٬]/g, match => match === "٫" ? "." : "").replace(/,/g, "")) || 0;
  const loadFile = async (source: "invoiceBank" | "transferBank" | "invoiceWallet" | "transferWallet" | "transferJawwal", file: File) => {
    try {
      await saveStageASourceFiles("stageB", source, [file]);
      const parsed = parseSheet(await readFileBuf(file));
      setFiles(prev => ({ ...prev, [source]: file }));
      setHeaders(prev => ({ ...prev, [source]: parsed.headers }));
      setRows(prev => ({ ...prev, [source]: parsed.rows }));
      const isInvoiceFile = source === "invoiceBank" || source === "invoiceWallet";
      setMaps(prev => ({ ...prev, [source]: {
        name: autoDetect(parsed.headers, ["البيان", "بيان", "إسم الزبون", "اسم الزبون", "اسم العميل", "اسم", "name", "customer", "client"]),
        amount: detectCardPaidColumn(parsed.headers) || autoDetect(parsed.headers, HINTS.debit) || autoDetect(parsed.headers, ["amount", "مبلغ"]),
        paidAmount: isInvoiceFile
          ? detectCardPaidColumn(parsed.headers) || autoDetect(parsed.headers, ["paid", "مدفوع", "المبلغ المدفوع"]) || autoDetect(parsed.headers, HINTS.credit)
          : detectCardPaidColumn(parsed.headers) || autoDetect(parsed.headers, ["paid", "مدفوع", "المبلغ المدفوع"]) || autoDetect(parsed.headers, HINTS.debit),
        receivedAmount: isInvoiceFile
          ? autoDetect(parsed.headers, ["received", "مستلم", "المبلغ المستلم"]) || autoDetect(parsed.headers, HINTS.debit)
          : autoDetect(parsed.headers, ["received", "مستلم", "المبلغ المستلم"]) || autoDetect(parsed.headers, HINTS.credit),
        accountType: autoDetect(parsed.headers, HINTS.accountType),
        userId: "",
        date: autoDetect(parsed.headers, HINTS.date)
      } }));
      if (isInvoiceFile) setStageBSwaps(prev => ({ ...prev, [source]: true }));
      if (source === "transferJawwal") setStageBSwaps(prev => ({ ...prev, [source]: true }));
      setError(null);
    } catch (e) { setError((e as Error).message); }
  };
  useEffect(() => {
    if (initialFile?.stage === "stageB" && ["invoiceBank", "transferBank", "invoiceWallet", "transferWallet", "transferJawwal"].includes(initialFile.source)) {
      void loadFile(initialFile.source as "invoiceBank" | "transferBank" | "invoiceWallet" | "transferWallet" | "transferJawwal", initialFile.file);
    }
  }, [initialFile]);
  const processInvoiceFileHeaders = (source: "invoiceBank" | "invoiceWallet") => {
    if (!removeRepeatedInvoiceHeaders) return;
    const fileHeaders = headers[source] ?? [];
    const sourceRows = rows[source] ?? [];
    const headerKeys = fileHeaders.map(header => normName(header));
    const cleanedRows = sourceRows.filter(row => !fileHeaders.every((header, index) =>
      normName(String(row[header] ?? "")) === headerKeys[index]
    ));
    setRows(previous => ({ ...previous, [source]: cleanedRows }));
    setError(cleanedRows.length === sourceRows.length
      ? "لم يتم العثور على ترويسات مكررة لإزالتها."
      : `تمت إزالة ${sourceRows.length - cleanedRows.length} صفوف ترويسة مكررة من ${files[source]?.name ?? "ملف الفواتير"}.`);
  };

  const invoiceData = useMemo<StageBInvoice[]>(() => {
    const sources = (["invoiceBank", "invoiceWallet"] as const).flatMap(source => {
      const swap = stageBSwaps[source];
      return (rows[source] ?? []).map((row, index) => ({
        id: `${source}-${index}`, name: String(row[maps[source].name] ?? "").trim(),
        originalName: String(row[maps[source].name] ?? "").trim(), amount: parseAmount(swap ? (row[maps[source].receivedAmount] || row[maps[source].amount]) : (row[maps[source].paidAmount] || row[maps[source].amount])),
        paidAmount: parseAmount(swap ? (row[maps[source].receivedAmount] || row[maps[source].amount]) : (row[maps[source].paidAmount] || row[maps[source].amount])),
        accountType: source === "invoiceBank" ? "بنك فلسطين" : "بال بي", userId: "",
        registrationTime: "", customerNumber: "", invoiceNumber: "",
        date: String(row[maps[source].date] ?? "").trim(), source: source === "invoiceBank" ? "فاتورة بنك فلسطين" : "فاتورة بال بي", raw: row
      })).filter(row => row.name || row.amount);
    });
    if (sources.length) return sources;
    return invoices.map((invoice, index) => ({
      id: `stage-a-${invoice.id}-${index}`, name: invoice.name, originalName: invoice.originalName,
      amount: invoice.amount, accountType: invoice.category || invoice.source, userId: invoice.userId,
      paidAmount: invoice.amount,
      registrationTime: invoice.registrationTime, customerNumber: invoice.customerNumber, invoiceNumber: invoice.invoiceNumber,
      date: invoice.invoiceDate, source: invoice.category, raw: invoice.invoice, stageAInvoice: invoice
    }));
  }, [rows, maps, invoices, stageBSwaps]);

  const allTransferPlatforms = useMemo(() => [...new Set(invoiceData.map(stageBAccountType).filter(Boolean))], [invoiceData]);
  const transferPlatforms = useMemo(() => allTransferPlatforms.filter(type => !["بنك فلسطين", "بال بي", "جوال بي"].includes(type)), [allTransferPlatforms]);
  const loadExtraTransfer = async (platform: string, file: File) => {
    try {
      const parsed = parseSheet(await readFileBuf(file));
      setExtraTransfers(previous => ({
        ...previous,
        [platform]: {
          fileName: file.name,
          headers: parsed.headers,
          rows: parsed.rows,
          map: {
            name: autoDetect(parsed.headers, ["بيان الحوالة", "اسم الزبون", "اسم العميل", "description", "name", "customer"]),
            amount: detectCardPaidColumn(parsed.headers) || autoDetect(parsed.headers, HINTS.debit) || autoDetect(parsed.headers, ["amount", "مبلغ"]),
            paidAmount: autoDetect(parsed.headers, ["مدفوع", "debit", "paid"]) || autoDetect(parsed.headers, HINTS.debit),
            receivedAmount: autoDetect(parsed.headers, ["مستلم", "credit", "received"]) || autoDetect(parsed.headers, HINTS.credit),
            accountType: "", userId: "", date: autoDetect(parsed.headers, HINTS.date),
          },
          swap: false,
        },
      }));
      setError(null);
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const transferData = useMemo<StageBTransfer[]>(() => (["transferBank", "transferWallet", "transferJawwal"] as const).flatMap(source => {
    const swap = stageBSwaps[source];
    const accountType = source === "transferBank" ? "بنك فلسطين" : source === "transferWallet" ? "بال بي" : "جوال بي";
    return (rows[source] ?? []).map((row, index) => ({
      id: `${source}-${index}`, name: String(row[maps[source].name] ?? "").trim(),
      description: String(row[maps[source].name] ?? "").trim(), amount: parseAmount(swap ? (row[maps[source].paidAmount] || row[maps[source].amount]) : (row[maps[source].receivedAmount] || row[maps[source].amount])),
      receivedAmount: parseAmount(swap ? (row[maps[source].paidAmount] || row[maps[source].amount]) : (row[maps[source].receivedAmount] || row[maps[source].amount])),
      accountType, type: accountType,
      date: String(row[maps[source].date] ?? "").trim(), raw: row
    })).filter(row => row.name || row.amount);
  }).concat(Object.entries(extraTransfers).flatMap(([platform, fileData]) => {
    const swap = fileData.swap;
    const extraRows: StageBTransfer[] = fileData.rows.map((row, index) => {
      const receivedAmount = parseAmount(swap ? row[fileData.map.paidAmount] || row[fileData.map.amount] : row[fileData.map.receivedAmount] || row[fileData.map.amount]);
      return {
        id: `extra-${platform}-${index}`,
        name: String(row[fileData.map.name] ?? "").trim(),
        description: String(row[fileData.map.name] ?? "").trim(),
        amount: receivedAmount, receivedAmount, accountType: platform, type: platform,
        date: String(row[fileData.map.date] ?? "").trim(), raw: row,
      };
    }).filter(row => row.name || row.amount);
    return extraRows;
  })), [rows, maps, stageBSwaps, extraTransfers]);

  const runStageB = () => {
    if (!invoiceData.length) { setError(invoices.length ? "لا توجد فواتير مرحلة من قاعدة البيانات." : "ارفع ملفات الفواتير أولاً."); return; }
    if (!transferData.length) { setError(`ارفع ملفات الحوالات المطلوبة: ${allTransferPlatforms.join("، ") || "المنصات المرحّلة"}.`); return; }
    const used = new Set<string>();
    const output = invoiceData.map(invoice => {
      const candidates = transferData
        .filter(transfer => !used.has(transfer.id))
        .map(transfer => {
          const amountDiff = Math.abs(transfer.receivedAmount - invoice.paidAmount);
          const nameScore = stageBNameMatch(invoice.name, transfer.description, aliases);
          const accountDiff = stageBAccountType(invoice) && stageBTransferAccountType(transfer) && normName(stageBAccountType(invoice)) !== normName(stageBTransferAccountType(transfer));
          const score = nameScore * 70 + (amountDiff <= 0.01 ? 30 : amountDiff <= Math.max(0.01, invoice.amount * 0.005) ? 20 : 0);
          return { transfer, score, amountDiff, accountDiff };
        })
        .filter(candidate => candidate.amountDiff <= Math.max(0.01, invoice.amount * (amountTolerancePercent / 100)))
        .sort((a, b) => b.score - a.score)[0];
      if (!candidates) return { id: `m-${invoice.id}`, invoice, score: 0, status: "held", reason: "لم توجد حوالة مناسبة" } satisfies StageBMatch;
      used.add(candidates.transfer.id);
      const status = candidates.amountDiff > 0.01 || candidates.accountDiff ? "difference" : candidates.score >= 90 ? "pending" : "held";
      const reasons = [
        stageBNameMatch(invoice.name, candidates.transfer.description, aliases) < 0.5 ? "اختلاف اسم يحتاج مراجعة" : "",
        candidates.accountDiff ? `فرق نوع الحساب: ${stageBAccountType(invoice)} مقابل ${stageBTransferAccountType(candidates.transfer)}` : "",
        candidates.amountDiff > 0.01 ? `فرق مبلغ: الحوالة ${candidates.transfer.receivedAmount > invoice.paidAmount ? "أكثر" : "أقل"} بـ ${fmtNum(candidates.amountDiff)} شيكل` : ""
      ].filter(Boolean);
      return { id: `m-${invoice.id}-${candidates.transfer.id}`, invoice, transfer: candidates.transfer, score: Math.round(candidates.score), status, reason: reasons.join("؛ ") || "مطابقة اسم ومبلغ" } satisfies StageBMatch;
    });
    setMatches(output);
    setSelected(new Set());
    setStageBPage(1);
    setStageBView("pending");
    setError(null);
  };

  const updateStatus = (ids: Set<string>, status: StageBMatch["status"]) => setMatches(prev => {
    undoStack.current.push(prev);
    return prev.map(match => ids.has(match.id) ? { ...match, status } : match);
  });
  const replaceTransfer = (match: StageBMatch, transfer: StageBTransfer) => {
    setMatches(prev => {
      undoStack.current.push(prev);
      return prev.map(item => item.id === match.id ? {
      ...item, transfer, status: Math.abs(transfer.receivedAmount - item.invoice.paidAmount) > 0.01 || normName(stageBAccountType(item.invoice)) !== normName(stageBTransferAccountType(transfer)) ? "difference" : "pending",
      score: Math.round(stageBNameMatch(item.invoice.name, transfer.description, aliases) * 70 + (Math.abs(transfer.receivedAmount - item.invoice.paidAmount) <= 0.01 ? 30 : 0)),
      reason: "تم اختيار حوالة بديلة يدوياً"
      } : item);
    });
    setDrawerMatchId(null);
  };
  const replaceInvoice = (match: StageBMatch, invoice: StageBInvoice) => {
    setMatches(prev => {
      undoStack.current.push(prev);
      return prev.map(item => item.id === match.id ? {
      ...item, invoice, status: item.transfer && (Math.abs(item.transfer.receivedAmount - invoice.paidAmount) > 0.01 || normName(stageBAccountType(invoice)) !== normName(stageBTransferAccountType(item.transfer))) ? "difference" : "pending",
      score: item.transfer ? Math.round(stageBNameMatch(invoice.name, item.transfer.description, aliases) * 70 + (Math.abs(item.transfer.receivedAmount - invoice.paidAmount) <= 0.01 ? 30 : 0)) : 0,
      reason: "تم اختيار فاتورة بديلة يدوياً"
      } : item);
    });
    setDrawerMatchId(null);
  };
  const saveAliases = (items: StageBMatch[]) => {
    const next = { ...aliases };
    items.forEach(item => { if (item.transfer) next[stageBNameKey(item.invoice.name)] = stageBNameKey(item.transfer.name); });
    setAliases(next);
    void storageSet("stage_b_aliases", next);
  };
  const confirmSelected = () => {
    const items = matches.filter(match => selected.has(match.id) && match.transfer);
    if (!items.length) return;
    updateStatus(new Set(items.map(item => item.id)), "confirmed");
    saveAliases(items);
    onSafeActivities?.(items.map(item => `${item.invoice.name} ↔ ${item.transfer?.description || "—"} · ${fmtNum(item.invoice.paidAmount)} ₪`));
    setSelected(new Set());
  };
  const visibleMatches = matches.filter(match => {
    const invoiceHit = !invoiceSearch.trim() || `${match.invoice.name} ${match.invoice.userId}`.toLowerCase().includes(invoiceSearch.trim().toLowerCase());
    const transferHit = !transferSearch.trim() || match.transfer?.description.toLowerCase().includes(transferSearch.trim().toLowerCase());
    const accountHit = stageBAccountFilter === "all" || stageBAccountFilter === stageBAccountType(match.invoice) || stageBAccountFilter === (match.transfer ? stageBTransferAccountType(match.transfer) : "");
    const viewHit = stageBView === "confirmed"
      ? match.status === "confirmed"
      : stageBView === "manual"
        ? ["heldInvoice", "heldTransfer", "difference", "rejected"].includes(match.status)
        : !["confirmed", "heldInvoice", "heldTransfer", "difference", "rejected"].includes(match.status);
    return invoiceHit && transferHit && accountHit && viewHit;
  }).sort((a, b) => {
    const compare = (sort: StageBSort, left: { name: string; amount: number }, right: { name: string; amount: number }) => {
      if (sort === "name") return left.name.localeCompare(right.name, "ar");
      if (sort === "amountAsc") return left.amount - right.amount;
      if (sort === "amountDesc") return right.amount - left.amount;
      return 0;
    };
    const invoiceResult = compare(invoiceSort, a.invoice, b.invoice);
    if (invoiceResult) return invoiceResult;
    return compare(
      transferSort,
      { name: a.transfer?.description || "", amount: a.transfer?.receivedAmount || 0 },
      { name: b.transfer?.description || "", amount: b.transfer?.receivedAmount || 0 },
    );
  });
  const stageBPageSize = 20;
  const totalStageBPages = Math.max(1, Math.ceil(visibleMatches.length / stageBPageSize));
  const currentStageBPage = Math.min(stageBPage, totalStageBPages);
  const pagedMatches = visibleMatches.slice((currentStageBPage - 1) * stageBPageSize, currentStageBPage * stageBPageSize);
  const pageMatchIds = new Set(pagedMatches.map(match => match.id));
  const pageIsSelected = pagedMatches.length > 0 && pagedMatches.every(match => selected.has(match.id));
  const togglePage = () => setSelected(previous => {
    const next = new Set(previous);
    if (pageIsSelected) {
      pageMatchIds.forEach(id => next.delete(id));
    } else {
      pageMatchIds.forEach(id => next.add(id));
    }
    return next;
  });
  const toggleAll = togglePage;
  const openDetails = (matchId: string, mode: "invoice" | "transfer") => {
    setDetailMode(mode);
    setDetailMatchId(matchId);
  };
  const openReplacement = (matchId: string, mode: "invoice" | "transfer") => {
    setDrawerMode(mode);
    setDrawerMatchId(matchId);
    window.setTimeout(() => rowRefs.current[matchId]?.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" }), 0);
  };
  useEffect(() => {
    const onFocusIn = (event: FocusEvent) => {
      const target = event.target;
      if (!(target instanceof HTMLElement) || !target.closest(".stage-b-results")) return;
      const entry = Object.entries(rowRefs.current).find(([, row]) => row?.contains(target));
      if (entry) setActiveStageBMatchId(entry[0]);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      const target = event.target;
      const isTextEntry = target instanceof HTMLElement && (
        target.isContentEditable ||
        target.matches("input:not([type='checkbox']), textarea, select")
      );
      if (event.key === "Escape") {
        setDrawerMatchId(null);
        setDetailMatchId(null);
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.code === "KeyZ") {
        if (isTextEntry) return;
        event.preventDefault();
        const previous = undoStack.current.pop();
        if (previous) setMatches(previous);
        return;
      }
      if (isTextEntry || event.ctrlKey || event.metaKey || event.altKey) return;

      const focusedCheckbox = target instanceof HTMLInputElement && target.type === "checkbox" ? target : null;
      if (focusedCheckbox?.closest(".stage-b-results") && (event.code === "ArrowDown" || event.code === "ArrowUp" || event.code === "Tab")) {
        const focusedId = pagedMatches.find(match => rowRefs.current[match.id]?.contains(focusedCheckbox))?.id;
        const currentIndex = pagedMatches.findIndex(match => match.id === (focusedId || activeStageBMatchId));
        const direction = event.code === "ArrowUp" || (event.code === "Tab" && event.shiftKey) ? -1 : 1;
        const nextIndex = currentIndex + direction;
        if (nextIndex >= 0 && nextIndex < pagedMatches.length) {
          event.preventDefault();
          const nextMatch = pagedMatches[nextIndex];
          setActiveStageBMatchId(nextMatch.id);
          stageBCheckboxRefs.current[nextMatch.id]?.focus();
        }
        return;
      }

      const currentId = activeStageBMatchId && pagedMatches.some(match => match.id === activeStageBMatchId)
        ? activeStageBMatchId
        : [...selected].find(id => pagedMatches.some(match => match.id === id));
      if (!currentId) return;
      const key = event.code;
      if (key === "KeyR" || key === "KeyT") {
        event.preventDefault();
        const match = matches.find(item => item.id === currentId);
        const openedReference = key === "KeyR" && match?.invoice.stageAInvoice
          ? onOpenInvoiceReference?.(match.invoice.stageAInvoice)
          : false;
        if (!openedReference) {
          openDetails(currentId, key === "KeyR" ? "invoice" : "transfer");
        }
      } else if (key === "KeyF" || key === "KeyH") {
        event.preventDefault();
        openReplacement(currentId, key === "KeyF" ? "invoice" : "transfer");
      } else if (key === "KeyX") {
        event.preventDefault();
        updateStatus(new Set([currentId]), "rejected");
      }
    };
    window.addEventListener("focusin", onFocusIn);
    window.addEventListener("keydown", onKeyDown);
    return () => {
      window.removeEventListener("focusin", onFocusIn);
      window.removeEventListener("keydown", onKeyDown);
    };
  }, [activeStageBMatchId, matches, onOpenInvoiceReference, pagedMatches, selected]);
  const suggestionClass = (amountDiff: number, nameMatch: number) =>
    amountDiff <= 0.01 && nameMatch >= 0.5
      ? "border-green-300 bg-green-50 text-green-800"
      : amountDiff <= 0.01
        ? "border-amber-300 bg-amber-50 text-amber-800"
        : "border-red-300 bg-red-50 text-red-800";

  const exportStageB = () => {
    const workbook = XLSX.utils.book_new();
    const rowsFor = (items: StageBMatch[]) => items.map(match => [
      match.invoice.userId, match.invoice.registrationTime, match.invoice.customerNumber, match.invoice.invoiceNumber,
      match.invoice.name, match.invoice.amount, match.invoice.accountType || "—",
      match.transfer?.description || "—", match.transfer?.amount ?? "", match.transfer?.accountType || "—",
      match.transfer?.type || "—", match.transfer ? fmtNum(match.transfer.receivedAmount - match.invoice.paidAmount) : "", stageBAccountType(match.invoice) || "—",
      match.transfer ? stageBTransferAccountType(match.transfer) : "—", match.score, match.reason
    ]);
    const headersOut = ["رقم المستخدم", "الساعة", "رقم الزبون", "رقم الفاتورة", "الفاتورة", "مبلغ الفاتورة", "نوع حساب الفاتورة", "الحوالة", "مبلغ الحوالة", "نوع حوالة الملف", "فرق المبلغ", "نوع الحساب الفعلي للفاتورة", "نوع الحساب الفعلي للحوالة", "الثقة", "الملاحظة"];
    const add = (name: string, items: StageBMatch[]) => XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([headersOut, ...rowsFor(items)]), name);
    add("المطابقة العامة", matches);
    add("المؤكد والمطابق", matches.filter(match => match.status === "confirmed"));
    add("المنتظر / قيد المطابقة", matches.filter(match => ["held", "heldInvoice", "heldTransfer", "rejected"].includes(match.status)));
    add("الفيزا", matches.filter(match => match.status === "visa"));
    add("فروقات تحتاج تدقيق", matches.filter(match => match.status === "difference"));
    XLSX.writeFile(workbook, "مطابقة_B_الفواتير_والحوالات.xlsx");
  };

  const saveStageBSession = () => {
    void storageSet("stage_b_session", { headers, rows, maps, matches, amountTolerancePercent, extraTransfers, stageBSwaps });
    setError(null);
  };

  const exportStageBSession = () => {
    const blob = new Blob([JSON.stringify({ headers, rows, maps, matches, amountTolerancePercent, extraTransfers, stageBSwaps })], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "جلسة_المرحلة_B.json";
    link.click();
    URL.revokeObjectURL(url);
  };

  const clearStageBSession = () => {
    if (!window.confirm("مسح ملفات ونتائج المرحلة B الحالية؟")) return;
    setFiles({ invoiceBank: null, transferBank: null, invoiceWallet: null, transferWallet: null, transferJawwal: null });
    setHeaders({ invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [], transferJawwal: [] });
    setRows({ invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [], transferJawwal: [] });
    setMaps({ invoiceBank: emptyMap, transferBank: emptyMap, invoiceWallet: emptyMap, transferWallet: emptyMap, transferJawwal: emptyMap });
    setExtraTransfers({});
    setStageBSwaps({ invoiceBank: true, transferBank: false, invoiceWallet: true, transferWallet: false, transferJawwal: true });
    setMatches([]);
    setSelected(new Set());
    void storageSet("stage_b_session", null);
  };

  const importStageBSession = async (file: File) => {
    try {
      const session = JSON.parse(new TextDecoder().decode(await readFileBuf(file)));
      if (!session || typeof session !== "object" || !session.rows || !session.maps) throw new Error("ملف المرحلة B غير صالح.");
      setHeaders({ invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [], transferJawwal: [], ...(session.headers ?? {}) });
      setRows({ invoiceBank: [], transferBank: [], invoiceWallet: [], transferWallet: [], transferJawwal: [], ...(session.rows ?? {}) });
      const restoredMaps = { invoiceBank: emptyMap, transferBank: emptyMap, invoiceWallet: emptyMap, transferWallet: emptyMap, transferJawwal: emptyMap, ...(session.maps ?? {}) };
      (["invoiceBank", "invoiceWallet"] as const).forEach(source => {
        const fileHeaders = session.headers?.[source] ?? [];
        if (session.stageBSwaps?.[source] === true || !fileHeaders.length) return;
        const detected = detectInvoiceFileColumns(fileHeaders);
        restoredMaps[source] = { ...restoredMaps[source], ...detected };
      });
      setMaps(restoredMaps);
      setExtraTransfers(session.extraTransfers ?? {});
      const restoredSwaps = { invoiceBank: true, transferBank: false, invoiceWallet: true, transferWallet: false, transferJawwal: true, ...(session.stageBSwaps ?? {}) };
      restoredSwaps.invoiceBank = session.stageBSwaps?.invoiceBank ?? true;
      restoredSwaps.invoiceWallet = session.stageBSwaps?.invoiceWallet ?? true;
      setStageBSwaps(restoredSwaps);
      setMatches(session.matches || []);
      setAmountTolerancePercent(session.amountTolerancePercent ?? 0.5);
      setError(null);
      void storageSet("stage_b_session", session);
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const stageBStats = [
    { label: "مؤكد", value: matches.filter(match => match.status === "confirmed").length, color: "border-blue-300 bg-blue-50 text-green-700" },
    { label: "منتظر", value: matches.filter(match => ["pending", "held"].includes(match.status)).length, color: "border-slate-200 bg-white text-blue-700" },
    { label: "مرفوض", value: matches.filter(match => match.status === "rejected").length, color: "border-slate-200 bg-white text-red-600" },
    { label: "مطابقة يدوية", value: matches.filter(match => ["heldInvoice", "heldTransfer", "difference"].includes(match.status)).length, color: "border-slate-200 bg-white text-purple-700" },
    { label: "فيزا", value: matches.filter(match => match.status === "visa").length, color: "border-slate-200 bg-white text-teal-700" },
    { label: "فواتير", value: invoiceData.length, color: "border-slate-200 bg-white text-orange-600" },
    { label: "حوالات", value: transferData.length, color: "border-slate-200 bg-white text-orange-600" },
    { label: "إجمالي المطابقات", value: matches.length, color: "border-slate-200 bg-white text-slate-700" },
  ];

  const renderFile = (source: "invoiceBank" | "transferBank" | "invoiceWallet" | "transferWallet" | "transferJawwal", label: string) => (
    <div className="space-y-3 rounded-xl border border-slate-200 bg-white p-4">
      <h2 className="text-sm font-bold">{label}</h2>
      <DropZone file={files[source]} rowCount={rows[source]?.length ?? 0} onFile={file => loadFile(source, file)} onClear={() => { setFiles(prev => ({ ...prev, [source]: null })); setHeaders(prev => ({ ...prev, [source]: [] })); setRows(prev => ({ ...prev, [source]: [] })); }} />
      {!!headers[source].length && <>
        {(source === "invoiceBank" || source === "invoiceWallet") && <div className="flex flex-wrap items-center gap-3">
          <label className="flex items-center gap-2 text-xs text-slate-600"><input type="checkbox" checked={removeRepeatedInvoiceHeaders} onChange={event => setRemoveRepeatedInvoiceHeaders(event.target.checked)} />إزالة الترويسات المكررة</label>
          <button type="button" onClick={() => processInvoiceFileHeaders(source)} className="rounded-lg border border-blue-200 bg-blue-50 px-3 py-2 text-xs font-semibold text-blue-800 hover:bg-blue-100">معالجة الملف وإزالة الترويسات</button>
        </div>}
        <div className="grid grid-cols-2 gap-2">
          <Sel label="الاسم / البيان" headers={headers[source]} value={maps[source].name} onChange={value => setMaps(prev => ({ ...prev, [source]: { ...prev[source], name: value } }))} />
          <Sel label={source === "invoiceBank" || source === "invoiceWallet" ? "المبلغ المدفوع (Credit)" : "المبلغ المدفوع"} headers={headers[source]} value={maps[source].paidAmount} onChange={value => setMaps(prev => ({ ...prev, [source]: { ...prev[source], paidAmount: value, amount: value } }))} />
          <Sel label={source === "invoiceBank" || source === "invoiceWallet" ? "المبلغ المستلم (Debit)" : "المبلغ المستلم"} headers={headers[source]} value={maps[source].receivedAmount} onChange={value => setMaps(prev => ({ ...prev, [source]: { ...prev[source], receivedAmount: value } }))} />
          <Sel label="نوع الحساب" headers={headers[source]} value={maps[source].accountType} onChange={value => setMaps(prev => ({ ...prev, [source]: { ...prev[source], accountType: value } }))} />
        </div>
        <label className="flex w-fit cursor-pointer items-center gap-2 text-xs text-slate-600">
          <input type="checkbox" checked={stageBSwaps[source]} onChange={event => setStageBSwaps(prev => ({ ...prev, [source]: event.target.checked }))} className="rounded" />
          {source === "invoiceBank" || source === "invoiceWallet"
            ? stageBSwaps[source] ? "المستلم مدين والمدفوع دائن" : "المدفوع مدين والمستلم دائن"
            : "عكس المدين والدائن"}
        </label>
      </>}
    </div>
  );
  const renderExtraTransferFile = (platform: string) => {
    const fileData = extraTransfers[platform];
    return <div key={platform} className="space-y-3 rounded-xl border border-indigo-200 bg-white p-4">
      <h2 className="text-sm font-bold">حوالات {platform}</h2>
      <DropZone file={files[`transferExtra:${platform}`] ?? null} rowCount={fileData?.rows.length ?? 0} onFile={file => {
        setFiles(previous => ({ ...previous, [`transferExtra:${platform}`]: file }));
        void loadExtraTransfer(platform, file);
      }} onClear={() => {
        setFiles(previous => ({ ...previous, [`transferExtra:${platform}`]: null }));
        setExtraTransfers(previous => { const next = { ...previous }; delete next[platform]; return next; });
      }} />
      {!files[`transferExtra:${platform}`] && fileData?.fileName && <p className="text-xs text-emerald-700">تم استعادة بيانات الملف: {fileData.fileName}</p>}
      {!!fileData?.headers.length && <>
        <div className="grid grid-cols-2 gap-2">
          <Sel label="الاسم / البيان" headers={fileData.headers} value={fileData.map.name} onChange={value => setExtraTransfers(previous => ({ ...previous, [platform]: { ...fileData, map: { ...fileData.map, name: value } } }))} />
          <Sel label="المبلغ" headers={fileData.headers} value={fileData.map.amount} onChange={value => setExtraTransfers(previous => ({ ...previous, [platform]: { ...fileData, map: { ...fileData.map, amount: value } } }))} />
          <Sel label="المدفوع (Debit)" headers={fileData.headers} value={fileData.map.paidAmount} onChange={value => setExtraTransfers(previous => ({ ...previous, [platform]: { ...fileData, map: { ...fileData.map, paidAmount: value } } }))} />
          <Sel label="المستلم (Credit)" headers={fileData.headers} value={fileData.map.receivedAmount} onChange={value => setExtraTransfers(previous => ({ ...previous, [platform]: { ...fileData, map: { ...fileData.map, receivedAmount: value } } }))} />
          <Sel label="التاريخ" headers={fileData.headers} value={fileData.map.date} onChange={value => setExtraTransfers(previous => ({ ...previous, [platform]: { ...fileData, map: { ...fileData.map, date: value } } }))} />
        </div>
        <label className="flex w-fit cursor-pointer items-center gap-2 text-xs text-slate-600"><input type="checkbox" checked={fileData.swap} onChange={event => setExtraTransfers(previous => ({ ...previous, [platform]: { ...fileData, swap: event.target.checked } }))} />عكس المدين والدائن</label>
      </>}
    </div>;
  };

  return <div dir="rtl" className="min-h-screen bg-slate-50 p-6 text-slate-900"><div className="mx-auto max-w-7xl space-y-5">
    <div className="sticky top-2 z-20 flex flex-wrap items-center justify-between gap-4 rounded-2xl border border-slate-200 bg-white/95 p-5">
      <div className="flex items-center gap-3">
        <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-blue-600"><ChevronLeft className="h-4 w-4" />العودة للمنصة الرئيسية</button>
        <span className="hidden h-8 w-px bg-slate-200 sm:block" />
        <div><h1 className="text-xl font-black text-slate-900">B — المطابقة الذكية للفواتير والحوالات</h1><p className="mt-1 text-xs text-slate-500">مطابقة المبالغ والأسماء بين الفواتير وحوالات بنك فلسطين والمحفظة التجارية</p></div>
      </div>
      <div className="flex flex-wrap gap-2">
        <button onClick={() => setShowSettings(value => !value)} className="rounded-lg border px-3 py-2 text-xs">الإعدادات</button>
        <button onClick={saveStageBSession} className="rounded-lg border border-blue-200 bg-blue-50 px-3 py-2 text-xs text-blue-700">حفظ المشروع</button>
        <button onClick={exportStageBSession} className="rounded-lg border px-3 py-2 text-xs">تصدير المشروع</button>
        <button onClick={() => sessionFileRef.current?.click()} className="rounded-lg border px-3 py-2 text-xs">استيراد ملف سابق</button>
        <input ref={sessionFileRef} type="file" accept=".json,application/json" className="hidden" onChange={event => { const file = event.target.files?.[0]; if (file) void importStageBSession(file); event.target.value = ""; }} />
        <button onClick={clearStageBSession} className="rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-xs text-red-700">أدوات الحذف</button>
        <button onClick={runStageB} className="rounded-lg bg-emerald-600 px-4 py-2 text-sm font-medium text-white"><Sparkles className="mr-1 inline h-4 w-4" />تشغيل المطابقة</button>
        {matches.length > 0 && <button onClick={exportStageB} className="rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white"><Download className="mr-1 inline h-4 w-4" />تصدير الشيتات</button>}
      </div>
    </div>
    {showSettings && <div className="rounded-xl border border-blue-100 bg-blue-50/60 p-4">
      <label className="flex items-center gap-2 text-xs">نسبة التسامح في فرق المبلغ
        <input type="number" min="0" max="10" step="0.1" value={amountTolerancePercent} onChange={event => setAmountTolerancePercent(Math.min(10, Math.max(0, Number(event.target.value) || 0)))} className="w-20 rounded-lg border bg-white px-2 py-1.5 text-center" />%
      </label>
    </div>}
    <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-4 text-sm leading-6 text-emerald-900">تتم المطابقة على أساس المبلغ والاسم معًا. عند الترحيل من A تظهر خانة حوالات لكل منصة رُحّلت فواتيرها، بما فيها بنك فلسطين وبال بي وجوال بي والمنصات الإضافية. نوع الحساب يُحدد من المنصة التي رُحّلت منها الفاتورة.</div>
    {invoices.length > 0 && <div className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">
      <div><strong>فواتير المرحلة B</strong><div className="mt-1 text-xs">يمكنك نقل كل الفواتير إلى منصة C بصرف النظر عن تصنيفها الأصلي.</div></div>
      <button onClick={() => onTransferVisaToC(invoices)} className="rounded-lg bg-amber-600 px-4 py-2 text-xs font-medium text-white hover:bg-amber-700">ترحيل كل الفواتير إلى C ({invoices.length})</button>
    </div>}
    {error && <div className="rounded-lg bg-red-100 p-3 text-sm text-red-800">{error}</div>}
    <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
      {hasTransferredInvoices ? <>{renderFile("transferBank", "حوالات بنك فلسطين")}{renderFile("transferWallet", "حوالات بال بي")}{renderFile("transferJawwal", "حوالات جوال بي")}{transferPlatforms.map(renderExtraTransferFile)}</> : <>{renderFile("invoiceBank", "فواتير بنك فلسطين")}{renderFile("invoiceWallet", "فواتير بال بي")}{renderFile("transferBank", "حوالات بنك فلسطين")}{renderFile("transferWallet", "حوالات بال بي")}{renderFile("transferJawwal", "حوالات جوال بي")}</>}
    </div>
    <div className="grid grid-cols-2 gap-2 sm:grid-cols-4 lg:grid-cols-8">
      {stageBStats.map(stat => <div key={stat.label} className={`rounded-xl border p-3 shadow-[0_3px_10px_rgba(30,64,175,0.06)] ${stat.color}`}>
        <div className="text-[11px] font-medium text-slate-600">{stat.label}</div>
        <div className="mt-1 text-2xl font-bold">{stat.value}</div>
      </div>)}
    </div>
    {!!matches.length && <div className={`stage-b-results overflow-hidden rounded-xl border bg-white ${stageBView === "confirmed" ? "stage-b-confirmed" : ""}`}>
      <div className="flex flex-wrap items-center gap-2 border-b bg-slate-50 p-3"><div className="flex gap-1 rounded-lg bg-white p-1"><button onClick={() => { setStageBView("pending"); setStageBPage(1); }} className={`rounded px-3 py-1.5 text-xs ${stageBView === "pending" ? "bg-amber-100 text-amber-800" : "text-slate-600"}`}>المنتظر {matches.filter(match => !["confirmed", "heldInvoice", "heldTransfer", "difference", "rejected"].includes(match.status)).length}</button><button onClick={() => { setStageBView("confirmed"); setStageBPage(1); }} className={`rounded px-3 py-1.5 text-xs ${stageBView === "confirmed" ? "bg-green-100 text-green-800" : "text-slate-600"}`}>المؤكد {matches.filter(match => match.status === "confirmed").length}</button><button onClick={() => { setStageBView("manual"); setStageBPage(1); }} className={`rounded px-3 py-1.5 text-xs ${stageBView === "manual" ? "bg-blue-100 text-blue-800" : "text-slate-600"}`}>مطابقة يدوية {matches.filter(match => ["heldInvoice", "heldTransfer", "difference", "rejected"].includes(match.status)).length}</button></div><input value={invoiceSearch} onChange={event => { setInvoiceSearch(event.target.value); setStageBPage(1); }} placeholder="بحث بالاسم أو البيان أو رقم المستخدم..." className="min-w-52 flex-1 rounded-lg border px-3 py-2 text-xs" /><input value={transferSearch} onChange={event => { setTransferSearch(event.target.value); setStageBPage(1); }} placeholder="بحث في الحوالات..." className="min-w-52 flex-1 rounded-lg border px-3 py-2 text-xs" /><select value={invoiceSort} onChange={event => { setInvoiceSort(event.target.value as StageBSort); setStageBPage(1); }} className="rounded-lg border bg-white px-3 py-2 text-xs"><option value="none">فواتير: بدون فرز</option><option value="name">فواتير: بالاسم</option><option value="amountDesc">فواتير: من الكبير للصغير</option><option value="amountAsc">فواتير: من الصغير للكبير</option></select><select value={transferSort} onChange={event => { setTransferSort(event.target.value as StageBSort); setStageBPage(1); }} className="rounded-lg border bg-white px-3 py-2 text-xs"><option value="none">حوالات: بدون فرز</option><option value="name">حوالات: بالاسم</option><option value="amountDesc">حوالات: من الكبير للصغير</option><option value="amountAsc">حوالات: من الصغير للكبير</option></select><select value={stageBAccountFilter} onChange={event => { setStageBAccountFilter(event.target.value); setStageBPage(1); }} className="rounded-lg border bg-white px-3 py-2 text-xs"><option value="all">كل الأنواع</option>{allTransferPlatforms.map(platform => <option key={platform} value={platform}>{platform}</option>)}</select>      <button onClick={togglePage} className="rounded-lg border px-3 py-2 text-xs">{pageIsSelected ? "إلغاء تحديد الصفحة" : `تحديد الصفحة (${pagedMatches.length})`}</button><button onClick={confirmSelected} disabled={!selected.size} className="rounded-lg bg-blue-600 px-3 py-2 text-xs text-white disabled:opacity-40">حفظ المحدد ({selected.size})</button><button onClick={() => updateStatus(selected, "rejected")} disabled={!selected.size} className="rounded-lg bg-red-600 px-3 py-2 text-xs text-white disabled:opacity-40">رفض المحدد</button></div>
      <div className="mb-3 rounded-lg border border-indigo-100 bg-indigo-50 px-3 py-2 text-[11px] leading-5 text-indigo-900">اختصارات للصف المحدد (تعمل حتى مع تخطيط لوحة مفاتيح عربي): <b>R</b> أماكن الفاتورة · <b>T</b> بيانات الحوالة · <b>F</b> استبدال فاتورة · <b>H</b> استبدال حوالة · <b>X</b> رفض · <b>Space</b> تحديد مربع الاختيار · <b>↑/↓</b> أو <b>Tab</b> للتنقل بين مربعات الصفوف.</div><div className="flex items-start gap-4"><div className="min-w-0 flex-1 overflow-auto"><table className="w-full min-w-[1000px] text-sm"><thead className="sticky top-0 z-10 bg-slate-100 text-xs"><tr><th className="p-3">      <input type="checkbox" checked={pagedMatches.length > 0 && pagedMatches.every(match => selected.has(match.id))} onChange={toggleAll} /></th><th className="p-3 text-right">الفاتورة / المستخدم</th><th className="p-3 text-right">المبلغ</th><th className="p-3 text-right">نوع الحساب</th><th className="p-3 text-right">الحوالة</th><th className="p-3 text-right">مبلغ الحوالة</th><th className="p-3 text-right">نوع الحساب</th><th className="p-3 text-right">الحالة</th><th className="p-3 text-right">إجراءات</th></tr></thead>            <tbody>{pagedMatches.map(match => <tr ref={row => { rowRefs.current[match.id] = row; stageBCheckboxRefs.current[match.id] = row?.querySelector<HTMLInputElement>("input[type=checkbox]") ?? null; }} key={match.id} className="border-t align-top"><td className="p-3"><input type="checkbox" checked={selected.has(match.id)} onFocus={() => setActiveStageBMatchId(match.id)} onClick={() => setActiveStageBMatchId(match.id)} onChange={() => setSelected(prev => { const next = new Set(prev); if (next.has(match.id)) next.delete(match.id); else next.add(match.id); return next; })} /></td><td className="p-3 font-medium">{match.invoice.name}      <div className="text-[10px] text-slate-500">المستخدم: {match.invoice.userId || "—"} · الزبون: {match.invoice.customerNumber || "—"} · الفاتورة: {match.invoice.invoiceNumber || "—"} · الساعة: {match.invoice.registrationTime || "—"}</div></td><td className="p-3 font-mono">{fmtNum(match.invoice.amount)}</td><td className="p-3">{match.invoice.accountType || "—"}</td><td className="p-3">{match.transfer?.description || "لا توجد حوالة مناسبة"}</td><td className="p-3 font-mono">{match.transfer ? fmtNum(match.transfer.amount) : "—"}</td><td className="p-3">{match.transfer?.accountType || "—"}</td><td className="p-3"><span className={match.status === "confirmed" ? "text-green-700" : match.status === "difference" ? "text-orange-700" : match.status === "visa" ? "text-teal-700" : "text-amber-700"}>{match.status === "confirmed" ? "مؤكد ومطابق" : match.status === "rejected" ? "مرفوض" : match.status === "heldInvoice" ? "معلق: فاتورة" : match.status === "heldTransfer" ? "معلق: حوالة" : match.status === "held" ? "معلق" : match.status === "visa" ? "فيزا" : "فرق يحتاج تدقيق"}</span><div className="text-[10px] text-slate-500">{match.reason} · {match.score}%</div></td><td className="p-3"><div className="flex flex-wrap gap-1"><button onClick={() => { const one = new Set([match.id]); if (match.transfer) saveAliases([match]); updateStatus(one, "confirmed"); }} className="rounded border border-green-200 bg-green-50 px-2 py-1 text-[10px] text-green-700">تأكيد</button><button onClick={() => updateStatus(new Set([match.id]), "rejected")} className="rounded border border-red-200 bg-red-50 px-2 py-1 text-[10px] text-red-700">رفض</button>            <button onClick={() => openReplacement(match.id, "invoice")} className="rounded border border-emerald-200 bg-emerald-50 px-2 py-1 text-[10px] text-emerald-700">سحب كاشير</button><button onClick={() => openReplacement(match.id, "transfer")} className="rounded border border-blue-200 bg-blue-50 px-2 py-1 text-[10px] text-blue-700">سحب حوالة</button><button onClick={() => openDetails(match.id, "invoice")} className="rounded border border-slate-200 bg-slate-50 px-2 py-1 text-[10px] text-slate-700" title="تفاصيل">i تفاصيل</button><button onClick={() => updateStatus(new Set([match.id]), "heldInvoice")} className="rounded border border-amber-200 bg-amber-50 px-2 py-1 text-[10px] text-amber-700">علق فاتورة</button><button onClick={() => updateStatus(new Set([match.id]), "heldTransfer")} className="rounded border border-amber-200 bg-amber-50 px-2 py-1 text-[10px] text-amber-700">علق حوالة</button>      {stageAVisaNumber(match.invoice.originalName) && <button onClick={() => updateStatus(new Set([match.id]), "visa")} className="rounded border border-teal-200 bg-teal-50 px-2 py-1 text-[10px] text-teal-700">ترحيل للفيزا</button>}</div></td></tr>)}            </tbody></table></div>{drawerMatchId && (() => { const current = matches.find(match => match.id === drawerMatchId); if (!current) return null;             const transferCandidates = transferData.filter(item => item.id !== current.transfer?.id && (!drawerMinAmount || item.receivedAmount >= Number(drawerMinAmount)) && (!drawerMaxAmount || item.receivedAmount <= Number(drawerMaxAmount))).slice(0, 30); const invoiceCandidates = invoiceData.filter(item => item.id !== current.invoice.id && (!drawerMinAmount || item.paidAmount >= Number(drawerMinAmount)) && (!drawerMaxAmount || item.paidAmount <= Number(drawerMaxAmount))).slice(0, 30);       return <aside className="replacement-drawer sticky top-3 shrink-0 rounded-xl border bg-white p-3 shadow-lg"><div className="mb-2 flex items-center justify-between"><strong className="text-sm">بحث واستبدال</strong><button onClick={() => setDrawerMatchId(null)} className="text-xs text-slate-500">إغلاق</button></div>      <div className="space-y-3">      {drawerMode === "invoice" && <section><p className="mb-1 text-[10px] font-bold text-emerald-700">بحث وسحب من الكاشير</p><input value={drawerInvoiceSearch} onChange={event => setDrawerInvoiceSearch(event.target.value)} placeholder="ابحث باسم أو مبلغ الفاتورة..." className="mb-2 w-full rounded-lg border px-2 py-1.5 text-xs" /><div className="max-h-48 space-y-1 overflow-auto">{invoiceCandidates.filter(item => !drawerInvoiceSearch || `${item.name} ${item.amount}`.toLowerCase().includes(drawerInvoiceSearch.toLowerCase())).map(item =>       <button key={item.id} onClick={() => replaceInvoice(current, item)} className={`block w-full rounded border p-2 text-right text-[10px] ${suggestionClass(Math.abs(item.paidAmount - current.invoice.paidAmount), stageBNameMatch(item.name, current.invoice.name, aliases))}`}>{item.name} · {fmtNum(item.paidAmount)} · المستخدم {item.userId || "—"}</button>)}      </div></section>}      <section className={drawerMode === "transfer" ? "" : "hidden"}><p className="mb-1 text-[10px] font-bold text-blue-700">بحث وسحب من الحوالات</p><input value={drawerTransferSearch} onChange={event => setDrawerTransferSearch(event.target.value)} placeholder="ابحث باسم أو مبلغ الحوالة..." className="mb-2 w-full rounded-lg border px-2 py-1.5 text-xs" /><div className="max-h-48 space-y-1 overflow-auto">{transferCandidates.filter(item => !drawerTransferSearch || `${item.description} ${item.receivedAmount}`.toLowerCase().includes(drawerTransferSearch.toLowerCase())).map(item =>       <button key={item.id} onClick={() => replaceTransfer(current, item)} className={`block w-full rounded border p-2 text-right text-[10px] ${suggestionClass(Math.abs(item.receivedAmount - current.invoice.paidAmount), stageBNameMatch(current.invoice.name, item.description, aliases))}`}>{item.description} · {fmtNum(item.receivedAmount)} · {item.accountType || "—"}</button>)}</div></section></div>            {detailMatchId === current.id && <div className="mt-3 rounded bg-slate-50 p-2 text-[10px]"><strong>تفاصيل الفاتورة:</strong><div className="mt-1 space-y-1"><div>الساعة: {current.invoice.registrationTime || "—"}</div><div>رقم الزبون: {current.invoice.customerNumber || "—"}</div><div>رقم الفاتورة: {current.invoice.invoiceNumber || "—"}</div><div>رقم المستخدم: {current.invoice.userId || "—"}</div></div><pre className="mt-1 max-h-28 overflow-auto whitespace-pre-wrap">      {JSON.stringify(current.invoice.raw, null, 2)}</pre></div>}      </aside>; })()}</div><div className="flex items-center justify-center gap-3 border-t bg-slate-50 p-3 text-xs"><button onClick={() => setStageBPage(page => Math.max(1, page - 1))} disabled={currentStageBPage === 1} className="rounded border px-3 py-1.5 disabled:opacity-40">السابق</button><span>صفحة {currentStageBPage} من {totalStageBPages} · عرض {pagedMatches.length} من {visibleMatches.length}</span><button onClick={() => setStageBPage(page => Math.min(totalStageBPages, page + 1))} disabled={currentStageBPage === totalStageBPages} className="rounded border px-3 py-1.5 disabled:opacity-40">التالي</button></div>
    </div>}
    {detailMatchId && (() => {
      const match = matches.find(item => item.id === detailMatchId);
      if (!match) return null;
      const values = detailMode === "invoice"
        ? [["الاسم", match.invoice.name], ["رقم الفاتورة", match.invoice.invoiceNumber], ["رقم الزبون", match.invoice.customerNumber], ["رقم المستخدم", match.invoice.userId], ["المبلغ", fmtNum(match.invoice.amount)], ["نوع الحساب", match.invoice.accountType || match.invoice.source], ["التاريخ", match.invoice.date], ["الوقت", match.invoice.registrationTime]]
        : [["البيان", match.transfer?.description || "لا توجد حوالة مرتبطة"], ["المبلغ", match.transfer ? fmtNum(match.transfer.amount) : "—"], ["نوع الحساب", match.transfer?.accountType || "—"], ["التاريخ", match.transfer?.date || "—"]];
      return <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4" role="dialog" aria-modal="true" aria-label={detailMode === "invoice" ? "بيانات الفاتورة" : "بيانات الحوالة"} onMouseDown={event => { if (event.target === event.currentTarget) setDetailMatchId(null); }}>
        <section className="w-full max-w-xl rounded-2xl bg-white p-5 shadow-2xl">
          <header className="flex items-center justify-between gap-3 border-b pb-3"><h2 className="text-base font-bold">{detailMode === "invoice" ? "بيانات الفاتورة" : "بيانات الحوالة"}</h2><button type="button" onClick={() => setDetailMatchId(null)} className="rounded-lg border px-3 py-1.5 text-xs hover:bg-slate-50">إغلاق (Esc)</button></header>
          <dl className="mt-4 grid gap-2 sm:grid-cols-2">{values.map(([label, value]) => <div key={label} className="rounded-lg bg-slate-50 p-3"><dt className="text-[11px] text-slate-500">{label}</dt><dd className="mt-1 break-words text-sm font-semibold text-slate-900">{value || "—"}</dd></div>)}</dl>
        </section>
      </div>;
    })()}
  </div></div>;
}

interface StageCMatch {
  kind: "invoice" | "sony-only";
  invoice: StageAResult;
  sonyName: string;
  sonyAmount: number;
  sonyVisa: string;
  authNumber: string;
  matched: boolean;
  issue: string;
}

export interface StageDBankInput {
  id?: string;
  description: string;
  amount: number;
  date: string;
}

interface StageDPending extends StageDBankInput {
  authNumber: string;
}

function extractFourDigitVisa(value: string): string {
  const matches = value.match(/\d{4}(?!\d)/g);
  return matches?.[matches.length - 1] || "";
}

function extractMerchantAuthorizationNumber(description: string): string {
  const numbers = description.match(/\d+/g) || [];
  const candidate = numbers.length >= 3 ? numbers[1] : numbers.find(number => number.length === 6);
  return candidate?.length === 6 ? candidate : "";
}

function normalizeAuthorizationNumber(value: string): string {
  return value.replace(/\D/g, "");
}

function StageDPlatform({ matches, pending, onBack }: { matches: StageCMatch[]; pending: StageDBankInput[]; onBack: () => void }) {
  const [file, setFile] = useState<File | null>(null);
  const [headers, setHeaders] = useState<string[]>([]);
  const [rawRows, setRawRows] = useState<Record<string, unknown>[]>([]);
  const [map, setMap] = useState({ description: "", amount: "", date: "" });
  const [error, setError] = useState<string | null>(null);
  const parseAmount = (value: unknown) => Number.parseFloat(String(value ?? "").replace(/[٫٬]/g, match => match === "٫" ? "." : "").replace(/,/g, "")) || 0;
  const loadFile = async (nextFile: File) => {
    try {
      const parsed = parseSheet(await readFileBuf(nextFile));
      setFile(nextFile); setHeaders(parsed.headers); setRawRows(parsed.rows);
      setMap({
        description: autoDetect(parsed.headers, ["مشتريات عمولة تجار", ...HINTS.desc]),
        amount: autoDetect(parsed.headers, HINTS.debit) || autoDetect(parsed.headers, HINTS.credit) || autoDetect(parsed.headers, ["amount", "مبلغ"]),
        date: autoDetect(parsed.headers, HINTS.date)
      });
      setError(null);
    } catch (e) { setError((e as Error).message); }
  };
  const uploadedRows = rawRows
    .filter(row => isMerchantPurchaseBankDescription(String(row[map.description] ?? "")))
    .map((row, index) => ({ id: `upload-${index}`, description: String(row[map.description] ?? ""), amount: parseAmount(row[map.amount]), date: String(row[map.date] ?? "") }));
  const pendingRows = uploadedRows.length ? [...pending, ...uploadedRows] : pending;
  const bankRows: StageDPending[] = pendingRows.map(bank => ({
    ...bank,
    authNumber: extractMerchantAuthorizationNumber(bank.description)
  }));
  const usedAuthorizations = new Set<string>();
  const rows = bankRows.map(bank => {
    const authorization = normalizeAuthorizationNumber(bank.authNumber);
    const match = authorization ? matches.find(item => {
      if (!item.authNumber) return false;
      return item.authNumber.split(/\s*\+\s*/)
        .some(number => normalizeAuthorizationNumber(number) === authorization && !usedAuthorizations.has(normalizeAuthorizationNumber(number)));
    }) : undefined;
    if (match) usedAuthorizations.add(authorization);
    return { bank, match };
  });
  const invoiceRows = matches.filter(item => item.kind === "invoice");
  const unmatchedInvoices = invoiceRows.filter(item =>
    !item.authNumber || !item.authNumber.split(/\s*\+\s*/).some(number => usedAuthorizations.has(normalizeAuthorizationNumber(number)))
  );
  const unmatchedSonyRows = matches.filter(item => item.kind === "sony-only" &&
    !item.authNumber.split(/\s*\+\s*/).some(number => usedAuthorizations.has(normalizeAuthorizationNumber(number)))
  );
  const reportRows = [
    ...rows.map(({ bank, match }) => [
      "مشتريات تجار", bank.description, bank.date, bank.amount, bank.authNumber,
      match?.kind === "invoice" ? match.invoice.name : match?.sonyName || "",
      match?.sonyVisa || "", match?.authNumber || "",
      match ? match.kind === "invoice" ? "مطابق بالتفويض" : "تفويض مطابق، دون فاتورة في C" : bank.authNumber ? "مشتريات زائدة بلا تفويض سوني مطابق" : "رقم التفويض غير موجود"
    ]),
    ...unmatchedInvoices.map(item => [
      "فاتورة فيزا", item.invoice.originalName, "", item.invoice.amount, "",
      item.invoice.name, stageAVisaNumber(item.invoice.originalName), item.authNumber,
      item.matched ? "فاتورة لم تطابق مشتريات التجار في D" : "فاتورة لم تطابق سوني في C"
    ]),
    ...unmatchedSonyRows.map(item => [
      "حركة سوني زائدة", item.sonyName, "", item.sonyAmount, "",
      "", item.sonyVisa, item.authNumber, "حركة سوني بلا فاتورة أو مشتريات تجار مطابقة"
    ])
  ];
  const exportStageD = () => {
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([
      ["نوع السجل", "البيان", "التاريخ", "المبلغ", "تفويض البيان البنكي", "اسم الزبون", "رقم الفيزا", "تفويض سوني", "الحالة"],
      ...reportRows
    ]);
    sheet["!cols"] = [18, 44, 16, 14, 22, 28, 14, 20, 42].map(wch => ({ wch }));
    sheet["!autofilter"] = { ref: sheet["!ref"] || "A1:I1" };
    sheet["!rows"] = [{ hpt: 26 }];
    XLSX.utils.book_append_sheet(workbook, sheet, "مطابقة D");
    XLSX.writeFile(workbook, "مطابقة_الفيزا_المرحلة_D.xlsx");
  };
  return <div dir="rtl" className="min-h-screen bg-slate-50 p-6 text-slate-900">
    <div className="mx-auto max-w-6xl space-y-5">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-blue-600"><ChevronLeft className="h-4 w-4" />العودة إلى C</button>
        <h1 className="text-xl font-black">D — مطابقة مشتريات التجار مع تفويضات سوني</h1>
        <button onClick={exportStageD} disabled={!reportRows.length} className="rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white disabled:opacity-40"><Download className="mr-1 inline h-4 w-4" />تصدير مطابقة D</button>
      </div>
      <div className="rounded-xl border border-indigo-200 bg-indigo-50 p-4 text-sm text-indigo-900">تتم مطابقة مشتريات عمولة التجار مع رقم التفويض المحفوظ من سوني كاشير. يُقرأ رقم التفويض من الجزء الأوسط في البيان، مثل 725047 في 65880/725047/30285.</div>
      {error && <div className="rounded-lg bg-red-100 p-3 text-sm text-red-800">{error}</div>}
      <div className="rounded-xl border border-indigo-200 bg-white p-4">
        <h2 className="mb-3 text-sm font-bold text-indigo-800">حركات بنكية معلقة — مشتريات عمولة تجار</h2>
        {pending.length > 0 && <p className="mb-2 rounded-lg bg-indigo-50 p-3 text-xs text-indigo-800">تم ترحيل {pending.length} حركة من المعلّقات.</p>}
        <DropZone file={file} rowCount={rawRows.length} onFile={loadFile} onClear={() => { setFile(null); setHeaders([]); setRawRows([]); }} />
        {!!headers.length && <div className="mt-3 grid grid-cols-1 gap-2 sm:grid-cols-3">
          <Sel label="بيان الحوالة" headers={headers} value={map.description} onChange={value => setMap(prev => ({ ...prev, description: value }))} />
          <Sel label="المبلغ" headers={headers} value={map.amount} onChange={value => setMap(prev => ({ ...prev, amount: value }))} />
          <Sel label="التاريخ" headers={headers} value={map.date} onChange={value => setMap(prev => ({ ...prev, date: value }))} />
        </div>}
        <p className="mt-2 text-xs text-slate-500">عدد حركات مشتريات التجار في هذه الدفعة: {bankRows.length}.</p>
      </div>
      <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
        <table className="w-full text-sm"><thead className="bg-slate-100 text-xs"><tr><th className="p-3 text-right">نوع السجل</th><th className="p-3 text-right">البيان</th><th className="p-3 text-right">التاريخ</th><th className="p-3 text-right">المبلغ</th><th className="p-3 text-right">تفويض البيان البنكي</th><th className="p-3 text-right">اسم الزبون</th><th className="p-3 text-right">رقم الفيزا</th><th className="p-3 text-right">تفويض سوني</th><th className="p-3 text-right">الحالة</th></tr></thead>
          <tbody>{reportRows.map((reportRow, index) => <tr key={`${reportRow[0]}-${index}`} className="border-t"><td className="p-3">{reportRow[0]}</td><td className="p-3">{reportRow[1]}</td><td className="p-3">{reportRow[2] || "—"}</td><td className="p-3 font-mono">{reportRow[3] ? fmtNum(Number(reportRow[3])) : "—"}</td><td className="p-3 font-mono">{reportRow[4] || "—"}</td><td className="p-3">{reportRow[5] || "—"}</td><td className="p-3 font-mono">{reportRow[6] || "—"}</td><td className="p-3 font-mono">{reportRow[7] || "—"}</td><td className="p-3">{reportRow[8]}</td></tr>)}</tbody>
        </table>
        {!reportRows.length && <div className="p-10 text-center text-sm text-slate-500">لا توجد فواتير أو حركات لعرضها في D.</div>}
        {!!(unmatchedInvoices.length + unmatchedSonyRows.length) && <div className="border-t bg-amber-50 p-3 text-xs text-amber-900">
          غير مرتبطة في D: {unmatchedInvoices.length} فاتورة و{unmatchedSonyRows.length} حركة سوني زائدة.
        </div>}
      </div>
    </div>
  </div>;
}

function StageCPlatform({ invoices, pendingBankInputs, onBack, onTransferToD, onTransferToB, onSafeActivities, initialFile }: { invoices: StageAResult[]; pendingBankInputs: StageDBankInput[]; onBack: () => void; onTransferToD?: (matches: StageCMatch[], pending: StageDBankInput[]) => void; onTransferToB: (invoices: StageAResult[]) => void; onSafeActivities?: (activities: string[]) => void; initialFile?: ProjectSourceFile }) {
  const [file, setFile] = useState<File | null>(null);
  const [sonyFileName, setSonyFileName] = useState("");
  const [headers, setHeaders] = useState<string[]>([]);
  const [rows, setRows] = useState<Record<string, unknown>[]>([]);
  const [map, setMap] = useState({ name: "", amount: "", auth: "" });
  const [invoiceFile, setInvoiceFile] = useState<File | null>(null);
  const [invoiceFileName, setInvoiceFileName] = useState("");
  const [invoiceHeaders, setInvoiceHeaders] = useState<string[]>([]);
  const [invoiceRows, setInvoiceRows] = useState<Record<string, unknown>[]>([]);
  const [invoiceMap, setInvoiceMap] = useState({ name: "", amount: "" });
  const [results, setResults] = useState<StageCMatch[]>([]);
  const stageCDataLoaded = useRef(false);
  const [page, setPage] = useState(1);
  const [error, setError] = useState<string | null>(null);
  useEffect(() => {
    storageGet<any>("stage_c_platform_session", null).then(session => {
      if (session) {
        setSonyFileName(session.sonyFileName ?? "");
        setHeaders(session.headers ?? []);
        setRows(session.rows ?? []);
        setMap(session.map ?? { name: "", amount: "", auth: "" });
        setInvoiceFileName(session.invoiceFileName ?? "");
        setInvoiceHeaders(session.invoiceHeaders ?? []);
        setInvoiceRows(session.invoiceRows ?? []);
        setInvoiceMap(session.invoiceMap ?? { name: "", amount: "" });
        setResults(session.results ?? []);
      }
      stageCDataLoaded.current = true;
    });
  }, []);
  useEffect(() => {
    if (!stageCDataLoaded.current) return;
    void storageSet("stage_c_platform_session", {
      sonyFileName, headers, rows, map, invoiceFileName, invoiceHeaders, invoiceRows, invoiceMap, results,
    });
  }, [sonyFileName, headers, rows, map, invoiceFileName, invoiceHeaders, invoiceRows, invoiceMap, results]);
  const parseAmount = (value: unknown) => Number.parseFloat(String(value ?? "").replace(/[٫٬]/g, match => match === "٫" ? "." : "").replace(/,/g, "")) || 0;
  const loadInvoiceFile = async (nextFile: File) => {
    try {
      await saveStageASourceFiles("stageC", "invoice", [nextFile]);
      const parsed = parseSheet(await readFileBuf(nextFile));
      setInvoiceFile(nextFile);
      setInvoiceFileName(nextFile.name);
      setInvoiceHeaders(parsed.headers);
      setInvoiceRows(parsed.rows);
      setInvoiceMap({
        name: autoDetect(parsed.headers, HINTS.name),
        amount: detectCardPaidColumn(parsed.headers) || autoDetect(parsed.headers, ["paid", "مدفوع", "المبلغ المدفوع"]) || autoDetect(parsed.headers, HINTS.debit) || autoDetect(parsed.headers, ["amount", "مبلغ"])
      });
      setError(null);
    } catch (e) {
      setError((e as Error).message);
    }
  };
  const effectiveInvoices = useMemo(() => {
    if (invoices.length) return invoices;
    return invoiceRows.map((row, index) => {
      const originalName = String(row[invoiceMap.name] ?? "").trim();
      const name = originalName.replace(/(?:^|\D)\d{4}(?=\D|$)/, " ").replace(/\s+/g, " ").trim();
      return {
        id: index,
        source: "فاتورة فيزا مرفوعة مباشرة",
        name,
        originalName,
        amount: parseAmount(row[invoiceMap.amount]),
        databaseName: "",
        databaseAmount: 0,
        invoiceTotal: 0,
        cashAmount: 0,
        checkAmount: 0,
        debtAmount: 0,
        discountAmount: 0,
        type: "مدفوع" as const,
        userId: "",
        invoiceDate: "",
        registrationTime: "",
        customerNumber: "",
        invoiceNumber: stageAVisaNumber(originalName),
        accountType: "",
        invoice: row,
        matched: true,
        issue: "",
        category: "فيزا" as const,
        categoryIssue: ""
      };
    }).filter(invoice => invoice.name || invoice.amount);
  }, [invoices, invoiceRows, invoiceMap]);
  const loadSonyFile = async (nextFile: File) => {
    try {
      await saveStageASourceFiles("stageC", "sony", [nextFile]);
      const parsed = parseSheet(await readFileBuf(nextFile));
      setFile(nextFile); setHeaders(parsed.headers); setRows(parsed.rows);
      setSonyFileName(nextFile.name);
      setMap({ name: autoDetect(parsed.headers, HINTS.name), amount: detectCardPaidColumn(parsed.headers) || autoDetect(parsed.headers, HINTS.debit) || autoDetect(parsed.headers, ["amount", "مبلغ"]), auth: detectAuthorizationColumn(parsed.headers) });
      setError(null);
    } catch (e) { setError((e as Error).message); }
  };
  useEffect(() => {
    if (initialFile?.stage !== "stageC") return;
    if (initialFile.source === "invoice") void loadInvoiceFile(initialFile.file);
    else if (initialFile.source === "sony") void loadSonyFile(initialFile.file);
  }, [initialFile]);
  const runStageC = () => {
    if (!effectiveInvoices.length) { setError("رحّل فواتير الفيزا من A أو ارفع ملف فواتير الفيزا هنا أولاً"); return; }
    if (!rows.length || !map.name || !map.amount) { setError("ارفع ملف سوني كاشير وحدد عمود البيان والمبلغ أولاً"); return; }
    const sonyRows = rows.map((row, index) => ({ index, name: String(row[map.name] ?? "").trim(), amount: parseAmount(row[map.amount]), visa: stageAVisaNumber(String(row[map.name] ?? "")), auth: String(row[map.auth] ?? "").trim(), used: false }))
      .filter(row => row.name || row.amount || row.visa);
    const invoiceResults = effectiveInvoices.map(invoice => {
      const visa = stageAVisaNumber(invoice.originalName);
      const direct = visa ? sonyRows.find(row => !row.used && row.visa === visa && Math.abs(row.amount - invoice.amount) <= 0.01) : undefined;
      if (direct) { direct.used = true; return { kind: "invoice" as const, invoice, sonyName: direct.name, sonyAmount: direct.amount, sonyVisa: direct.visa, authNumber: direct.auth, matched: true, issue: "" }; }
      const group = visa ? sonyRows.filter(row => !row.used && row.visa === visa) : [];
      const grouped = group.length > 1 ? group.reduce((sum, row) => sum + row.amount, 0) : 0;
      if (grouped && Math.abs(grouped - invoice.amount) <= 0.01) { group.forEach(row => { row.used = true; }); return { kind: "invoice" as const, invoice, sonyName: group.map(row => row.name).join(" + "), sonyAmount: grouped, sonyVisa: visa, authNumber: group.map(row => row.auth).filter(Boolean).join(" + "), matched: true, issue: "مطابقة مجمعة" }; }
      return { kind: "invoice" as const, invoice, sonyName: "", sonyAmount: 0, sonyVisa: visa, authNumber: "", matched: false, issue: visa ? "رقم الفيزا موجود لكن الرقم أو المبلغ غير موجود في سوني" : "لم يتم العثور على رقم فيزا من 4 أرقام" };
    });
    const unmatchedSonyResults: StageCMatch[] = sonyRows.filter(row => !row.used).map(row => {
      const invoice: StageAResult = {
        id: row.index,
        source: "حركة سوني كاشير",
        name: row.name.replace(/(?:^|\D)\d{4}(?=\D|$)/, " ").replace(/\s+/g, " ").trim(),
        originalName: row.name,
        amount: row.amount,
        databaseName: "",
        databaseAmount: 0,
        invoiceTotal: 0,
        cashAmount: 0,
        checkAmount: 0,
        debtAmount: 0,
        discountAmount: 0,
        type: "مدفوع",
        userId: "",
        invoiceDate: "",
        registrationTime: "",
        customerNumber: "",
        invoiceNumber: row.visa,
        accountType: "",
        invoice: rows[row.index],
        matched: false,
        issue: "حركة سوني لا تقابل فاتورة فيزا",
        category: "فيزا",
        categoryIssue: ""
      };
      return {
        kind: "sony-only",
        invoice,
        sonyName: row.name,
        sonyAmount: row.amount,
        sonyVisa: row.visa,
        authNumber: row.auth,
        matched: false,
        issue: "حركة سوني زائدة بلا فاتورة مطابقة"
      };
    });
    setResults([...invoiceResults, ...unmatchedSonyResults]);
    setPage(1);
    setError(null);
  };
  const pageSize = 20;
  const totalPages = Math.max(1, Math.ceil(results.length / pageSize));
  const currentPage = Math.min(page, totalPages);
  const pagedResults = results.slice((currentPage - 1) * pageSize, currentPage * pageSize);
  const exportStageC = () => {
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([
      ["نوع السجل", "اسم الزبون / بيان سوني", "رقم الفيزا", "مبلغ الفاتورة", "تفويض سوني", "بيان سوني", "مبلغ سوني", "الحالة", "ملاحظة"],
      ...results.map(result => [
        result.kind === "invoice" ? "فاتورة فيزا" : "حركة سوني زائدة",
        result.kind === "invoice" ? result.invoice.name : result.sonyName,
        result.sonyVisa || stageAVisaNumber(result.invoice.originalName),
        result.kind === "invoice" ? result.invoice.amount : "",
        result.authNumber,
        result.sonyName,
        result.sonyAmount || "",
        result.kind === "sony-only" ? "حركة سوني بلا فاتورة" : result.matched ? "مطابق" : "فاتورة بلا حركة سوني مطابقة",
        result.invoice.categoryIssue || result.issue
      ])
    ]);
    sheet["!cols"] = [18, 32, 14, 16, 18, 32, 16, 30, 46].map(wch => ({ wch }));
    sheet["!autofilter"] = { ref: sheet["!ref"] || "A1:I1" };
    sheet["!rows"] = [{ hpt: 26 }];
    XLSX.utils.book_append_sheet(workbook, sheet, "مطابقة فيزا");
    XLSX.writeFile(workbook, "مطابقة_الفيزا_مع_سوني.xlsx");
  };
  return <div dir="rtl" className="min-h-screen bg-slate-50 p-6 text-slate-900"><div className="mx-auto max-w-6xl space-y-5">
    <div className="flex flex-wrap items-center justify-between gap-3">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-blue-600"><ChevronLeft className="h-4 w-4" />العودة لمنصة الفيزا</button>
      <h1 className="text-xl font-black">C — مطابقة الفواتير مع سوني كاشير فيزا</h1>
      <div className="flex gap-2">
        <button onClick={runStageC} className="rounded-lg bg-amber-600 px-4 py-2 text-sm font-medium text-white"><Sparkles className="mr-1 inline h-4 w-4" />تشغيل المطابقة</button>
        {results.length > 0 && <button onClick={exportStageC} className="rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white"><Download className="mr-1 inline h-4 w-4" />تصدير مطابقة فيزا</button>}
        {results.length > 0 && <button onClick={() => {
          const confirmed = results.filter(result => result.matched);
          onSafeActivities?.(confirmed.map(result => `${result.invoice.name} ↔ ${result.sonyName} · ${fmtNum(result.invoice.amount)} ₪`));
          onTransferToD?.(results, pendingBankInputs);
        }} className="rounded-lg bg-indigo-600 px-4 py-2 text-sm font-medium text-white">ترحيل إلى D</button>}
      </div>
    </div>
    <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">{invoices.length ? `تم ترحيل ${invoices.length} فاتورة من A.` : `${effectiveInvoices.length} فاتورة جاهزة من الملف المرفوع مباشرة.`} تتم المطابقة برقم الفيزا المكوّن من 4 أرقام والمبلغ، ويُحفظ رقم التفويض من صف سوني نفسه.</div>
    {invoices.length > 0 && <div className="flex justify-end"><button type="button" onClick={() => onTransferToB(invoices)} className="rounded-lg bg-emerald-600 px-4 py-2 text-sm font-semibold text-white hover:bg-emerald-700">ترحيل فواتير منصة C إلى B ({invoices.length})</button></div>}
    {error && <div className="rounded-lg bg-red-100 p-3 text-sm text-red-800">{error}</div>}
    {!invoices.length && <div className="rounded-xl border border-slate-200 bg-white p-4">
      <h2 className="mb-3 text-sm font-bold">ملف فواتير الفيزا (إذا لم تُرحّل من A)</h2>
      <DropZone file={invoiceFile} rowCount={invoiceRows.length} onFile={loadInvoiceFile} onClear={() => { setInvoiceFile(null); setInvoiceFileName(""); setInvoiceHeaders([]); setInvoiceRows([]); }} />
      {!invoiceFile && invoiceFileName && <p className="mt-2 text-xs text-emerald-700">تم استعادة بيانات الملف المحفوظ: {invoiceFileName}</p>}
      {!!invoiceHeaders.length && <div className="mt-3 grid grid-cols-1 gap-2 sm:grid-cols-2">
        <Sel label="اسم الزبون مع رقم الفيزا" headers={invoiceHeaders} value={invoiceMap.name} onChange={value => setInvoiceMap(prev => ({ ...prev, name: value }))} />
        <Sel label="مبلغ الفاتورة" headers={invoiceHeaders} value={invoiceMap.amount} onChange={value => setInvoiceMap(prev => ({ ...prev, amount: value }))} />
      </div>}
    </div>}
    <div className="rounded-xl border border-slate-200 bg-white p-4">
      <h2 className="mb-3 text-sm font-bold">ملف سوني كاشير فيزا</h2>
      <DropZone file={file} rowCount={rows.length} onFile={loadSonyFile} onClear={() => { setFile(null); setSonyFileName(""); setHeaders([]); setRows([]); }} />
      {!file && sonyFileName && <p className="mt-2 text-xs text-emerald-700">تم استعادة بيانات الملف المحفوظ: {sonyFileName}</p>}
      {!!headers.length && <div className="mt-3 grid grid-cols-1 gap-2 sm:grid-cols-3">
        <Sel label="البيان الذي يحتوي رقم الفيزا" headers={headers} value={map.name} onChange={value => setMap(prev => ({ ...prev, name: value }))} />
        <Sel label="المبلغ" headers={headers} value={map.amount} onChange={value => setMap(prev => ({ ...prev, amount: value }))} />
        <Sel label="رقم التفويض" headers={headers} value={map.auth} onChange={value => setMap(prev => ({ ...prev, auth: value }))} />
      </div>}
    </div>
    {!!results.length && <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
      <div className="border-b bg-slate-50 px-4 py-3 text-sm font-bold">
        فواتير مطابقة: {results.filter(result => result.kind === "invoice" && result.matched).length}
        {" · "}فواتير بلا مطابقة: {results.filter(result => result.kind === "invoice" && !result.matched).length}
        {" · "}حركات سوني زائدة: {results.filter(result => result.kind === "sony-only").length}
      </div>
      <div className="overflow-x-auto"><table className="w-full text-sm"><thead className="bg-slate-100 text-xs"><tr><th className="p-3 text-right">نوع السجل</th><th className="p-3 text-right">اسم الزبون / بيان سوني</th><th className="p-3 text-right">رقم الفيزا</th><th className="p-3 text-right">مبلغ الفاتورة</th><th className="p-3 text-right">مبلغ سوني</th><th className="p-3 text-right">رقم التفويض</th><th className="p-3 text-right">بيان سوني</th><th className="p-3 text-right">الحالة</th></tr></thead><tbody>
        {pagedResults.map(result => <tr key={`${result.kind}-${result.invoice.source}-${result.invoice.id}`} className="border-t">
          <td className="p-3">{result.kind === "invoice" ? "فاتورة فيزا" : "حركة سوني زائدة"}</td>
          <td className="p-3">{result.kind === "invoice" ? result.invoice.name : result.sonyName}</td>
          <td className="p-3 font-mono">{result.sonyVisa || stageAVisaNumber(result.invoice.originalName) || "—"}</td>
          <td className="p-3 font-mono">{result.kind === "invoice" ? fmtNum(result.invoice.amount) : "—"}</td>
          <td className="p-3 font-mono">{result.sonyName ? fmtNum(result.sonyAmount) : "—"}</td>
          <td className="p-3 font-mono">{result.authNumber || "—"}</td>
          <td className="p-3">{result.sonyName || "—"}</td>
          <td className={`p-3 font-semibold ${result.matched ? "text-green-700" : "text-red-700"}`}>
            {result.kind === "sony-only" ? "حركة سوني بلا فاتورة مطابقة" : result.matched ? result.issue || "مطابق" : "فاتورة بلا حركة سوني مطابقة"}
          </td>
        </tr>)}
      </tbody></table></div>
      {totalPages > 1 && <div className="flex items-center justify-center gap-3 border-t p-3 text-xs">
        <button type="button" onClick={() => setPage(current => Math.max(1, current - 1))} disabled={currentPage <= 1} className="rounded border px-3 py-1.5 disabled:opacity-40">السابق</button>
        <span>صفحة {currentPage} من {totalPages}</span>
        <button type="button" onClick={() => setPage(current => Math.min(totalPages, current + 1))} disabled={currentPage >= totalPages} className="rounded border px-3 py-1.5 disabled:opacity-40">التالي</button>
      </div>}
    </div>}
  </div></div>;
}

// ─── Main Component ───────────────────────────────────────────────────────────
export default function Reconciliation2({ onBack, onStageAToMain, onStageBToMain, onSafeActivities, onOpenInvoiceReference, visaInvoices = [], initialStageDBankInputs = [], initialStage = "hub", projectSourceFiles = [] }: { onBack: () => void; onStageAToMain?: (results: StageAResult[], summary?: StageASummary) => void; onStageBToMain?: () => void; onSafeActivities?: (activities: string[]) => void; onOpenInvoiceReference?: (invoice: StageAResult) => boolean; visaInvoices?: StageAResult[]; initialStageDBankInputs?: StageDBankInput[]; initialStage?: "hub" | "stageA" | "stageB" | "stageC" | "stageD" | "stageF"; projectSourceFiles?: SavedProjectSourceFile[] }) {
  const [stagePage, setStagePage] = useState<"hub" | "stageA" | "stageB" | "stageC" | "stageD" | "stageF">(initialStage);
  const [stageDMatches, setStageDMatches] = useState<StageCMatch[]>([]);
  const [stageDPending, setStageDPending] = useState<StageDBankInput[]>(initialStageDBankInputs);
  const [stageBInvoices, setStageBInvoices] = useState<StageAResult[]>([]);
  const [stageCInvoices, setStageCInvoices] = useState<StageAResult[]>(visaInvoices);
  const [transferDialog, setTransferDialog] = useState<{ invoices: StageAResult[]; from: "B" | "C" } | null>(null);
  const [projectFileToOpen, setProjectFileToOpen] = useState<ProjectSourceFile | null>(null);
  const openProjectFile = (file: ProjectSourceFile) => {
    setProjectFileToOpen(file);
    setStagePage(file.stage);
  };
  const openTransferDialog = (invoices: StageAResult[], from: "B" | "C") => setTransferDialog({ invoices, from });
  const stageCSessionLoaded = useRef(false);
  const [sessionLoaded, setSessionLoaded] = useState(false);

  useEffect(() => {
    storageGet<StageAResult[]>("stage_c_invoices", []).then(saved => {
      if (visaInvoices.length) void storageSet("stage_c_invoices", visaInvoices);
      else if (saved.length) setStageCInvoices(saved);
      stageCSessionLoaded.current = true;
    });
  }, []);
  useEffect(() => {
    if (stageCSessionLoaded.current) void storageSet("stage_c_invoices", stageCInvoices);
  }, [stageCInvoices]);

  // ─── File states ──────────────────────────────────────────────────────────
  // 1. كشف البنك (يُفرز منه مشتريات عمولة تجار تلقائياً لخانة "فيزا بنك")
  const [bankFile, setBankFile] = useState<File|null>(null);
  const [bankHeaders, setBankHeaders] = useState<string[]>([]);
  const [bankRowsRaw, setBankRowsRaw] = useState<Record<string,unknown>[]>([]);
  const [bankMap, setBankMap] = useState({ date:"", desc:"", debit:"", credit:"", accountType:"", reference:"" });

  // 2. كشف الأستاذ (الاستاذ) — الفواتير الكاملة (~1050)
  const [azaFile, setAzaFile] = useState<File|null>(null);
  const [azaHeaders, setAzaHeaders] = useState<string[]>([]);
  const [azaRowsRaw, setAzaRowsRaw] = useState<Record<string,unknown>[]>([]);
  const [azaMap, setAzaMap] = useState({ date:"", name:"", debit:"", credit:"", accountType:"", reference:"" });

  // 3. كشف اليومي كاشير — حوالات يوم كامل (~1000)
  const [dailyFile, setDailyFile] = useState<File|null>(null);
  const [dailyHeaders, setDailyHeaders] = useState<string[]>([]);
  const [dailyRowsRaw, setDailyRowsRaw] = useState<Record<string,unknown>[]>([]);
  const [dailyMap, setDailyMap] = useState({ date:"", name:"", debit:"", credit:"", accountType:"", reference:"" });

  // 4. سوني كاشير فيزا — اسم الزبون + رقم 4 أرقام + مبلغ
  const [sonyFile, setSonyFile] = useState<File|null>(null);
  const [sonyHeaders, setSonyHeaders] = useState<string[]>([]);
  const [sonyRowsRaw, setSonyRowsRaw] = useState<Record<string,unknown>[]>([]);
  const [sonyMap, setSonyMap] = useState({ date:"", name:"", debit:"", credit:"", accountType:"", authNum:"" });

  // 5. رقم التفويض فيزا — ملف أرقام التفويض
  const [authFile, setAuthFile] = useState<File|null>(null);
  const [authHeaders, setAuthHeaders] = useState<string[]>([]);
  const [authRowsRaw, setAuthRowsRaw] = useState<Record<string,unknown>[]>([]);
  const [authMap, setAuthMap] = useState({ name:"", authNum:"", amount:"", date:"" });

  // ─── Matched & pending state ──────────────────────────────────────────────
  const [matchedVisa, setMatchedVisa] = useState<MatchedVisa[]>([]);
  const [pendingVisaBank, setPendingVisaBank] = useState<PendingVisaBank[]>([]);
  const [rejectedPairs, setRejectedPairs] = useState<Set<string>>(new Set());
  const [tab, setTab] = useState<R2Tab>("overview");
  const [error, setError] = useState<string|null>(null);
  const [expandedKey, setExpandedKey] = useState<string|null>(null);
  const [visaSwaps, setVisaSwaps] = useState<Record<string, boolean>>({ bank: false, aza: false, daily: false, sony: false, auth: false });

  // ─── لصق إيصال SoftPOS وتحويله لملف سوني كاشير فيزا ──────────────────────
  const [softposText, setSoftposText] = useState("");

  // ─── Session restore ──────────────────────────────────────────────────────
  useEffect(() => {
    (async () => {
      const s = await storageGet<any>("recon2_session", null);
      if (s) {
        setBankHeaders(s.bankHeaders || []); setBankRowsRaw(s.bankRowsRaw || []);
        setBankMap({ date:"", desc:"", debit:"", credit:"", accountType:"", reference:"", ...(s.bankMap||{}) });
        setAzaHeaders(s.azaHeaders || []); setAzaRowsRaw(s.azaRowsRaw || []);
        setAzaMap({ date:"", name:"", debit:"", credit:"", accountType:"", reference:"", ...(s.azaMap||{}) });
        setDailyHeaders(s.dailyHeaders || []); setDailyRowsRaw(s.dailyRowsRaw || []);
        setDailyMap({ date:"", name:"", debit:"", credit:"", accountType:"", reference:"", ...(s.dailyMap||{}) });
        setSonyHeaders(s.sonyHeaders || []); setSonyRowsRaw(s.sonyRowsRaw || []);
        setSonyMap({ date:"", name:"", debit:"", credit:"", accountType:"", authNum:"", ...(s.sonyMap||{}) });
        setAuthHeaders(s.authHeaders || []); setAuthRowsRaw(s.authRowsRaw || []);
        setAuthMap({ name:"", authNum:"", amount:"", date:"", ...(s.authMap||{}) });
        setMatchedVisa(s.matchedVisa || []);
        setPendingVisaBank(s.pendingVisaBank || []);
        setRejectedPairs(new Set(s.rejectedPairs || []));
        if (s.visaSwaps) setVisaSwaps(s.visaSwaps);
      }
      setSessionLoaded(true);
    })();
  }, []);

  // ─── Auto-save ─────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!sessionLoaded) return;
    const t = setTimeout(() => {
      storageSet("recon2_session", {
        bankHeaders, bankRowsRaw, bankMap,
        azaHeaders, azaRowsRaw, azaMap,
        dailyHeaders, dailyRowsRaw, dailyMap,
        sonyHeaders, sonyRowsRaw, sonyMap,
        authHeaders, authRowsRaw, authMap,
        visaSwaps, matchedVisa, pendingVisaBank,
        rejectedPairs: Array.from(rejectedPairs),
      });
    }, 500);
    return () => clearTimeout(t);
  }, [sessionLoaded, bankHeaders, bankRowsRaw, bankMap,
      azaHeaders, azaRowsRaw, azaMap,
      dailyHeaders, dailyRowsRaw, dailyMap,
      sonyHeaders, sonyRowsRaw, sonyMap,
      authHeaders, authRowsRaw, authMap,
      visaSwaps, matchedVisa, pendingVisaBank, rejectedPairs]);

  // ─── File loaders ──────────────────────────────────────────────────────────
  const loadBank = async (f: File) => {
    setBankFile(f); setError(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setBankHeaders(headers); setBankRowsRaw(rows);
      setBankMap({ date:autoDetect(headers,HINTS.date), desc:autoDetect(headers,HINTS.desc),
        debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit),
        accountType:autoDetect(headers,HINTS.accountType), reference:autoDetect(headers,HINTS.authNum) });
    } catch(e) { setError((e as Error).message); }
  };
  const loadAza = async (f: File) => {
    setAzaFile(f); setError(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setAzaHeaders(headers); setAzaRowsRaw(rows);
      setAzaMap({ date:autoDetect(headers,HINTS.date), name:autoDetect(headers,HINTS.name),
        debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit),
        accountType:autoDetect(headers,HINTS.accountType), reference:autoDetect(headers,HINTS.authNum) });
    } catch(e) { setError((e as Error).message); }
  };
  const loadDaily = async (f: File) => {
    setDailyFile(f); setError(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setDailyHeaders(headers); setDailyRowsRaw(rows);
      setDailyMap({ date:autoDetect(headers,HINTS.date), name:autoDetect(headers,HINTS.name),
        debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit),
        accountType:autoDetect(headers,HINTS.accountType), reference:autoDetect(headers,HINTS.authNum) });
    } catch(e) { setError((e as Error).message); }
  };
  const loadSony = async (f: File) => {
    setSonyFile(f); setError(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setSonyHeaders(headers); setSonyRowsRaw(rows);
      setSonyMap({ date:autoDetect(headers,HINTS.date), name:autoDetect(headers,HINTS.name),
        debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit),
        accountType:autoDetect(headers,HINTS.accountType), authNum:autoDetect(headers,HINTS.authNum) });
    } catch(e) { setError((e as Error).message); }
  };
  const loadAuth = async (f: File) => {
    setAuthFile(f); setError(null);
    try {
      const { headers, rows } = parseSheet(await readFileBuf(f));
      setAuthHeaders(headers); setAuthRowsRaw(rows);
      setAuthMap({ name:autoDetect(headers,HINTS.name), authNum:autoDetect(headers,HINTS.authNum),
        amount:autoDetect(headers,HINTS.debit), date:autoDetect(headers,HINTS.date) });
    } catch(e) { setError((e as Error).message); }
  };

  // ─── Parse bank rows + auto-separate مشتريات عمولة تجار ─────────────────────
  const isVisaBankDesc = useCallback((desc: string): boolean => {
    return isMerchantPurchaseBankDescription(desc);
  }, []);

  const parsedBank = useMemo((): BankRow2[] => {
    const swap = visaSwaps.bank;
    return bankRowsRaw.map((r, i) => {
      const desc = String(bankMap.desc ? r[bankMap.desc] : "").trim();
      const dRaw = Math.abs(toNum(bankMap.debit ? r[bankMap.debit] : 0));
      const cRaw = Math.abs(toNum(bankMap.credit ? r[bankMap.credit] : 0));
      const debitRaw = swap ? cRaw : dRaw;
      const creditRaw = swap ? dRaw : cRaw;
      let rawAmount = 0;
      let type: "مدفوع" | "مستلم" = "مستلم";
      if (debitRaw > 0 && creditRaw === 0) { rawAmount = debitRaw; type = "مدفوع"; }
      else if (creditRaw > 0 && debitRaw === 0) { rawAmount = creditRaw; type = "مستلم"; }
      else if (debitRaw > 0 && creditRaw > 0) { rawAmount = debitRaw; type = "مدفوع"; }
      if (rawAmount === 0) return null;
      return {
        id: i,
        date: fmtDate(bankMap.date ? r[bankMap.date] : ""),
        description: desc,
        rawAmount,
        type,
        accountType: String(bankMap.accountType ? r[bankMap.accountType] : "").trim(),
        isVisaBank: isVisaBankDesc(desc),
        orig: r,
      };
    }).filter((r): r is BankRow2 => r !== null);
  }, [bankRowsRaw, bankMap, isVisaBankDesc, visaSwaps.bank]);

  // فيزا بنك = مشتريات عمولة تجار المفرزة تلقائياً
  const visaBankRows = useMemo(() => parsedBank.filter(b => b.isVisaBank), [parsedBank]);
  const regularBankRows = useMemo(() => parsedBank.filter(b => !b.isVisaBank), [parsedBank]);

  // ─── Parse أستاذ rows ──────────────────────────────────────────────────────
  const parsedAza = useMemo((): AzaRow[] => {
    const swap = visaSwaps.aza;
    return azaRowsRaw.map((r, i) => {
      const rawName = String(azaMap.name ? r[azaMap.name] : "").trim();
      const name = rawName.replace(VISA_NUM_RE, "").replace(/\s+/g, " ").trim();
      const dRaw = Math.abs(toNum(azaMap.debit ? r[azaMap.debit] : 0));
      const cRaw = Math.abs(toNum(azaMap.credit ? r[azaMap.credit] : 0));
      const debitRaw = swap ? cRaw : dRaw;
      const creditRaw = swap ? dRaw : cRaw;
      let amount = 0;
      let type: "مدفوع" | "مستلم" = "مستلم";
      if (debitRaw > 0 && creditRaw === 0) { amount = debitRaw; type = "مدفوع"; }
      else if (creditRaw > 0 && debitRaw === 0) { amount = creditRaw; type = "مستلم"; }
      else if (debitRaw > 0 && creditRaw > 0) { amount = debitRaw; type = "مدفوع"; }
      if (amount === 0) return null;
      return {
        id: i, rawName, name, amount, type,
        accountType: String(azaMap.accountType ? r[azaMap.accountType] : "").trim(),
        date: fmtDate(azaMap.date ? r[azaMap.date] : ""),
        visaNumber: extractVisaNumber(rawName),
        orig: r,
      };
    }).filter((r): r is AzaRow => r !== null);
  }, [azaRowsRaw, azaMap, visaSwaps.aza]);

  // أستاذ فيزا = اللي فيها رقم فيزا (4 أرقام)
  const azaVisaRows = useMemo(() => parsedAza.filter(a => a.visaNumber), [parsedAza]);
  const azaRegularRows = useMemo(() => parsedAza.filter(a => !a.visaNumber), [parsedAza]);

  // ─── Parse daily cashier rows ──────────────────────────────────────────────
  const parsedDaily = useMemo((): AzaRow[] => {
    const swap = visaSwaps.daily;
    return dailyRowsRaw.map((r, i) => {
      const rawName = String(dailyMap.name ? r[dailyMap.name] : "").trim();
      const name = rawName.replace(VISA_NUM_RE, "").replace(/\s+/g, " ").trim();
      const dRaw = Math.abs(toNum(dailyMap.debit ? r[dailyMap.debit] : 0));
      const cRaw = Math.abs(toNum(dailyMap.credit ? r[dailyMap.credit] : 0));
      const debitRaw = swap ? cRaw : dRaw;
      const creditRaw = swap ? dRaw : cRaw;
      let amount = 0;
      let type: "مدفوع" | "مستلم" = "مستلم";
      if (debitRaw > 0 && creditRaw === 0) { amount = debitRaw; type = "مدفوع"; }
      else if (creditRaw > 0 && debitRaw === 0) { amount = creditRaw; type = "مستلم"; }
      else if (debitRaw > 0 && creditRaw > 0) { amount = debitRaw; type = "مدفوع"; }
      if (amount === 0) return null;
      return {
        id: i, rawName, name, amount, type,
        accountType: String(dailyMap.accountType ? r[dailyMap.accountType] : "").trim(),
        date: fmtDate(dailyMap.date ? r[dailyMap.date] : ""),
        visaNumber: extractVisaNumber(rawName),
        orig: r,
      };
    }).filter((r): r is AzaRow => r !== null);
  }, [dailyRowsRaw, dailyMap, visaSwaps.daily]);

  // الفرق: أستاذ - يومي = الـ50 فاتورة الزيادة (مش موجودة باليومي)
  const azaOnlyRows = useMemo(() => {
    const dailyNames = new Set(parsedDaily.map(d => normName(d.name)));
    return azaRegularRows.filter(a => !dailyNames.has(normName(a.name)));
  }, [azaRegularRows, parsedDaily]);

  // ─── Parse سوني كاشير فيزا ─────────────────────────────────────────────────
  const parsedSony = useMemo((): SonyVisaRow[] => {
    const swap = visaSwaps.sony;
    return sonyRowsRaw.map((r, i) => {
      const rawName = String(sonyMap.name ? r[sonyMap.name] : "").trim();
      const visaNum = extractVisaNumber(rawName);
      const name = rawName.replace(VISA_NUM_RE, "").replace(/\s+/g, " ").trim();
      const dRaw = Math.abs(toNum(sonyMap.debit ? r[sonyMap.debit] : 0));
      const cRaw = Math.abs(toNum(sonyMap.credit ? r[sonyMap.credit] : 0));
      const debitRaw = swap ? cRaw : dRaw;
      const creditRaw = swap ? dRaw : cRaw;
      let amount = 0;
      if (debitRaw > 0) amount = debitRaw;
      else if (creditRaw > 0) amount = creditRaw;
      if (amount === 0 && !visaNum) return null;
      const authNum = sonyMap.authNum ? String(r[sonyMap.authNum] || "").trim() : null;
      return {
        id: i, rawName, name, amount, visaNumber: visaNum, authNumber: authNum || null,
        date: fmtDate(sonyMap.date ? r[sonyMap.date] : ""), orig: r,
      };
    }).filter((r): r is SonyVisaRow => r !== null);
  }, [sonyRowsRaw, sonyMap, visaSwaps.sony]);

  // ─── Parse رقم التفويض ─────────────────────────────────────────────────────
  const parsedAuth = useMemo(() => {
    return authRowsRaw.map((r, i) => {
      const name = String(authMap.name ? r[authMap.name] : "").trim();
      const authNum = String(authMap.authNum ? r[authMap.authNum] : "").trim();
      const amount = toNum(authMap.amount ? r[authMap.amount] : 0);
      return { id: i, name, authNumber: authNum, amount, orig: r };
    }).filter(a => a.authNumber || a.name);
  }, [authRowsRaw, authMap]);

  // ─── Matched IDs ───────────────────────────────────────────────────────────
  const matchedSonyIds = useMemo(() => new Set(matchedVisa.map(m => m.sonyId)), [matchedVisa]);
  const matchedAzaIds = useMemo(() => new Set(matchedVisa.map(m => m.azaId)), [matchedVisa]);

  // ─── Auto-match: سوني كاشير فيزا ↔ أستاذ فيزا ──────────────────────────────
  // المطابقة بالاسم + المبلغ + رقم الفيزا (إن وجد)
  const autoMatchResults = useMemo(() => {
    const results: Array<{
      sony: SonyVisaRow; aza: AzaRow; score: number; matchType: string;
      amountDiff: number; visaMatch: boolean;
    }> = [];
    const usedAza = new Set<number>(matchedAzaIds);
    const usedSony = new Set<number>(matchedSonyIds);

    parsedSony.forEach(s => {
      if (usedSony.has(s.id)) return;
      let best: { aza: AzaRow; score: number; matchType: string; amountDiff: number; visaMatch: boolean } | null = null;

      azaVisaRows.forEach(a => {
        if (usedAza.has(a.id)) return;
        const pairKey = `${s.id}-${a.id}`;
        if (rejectedPairs.has(pairKey)) return;

        const amtMatch = Math.abs(a.amount - s.amount) <= 0.01;
        const ms = advancedMatchCheck(s.name, a.name);
        let visaMatch = false;
        if (s.visaNumber && a.visaNumber) visaMatch = s.visaNumber === a.visaNumber;

        let score = 0;
        if (ms.isMatch && amtMatch) score = 0.95;
        else if (ms.isMatch && !amtMatch) score = 0.5 + nameSim(s.name, a.name) * 0.3;
        else if (!ms.isMatch && amtMatch) score = 0.2;
        if (visaMatch) score += 0.1;
        if (score < 0.15) return;

        if (!best || score > best.score) {
          best = { aza: a, score, matchType: ms.matchType, amountDiff: Math.abs(a.amount - s.amount), visaMatch };
        }
      });

      const selected = best as { aza: AzaRow; score: number; matchType: string; amountDiff: number; visaMatch: boolean } | null;
      if (selected && selected.score >= 0.15) {
        results.push({ sony: s, aza: selected.aza, score: selected.score, matchType: selected.matchType, amountDiff: selected.amountDiff, visaMatch: selected.visaMatch });
      }
    });

    return results.sort((a, b) => b.score - a.score);
  }, [parsedSony, azaVisaRows, rejectedPairs, matchedSonyIds, matchedAzaIds]);

  // ─── Unmatched سوني ─────────────────────────────────────────────────────────
  const unmatchedSony = useMemo(() => {
    const matchedIds = new Set<number>([...matchedSonyIds, ...autoMatchResults.map(r => r.sony.id)]);
    return parsedSony.filter(s => !matchedIds.has(s.id));
  }, [parsedSony, matchedSonyIds, autoMatchResults]);

  // ─── Unmatched أستاذ فيزا ──────────────────────────────────────────────────
  const unmatchedAzaVisa = useMemo(() => {
    const matchedIds = new Set<number>([...matchedAzaIds, ...autoMatchResults.map(r => r.aza.id)]);
    return azaVisaRows.filter(a => !matchedIds.has(a.id));
  }, [azaVisaRows, matchedAzaIds, autoMatchResults]);

  // ─── Handlers ──────────────────────────────────────────────────────────────
  const handleAcceptMatch = (sony: SonyVisaRow, aza: AzaRow, note?: string) => {
    const pairKey = `${sony.id}-${aza.id}`;
    if (matchedSonyIds.has(sony.id) || matchedAzaIds.has(aza.id)) return;
    const m: MatchedVisa = {
      id: `mv-${Date.now()}-${sony.id}-${aza.id}`,
      sonyId: sony.id, azaId: aza.id,
      sonyName: sony.name, azaName: aza.name,
      visaNumber: sony.visaNumber || aza.visaNumber || "",
      authNumber: sony.authNumber,
      amount: aza.amount,
      sonyRawName: sony.rawName, azaRawName: aza.rawName,
      matchedAt: new Date().toLocaleString("ar-SA"),
      note,
    };
    setMatchedVisa(prev => [...prev, m]);
    onSafeActivities?.([`${sony.name} ↔ ${aza.name} · ${fmtNum(aza.amount)} ₪`]);
    const next = new Set(rejectedPairs);
    next.delete(pairKey);
    setRejectedPairs(next);
    setExpandedKey(null);
  };

  const handleRejectMatch = (sonyId: number, azaId: number) => {
    const pairKey = `${sonyId}-${azaId}`;
    setRejectedPairs(prev => new Set(prev).add(pairKey));
    setExpandedKey(null);
  };

  const handleUnmatch = (m: MatchedVisa) => {
    if (!window.confirm("إلغاء هذه المطابقة؟")) return;
    setMatchedVisa(prev => prev.filter(x => x.id !== m.id));
  };

  const handleClearSaved = () => {
    if (!matchedVisa.length && !pendingVisaBank.length) return;
    if (!window.confirm("مسح كل المطابقات المحفوظة والمعلقات في منصة الفيزا؟")) return;
    setMatchedVisa([]);
    setPendingVisaBank([]);
    setRejectedPairs(new Set());
    setExpandedKey(null);
  };

  const handleClearSession = () => {
    if (!window.confirm("مسح جلسة منصة الفيزا كاملة؟ سيتم حذف الملفات والنتائج المحفوظة.")) return;
    setBankFile(null); setBankHeaders([]); setBankRowsRaw([]);
    setAzaFile(null); setAzaHeaders([]); setAzaRowsRaw([]);
    setDailyFile(null); setDailyHeaders([]); setDailyRowsRaw([]);
    setSonyFile(null); setSonyHeaders([]); setSonyRowsRaw([]);
    setAuthFile(null); setAuthHeaders([]); setAuthRowsRaw([]);
    setMatchedVisa([]); setPendingVisaBank([]); setRejectedPairs(new Set());
    setSoftposText(""); setError(null); setTab("overview");
    void storageSet("recon2_session", null);
  };

  const handleAcceptAllAuto = () => {
    if (!autoMatchResults.length) return;
    if (!window.confirm(`تأكيد ${autoMatchResults.length} مطابقة تلقائية؟`)) return;
    const newMatches: MatchedVisa[] = [];
    const usedAzaNow = new Set(matchedAzaIds);
    const usedSonyNow = new Set(matchedSonyIds);
    autoMatchResults.forEach(r => {
      if (usedSonyNow.has(r.sony.id) || usedAzaNow.has(r.aza.id)) return;
      usedSonyNow.add(r.sony.id);
      usedAzaNow.add(r.aza.id);
      newMatches.push({
        id: `mv-${Date.now()}-${r.sony.id}-${r.aza.id}`,
        sonyId: r.sony.id, azaId: r.aza.id,
        sonyName: r.sony.name, azaName: r.aza.name,
        visaNumber: r.sony.visaNumber || r.aza.visaNumber || "",
        authNumber: r.sony.authNumber,
        amount: r.aza.amount,
        sonyRawName: r.sony.rawName, azaRawName: r.aza.rawName,
        matchedAt: new Date().toLocaleString("ar-SA"),
        note: r.amountDiff > 0.01 ? `فرق مبلغ: ${fmtNum(r.amountDiff)}` : undefined,
      });
    });
    setMatchedVisa(prev => [...prev, ...newMatches]);
    onSafeActivities?.(newMatches.map(match => `${match.sonyName} ↔ ${match.azaName} · ${fmtNum(match.amount)} ₪`));
  };

  // ─── نقل مشتريات عمولة تجار للمعلقة (للمطابقة ثاني يوم) ──────────────────────
  const handleMoveVisaBankToPending = () => {
    if (!visaBankRows.length) return;
    const newRows = visaBankRows.filter(bank => !pendingVisaBank.some(item => item.bankId === bank.id));
    if (!newRows.length) return;
    if (!window.confirm(`نقل ${newRows.length} حوالة "مشتريات عمولة تجار" للمعلقة؟\nستطابقها لاحقاً مع تفويض سوني كاشير.`)) return;
    const items: PendingVisaBank[] = newRows.map(b => ({
      id: `pb-${Date.now()}-${b.id}`,
      bankId: b.id,
      bankDesc: b.description,
      authNumber: extractMerchantAuthorizationNumber(b.description) || null,
      amount: b.rawAmount,
      date: b.date,
      movedAt: new Date().toLocaleString("ar-SA"),
    }));
    setPendingVisaBank(prev => [...prev, ...items]);
  };

  // ─── مطابقة المعلقة مع سوني مؤكد ─────────────────────────────────────────────
  // ثاني يوم: نطابق pendingVisaBank (مشتريات البنك) مع matchedVisa (سوني مؤكد)
  // عبر رقم التفويض الموجود في وصف مشتريات التجار
  const pendingMatchResults = useMemo(() => {
    if (!pendingVisaBank.length || !matchedVisa.length) return [];
    const results: Array<{ pending: PendingVisaBank; matched: MatchedVisa; score: number; reason: string }> = [];
    const usedMatched = new Set<string>();

    pendingVisaBank.forEach(p => {
      let best: { matched: MatchedVisa; score: number; reason: string } | null = null;
      matchedVisa.forEach(m => {
        if (usedMatched.has(m.id)) return;
        if (m.authNumber && p.authNumber &&
            normalizeAuthorizationNumber(m.authNumber) === normalizeAuthorizationNumber(p.authNumber) &&
            (!best || 1 > best.score)) {
          best = { matched: m, score: 1, reason: "رقم التفويض مطابق" };
        }
      });
      const selected = best as { matched: MatchedVisa; score: number; reason: string } | null;
      if (selected) {
        results.push({ pending: p, matched: selected.matched, score: selected.score, reason: selected.reason });
        usedMatched.add(selected.matched.id);
      }
    });
    return results.sort((a, b) => b.score - a.score);
  }, [pendingVisaBank, matchedVisa]);

  // ─── لصق إيصال SoftPOS وتحويله لملف سوني كاشير فيزا ──────────────────────
  // النص بيحتوي على عدة حركات، كل حركة فيها:
  //   - رقم البطاقة (****-4351) → رقم الفيزا (4 أرقام)
  //   - المبلغ (ILS 40.00 أو بمبلغ 40.00 ILS)
  //   - رقم التفويض (236189)
  //   - التاريخ والوقت (2026-08-25 08:44 AM)
  //   - نوع البطاقة (Visa / MasterCard)
  // البيانات كلها بأسماء أرقام بس (بدون اسم زبون) — الاسم بيجي من ملف الأستاذ لاحقاً
  const handleParseSoftPOS = () => {
    const text = softposText.trim();
    if (!text) { setError("الصق نص إيصال SoftPOS أولاً"); return; }
    setError(null);

    // نقسم النص لحركات — كل حركة تبدأ بـ "حركة على بطاقة" أو "حركة"
    const lines = text.split("\n").map(l => l.trim());
    const parseReceiptAmount = (value: string) => {
      const normalized = value.replace(/[٫]/g, ".").replace(/[٬]/g, "").replace(/,/g, ".");
      return Number.parseFloat(normalized) || 0;
    };
    const merchant = lines.find(line => line.startsWith("SoftPOS Receipt") === false && !line.includes(":")) || "";
    const deviceNumber = lines.find(line => line.includes("رقم الجهاز"))?.match(/(\d+)\s*:\s*رقم الجهاز/)?.[1] || "";
    const outletNumber = lines.find(line => line.includes("رقم المخرج"))?.match(/(\d+)\s*:\s*رقم المخرج/)?.[1] || "";
    const summary = lines.find(line => line.includes("عدد حركات"));
    const declaredCount = summary?.match(/عدد حركات\s*=\s*(\d+)/)?.[1] || "";
    const declaredTotal = summary?.match(/مجموع\s*=\s*([\d٫.,]+)\s*ILS/i)?.[1] || "";
    const transactions: Array<{
      visaNumber: string; amount: number; authNumber: string;
      date: string; cardType: string; txId: string; merchant: string;
    }> = [];

    let current: Partial<{
      visaNumber: string; amount: number; authNumber: string;
      date: string; cardType: string; txId: string; merchant: string;
    }> = {};

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];

      // بداية حركة جديدة
      if (line.includes("حركة على بطاقة") || (line.includes("حركة") && line.includes("بمبلغ"))) {
        if (current.visaNumber || current.amount) {
          transactions.push(current as any);
        }
        current = {};
        const merchantMatch = line.match(/من المتجر\s+(.+)$/);
        if (merchantMatch) current.merchant = merchantMatch[1].trim();
        // استخراج رقم البطاقة من السطر: ****-4351
        const cardMatch = line.match(/\*{4}[-\s]*(\d{4})/);
        if (cardMatch) current.visaNumber = cardMatch[1];
        // استخراج المبلغ: بمبلغ 40.00 ILS أو بمبلغ 40٫00 ILS
        const amtMatch = line.match(/بمبلغ\s*([\d٫.,]+)\s*ILS/i) || line.match(/([\d٫.,]+)\s*ILS/i);
        if (amtMatch) current.amount = parseReceiptAmount(amtMatch[1]);
        // نوع البطاقة
        if (line.includes("Visa")) current.cardType = "Visa";
        else if (line.includes("MasterCard")) current.cardType = "MasterCard";
      }

      // رقم التفويض: "236189 رقم التفويض" أو "رقم التفويض 236189"
      if (line.includes("رقم التفويض")) {
        const m = line.match(/(\d{4,})/);
        if (m) current.authNumber = m[1];
      }

      // التاريخ والوقت: "2026-08-25 08:44 AM الوقت والتاريخ" أو العكس
      if (line.includes("الوقت والتاريخ") || line.includes("التاريخ والوقت")) {
        const m = line.match(/(\d{4}-\d{2}-\d{2}\s+\d{1,2}:\d{2}\s*(?:AM|PM)?)/i);
        if (m) current.date = m[1].trim();
      }

      // رقم تعريف وتسلسل الحركة: "13937 - 499538 رقم تعريف وتسلسل الحركة"
      if (line.includes("رقم تعريف") && line.includes("تسلسل")) {
        const m = line.match(/(\d+)\s*[-\s]\s*(\d+)/);
        if (m) current.txId = `${m[1]}-${m[2]}`;
      }

      // نوع ورقم البطاقة: "Visa / ****-4351 نوع ورقم البطاقة"
      if (line.includes("نوع ورقم البطاقة")) {
        const cardMatch = line.match(/\*{4}[-\s]*(\d{4})/);
        if (cardMatch && !current.visaNumber) current.visaNumber = cardMatch[1];
        if (line.includes("Visa")) current.cardType = "Visa";
        else if (line.includes("MasterCard")) current.cardType = "MasterCard";
      }

      // المبلغ في سطر مستقل: "ILS 40.00 المبلغ"
      if (line.includes("المبلغ") && !current.amount) {
        const m = line.match(/([\d٫.,]+)\s*ILS/i) || line.match(/ILS\s*([\d٫.,]+)/i);
        if (m) current.amount = parseReceiptAmount(m[1]);
      }
    }
    // آخر حركة
    if (current.visaNumber || current.amount) {
      transactions.push(current as any);
    }

    if (!transactions.length) {
      setError("لم يتم العثور على أي حركة في النص. تأكد إنك لصقت إيصال SoftPOS كامل.");
      return;
    }

    // إنشاء ملف Excel: سوني كاشير فيزا
    const wb = XLSX.utils.book_new();
    const rows = transactions.map((t, i) => [
      i + 1,
      t.merchant || merchant,
      deviceNumber,
      outletNumber,
      declaredCount,
      parseReceiptAmount(declaredTotal),
      t.visaNumber || "",
      t.authNumber || "",
      t.amount || 0,
      t.date || "",
      t.cardType || "",
      t.txId || "",
    ]);
    const ws = XLSX.utils.aoa_to_sheet([
      ["#", "اسم المتجر", "رقم الجهاز", "رقم المخرج", "عدد الحركات المعلن", "مجموع الحركات المعلن", "رقم الفيزا", "رقم التفويض", "المبلغ", "التاريخ والوقت", "نوع البطاقة", "رقم التعريف والتسلسل"],
      ...rows,
    ]);
    ws["!cols"] = [{wch:5},{wch:28},{wch:14},{wch:14},{wch:18},{wch:20},{wch:12},{wch:15},{wch:12},{wch:22},{wch:14},{wch:20}];
    XLSX.utils.book_append_sheet(wb, ws, "سوني كاشير فيزا");
    XLSX.writeFile(wb, "سوني_كاشير_فيزا.xlsx");
  };

  // ─── Export ─────────────────────────────────────────────────────────────────
  const handleExport = () => {
    const wb = XLSX.utils.book_new();
    const add = (name: string, hdrs: string[], rows: unknown[][]) => {
      const ws = XLSX.utils.aoa_to_sheet([hdrs, ...rows]);
      ws["!cols"] = hdrs.map(() => ({ wch: 25 }));
      XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    };

    add("💳 فيزا بنك (مشتريات تجار)", ["البيان","المبلغ","النوع","التاريخ"],
      visaBankRows.map(b => [b.description, b.rawAmount, b.type, b.date]));

    add("📋 أستاذ - فيزا", ["الاسم","رقم الفيزا","المبلغ","النوع","نوع الحساب","التاريخ"],
      azaVisaRows.map(a => [a.name, a.visaNumber, a.amount, a.type, a.accountType, a.date]));

    add("📋 أستاذ - عادي", ["الاسم","المبلغ","النوع","نوع الحساب","التاريخ"],
      azaRegularRows.map(a => [a.name, a.amount, a.type, a.accountType, a.date]));

    add("📋 أستاذ فقط (زيادة)", ["الاسم","المبلغ","النوع","نوع الحساب","التاريخ"],
      azaOnlyRows.map(a => [a.name, a.amount, a.type, a.accountType, a.date]));

    add("🏪 سوني كاشير فيزا", ["الاسم","رقم الفيزا","رقم التفويض","المبلغ","التاريخ"],
      parsedSony.map(s => [s.name, s.visaNumber, s.authNumber, s.amount, s.date]));

    add("✅ مطابقات مؤكدة", ["اسم سوني","اسم أستاذ","رقم الفيزا","رقم التفويض","المبلغ","تاريخ المطابقة","ملاحظة"],
      matchedVisa.map(m => [m.sonyName, m.azaName, m.visaNumber, m.authNumber, m.amount, m.matchedAt, m.note || ""]));

    add("⏳ مشتريات معلقة (لثاني يوم)", ["البيان","المبلغ","التاريخ","تاريخ النقل"],
      pendingVisaBank.map(p => [p.bankDesc, p.amount, p.date, p.movedAt]));

    add("🔄 مطابقات المعلقة", ["بيان البنك","اسم سوني","المبلغ","سبب المطابقة"],
      pendingMatchResults.map(r => [r.pending.bankDesc, r.matched.sonyName, r.pending.amount, r.reason]));

    add("❌ سوني غير متطابق", ["الاسم","رقم الفيزا","رقم التفويض","المبلغ","التاريخ"],
      unmatchedSony.map(s => [s.name, s.visaNumber, s.authNumber, s.amount, s.date]));

    add("❌ أستاذ فيزا غير متطابق", ["الاسم","رقم الفيزا","المبلغ","النوع","التاريخ"],
      unmatchedAzaVisa.map(a => [a.name, a.visaNumber, a.amount, a.type, a.date]));

    XLSX.writeFile(wb, "تقرير_التسوية_2.xlsx");
  };

  // ─── Stats ──────────────────────────────────────────────────────────────────
  const stats = useMemo(() => ({
    bankTotal: parsedBank.length,
    visaBank: visaBankRows.length,
    azaTotal: parsedAza.length,
    azaVisa: azaVisaRows.length,
    azaRegular: azaRegularRows.length,
    azaOnly: azaOnlyRows.length,
    dailyTotal: parsedDaily.length,
    sonyTotal: parsedSony.length,
    matched: matchedVisa.length,
    pendingBank: pendingVisaBank.length,
    pendingMatched: pendingMatchResults.length,
    unmatchedSony: unmatchedSony.length,
    unmatchedAza: unmatchedAzaVisa.length,
    autoMatch: autoMatchResults.length,
  }), [parsedBank, visaBankRows, parsedAza, azaVisaRows, azaRegularRows, azaOnlyRows,
       parsedDaily, parsedSony, matchedVisa, pendingVisaBank, pendingMatchResults,
       unmatchedSony, unmatchedAzaVisa, autoMatchResults]);

  const canExport = parsedBank.length > 0 || parsedAza.length > 0 || parsedSony.length > 0;

  if (stagePage === "stageA") {
    return <><StageAPlatform initialFile={projectFileToOpen?.stage === "stageA" ? projectFileToOpen : undefined} onBack={() => { setProjectFileToOpen(null); setStagePage("hub"); }} onSafeActivities={onSafeActivities} onTransferToB={(results, summary) => {
      onSafeActivities?.(results.filter(result => result.matched && result.issue !== "تم اعتمادها يدوياً").map(result => `${result.originalName} ↔ ${result.databaseName} · ${fmtNum(result.amount)} ₪`));
      if (onStageAToMain) onStageAToMain(results, summary); else { setStageBInvoices(results); setStagePage("stageB"); }
    }} onTransferVisaToC={results => { setStageCInvoices(results); setStagePage("stageC"); }} /><ProjectSourceFileLauncher files={projectSourceFiles.filter(item => item.stage === "stageA")} onOpen={openProjectFile} /></>;
  }
  if (stagePage === "stageB") {
    return <><StageBPlatform initialFile={projectFileToOpen?.stage === "stageB" ? projectFileToOpen : undefined} invoices={stageBInvoices} onBack={() => { setProjectFileToOpen(null); setStagePage("hub"); }} onSafeActivities={onSafeActivities} onOpenInvoiceReference={onOpenInvoiceReference} onTransferVisaToC={results => openTransferDialog(results, "B")} /><ProjectSourceFileLauncher files={projectSourceFiles.filter(item => item.stage === "stageB")} onOpen={openProjectFile} /><StageTransferDialog dialog={transferDialog} onCancel={() => setTransferDialog(null)} onTransferToB={invoices => { setStageBInvoices(previous => [...previous, ...invoices]); setStagePage("stageB"); setTransferDialog(null); }} onTransferToC={invoices => { setStageCInvoices(previous => [...previous, ...invoices]); setStagePage("stageC"); setTransferDialog(null); }} /></>;
  }
  if (stagePage === "stageC") {
    const stageDBankInputs = [
      ...initialStageDBankInputs,
      ...pendingVisaBank.map(item => ({ id: item.id, description: item.bankDesc, amount: item.amount, date: item.date }))
    ];
    return <><StageCPlatform initialFile={projectFileToOpen?.stage === "stageC" ? projectFileToOpen : undefined} invoices={stageCInvoices} pendingBankInputs={stageDBankInputs} onBack={() => { setProjectFileToOpen(null); setStagePage("hub"); }} onTransferToB={invoices => openTransferDialog(invoices, "C")} onSafeActivities={onSafeActivities} onTransferToD={(matches, pending) => { setStageDMatches(matches); setStageDPending(pending); setStagePage("stageD"); }} /><ProjectSourceFileLauncher files={projectSourceFiles.filter(item => item.stage === "stageC")} onOpen={openProjectFile} /><StageTransferDialog dialog={transferDialog} onCancel={() => setTransferDialog(null)} onTransferToB={invoices => { setStageBInvoices(previous => [...previous, ...invoices]); setStagePage("stageB"); setTransferDialog(null); }} onTransferToC={invoices => { setStageCInvoices(previous => [...previous, ...invoices]); setStagePage("stageC"); setTransferDialog(null); }} /></>;
  }
  if (stagePage === "stageD") {
    return <StageDPlatform matches={stageDMatches} pending={stageDPending} onBack={() => setStagePage("stageC")} />;
  }
  if (stagePage === "stageF") {
    return <div dir="rtl" className="min-h-screen bg-slate-50 p-6 text-slate-900"><div className="mx-auto max-w-4xl space-y-5">
      <div className="flex flex-wrap items-center justify-between gap-3"><button onClick={() => setStagePage("hub")} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-900"><ChevronLeft className="h-4 w-4" />العودة لمنصة الفيزا</button><h1 className="text-xl font-black">F — تحويل إيصال SoftPOS إلى Excel</h1></div>
      {error && <div role="alert" className="rounded-lg bg-red-100 p-3 text-sm text-red-800">{error}</div>}
      <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm"><div className="mb-3 flex items-center gap-2"><FileSpreadsheet className="h-5 w-5 text-slate-700" /><div><h2 className="text-sm font-bold">لصق إيصال SoftPOS</h2><p className="text-xs text-slate-500">الصق نص الإيصال لتحويل كل حركة إلى صف مستقل في ملف سوني كاشير فيزا</p></div></div><textarea value={softposText} onChange={e => setSoftposText(e.target.value)} placeholder="الصق نص إيصال SoftPOS هنا..." className="h-72 w-full resize-y rounded-lg border border-slate-200 bg-slate-50 p-3 text-xs font-mono outline-none focus:ring-2 focus:ring-slate-200" />
      <div className="mt-3 flex flex-wrap gap-2"><button onClick={handleParseSoftPOS} disabled={!softposText.trim()} className="flex items-center gap-2 rounded-lg bg-slate-700 px-4 py-2 text-sm font-medium text-white hover:bg-slate-800 disabled:opacity-40"><Download className="h-4 w-4" />إنشاء ملف Excel</button>{softposText.trim() && <button onClick={() => setSoftposText("")} className="flex items-center gap-1 rounded-lg bg-slate-100 px-3 py-2 text-xs text-slate-600 hover:bg-slate-200"><X className="h-3.5 w-3.5" />مسح النص</button>}</div></div>
      <div className="rounded-xl border border-blue-200 bg-blue-50 p-4 text-xs leading-6 text-blue-900">يتم استخراج رقم البطاقة، المبلغ، رقم التفويض، الوقت والتاريخ، نوع البطاقة، ورقم تعريف وتسلسل الحركة، مع إضافة بيانات المتجر والجهاز والمخرج إلى ملف Excel.</div>
    </div></div>;
  }

  return (
    <div dir="rtl" className="min-h-screen bg-slate-50 text-slate-900 p-6">
      <div className="max-w-6xl mx-auto space-y-6">

        {/* Header */}
        <div className="relative overflow-hidden rounded-2xl border border-slate-200 bg-white px-5 py-6 text-center shadow-sm">
          <div className="mx-auto mb-4 flex h-16 w-16 items-center justify-center overflow-hidden rounded-2xl bg-white ring-4 ring-amber-100 shadow-lg">
            <img src="/taswiyah-logo.png" alt="Taswiyah Hub" className="h-full w-full object-contain p-1" />
          </div>
           <div className="mb-1 text-xs font-bold tracking-[.18em] text-amber-600">TASWIYAH HUB</div>
           <h1 className="text-3xl font-black tracking-tight text-slate-900">محطتك الأولى لتسوية الحسابات بدقة.</h1>
           <p className="mx-auto mt-2 max-w-2xl text-sm text-slate-500">اختر المرحلة المناسبة لمعالجة الفواتير والحوالات والفيزا</p>
          <div className="absolute left-5 top-5 flex items-center gap-2">
          {canExport && (
            <button onClick={handleExport} className="absolute left-5 top-5 flex items-center gap-2 rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white transition-colors hover:bg-green-700">
              <Download className="w-4 h-4"/>تصدير Excel
            </button>
          )}
          <div className="relative">
            <button onClick={e => e.currentTarget.nextElementSibling?.classList.toggle("hidden")} className="flex items-center gap-2 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-xs font-medium text-red-700 hover:bg-red-100">
              <Trash2 className="h-3.5 w-3.5" />المحفوظات
            </button>
            <div className="absolute left-0 top-10 z-20 hidden w-56 rounded-xl border border-slate-200 bg-white p-1 text-right shadow-lg">
              <button onClick={handleClearSaved} className="flex w-full items-center gap-2 rounded-lg px-3 py-2 text-xs text-red-700 hover:bg-red-50"><Trash2 className="h-3.5 w-3.5" />مسح المطابقات والمعلقات</button>
              <button onClick={handleClearSession} className="flex w-full items-center gap-2 rounded-lg border-t border-slate-100 px-3 py-2 text-xs text-amber-700 hover:bg-amber-50"><RotateCcw className="h-3.5 w-3.5" />مسح الجلسة كاملة</button>
            </div>
          </div>
          </div>
        </div>

        {/* Visa workflow hub */}
        <section className="grid grid-cols-1 gap-4 md:grid-cols-2">
          {[
            { id: "stage-a", code: "A", title: "فرز الفواتير حسب قاعدة البيانات", description: "مطابقة اسم ومبلغ كل فاتورة مع قاعدة البيانات، ثم استخراج المطابقات ومشاكل أخرى.", color: "bg-blue-600", text: "text-blue-600" },
            { id: "stage-b", code: "B", title: "المطابقة الذكية للفواتير والحوالات", description: "فتح أداة مطابقة الفواتير والحوالات.", color: "bg-emerald-600", text: "text-emerald-600" },
            { id: "stage-c", code: "C", title: "مطابقة السوني كاشير (فيزا)", description: "مطابقة العمليات المعلقة واسترجاع أرقام التفويض مع بيانات الفيزا.", color: "bg-amber-600", text: "text-amber-600" },
            { id: "stage-d", code: "D", title: "مطابقة الفيزا — المرحلة الثانية", description: "مطابقة أرقام الفيزا المؤكدة مع الحوالات البنكية المعلقة.", color: "bg-indigo-600", text: "text-indigo-600" },
            { id: "stage-f", code: "F", title: "تحويل إيصال SoftPOS إلى Excel", description: "لصق نص الإيصال وتحويل الحركات وأرقام البطاقات والتفويض إلى ملف سوني جاهز.", color: "bg-slate-700", text: "text-slate-700" },
          ].map(stage => (
            <button key={stage.id} onClick={() => stage.id === "stage-a" ? setStagePage("stageA") : stage.id === "stage-b" ? (onStageBToMain ? onStageBToMain() : onBack()) : stage.id === "stage-c" ? setStagePage("stageC") : stage.id === "stage-d" ? setStagePage("stageD") : setStagePage("stageF")}
              className="group flex min-h-36 items-center gap-5 rounded-2xl border border-slate-200 bg-white p-5 text-right shadow-sm transition-all hover:-translate-y-0.5 hover:border-slate-300 hover:shadow-md">
              <span className={`flex h-14 w-14 shrink-0 items-center justify-center rounded-xl ${stage.color} text-xl font-black text-white shadow-sm`}>{stage.code}</span>
              <span className="min-w-0 flex-1">
                <span className={`mb-1 block text-base font-bold ${stage.text}`}>{stage.title}</span>
                <span className="block text-xs leading-6 text-slate-500">{stage.description}</span>
                <span className={`mt-2 block text-xs font-semibold ${stage.text}`}>فتح المرحلة ←</span>
              </span>
            </button>
          ))}
        </section>

        {/* Results */}
        {(parsedBank.length > 0 || parsedAza.length > 0 || parsedSony.length > 0) && (
          <div className="space-y-4">
            {/* Stats */}
            <div className="grid grid-cols-3 sm:grid-cols-7 gap-2">
              {([
                { id: "overview", label: "نظرة عامة", count: "", color: "text-blue-700" },
                { id: "bank", label: "فيزا بنك", count: stats.visaBank, color: "text-teal-700" },
                { id: "aza", label: "أستاذ", count: stats.azaTotal, color: "text-indigo-700" },
                { id: "sony", label: "🏪 سوني", count: stats.sonyTotal, color: "text-purple-700" },
                { id: "matched", label: "مؤكد", count: stats.matched, color: "text-green-700" },
                { id: "pendingBank", label: "معلقة بنك", count: stats.pendingBank, color: "text-amber-700" },
                { id: "unmatched", label: "غير متطابق", count: stats.unmatchedSony + stats.unmatchedAza, color: "text-red-700" },
              ] as const).map(t => (
                <button key={t.id} onClick={() => setTab(t.id as R2Tab)}
                  className={`p-2.5 rounded-xl border text-right transition-all ${tab === t.id ? "border-blue-500 bg-blue-50" : "bg-card hover:bg-muted/30"}`}>
                  <span className="text-[10px] text-muted-foreground block">{t.label}</span>
                  <span className={`text-base font-mono font-bold block mt-0.5 ${t.color}`}>{t.count}</span>
                </button>
              ))}
            </div>

            <div className="border rounded-xl bg-card overflow-hidden">

              {/* Overview tab */}
              {tab === "overview" && (
                <div className="p-6 space-y-4">
                  <h2 className="font-bold text-sm">نظرة عامة على الآلية الجديدة</h2>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <div className="p-4 rounded-xl bg-teal-50 border border-teal-200">
                      <div className="text-xs text-teal-600 font-medium mb-1">كشف البنك</div>
                      <div className="text-2xl font-bold text-teal-800">{stats.bankTotal}</div>
                      <div className="text-xs text-teal-700 mt-1">فيزا بنك (مشتريات تجار): {stats.visaBank}</div>
                    </div>
                    <div className="p-4 rounded-xl bg-indigo-50 border border-indigo-200">
                      <div className="text-xs text-indigo-600 font-medium mb-1">كشف الأستاذ</div>
                      <div className="text-2xl font-bold text-indigo-800">{stats.azaTotal}</div>
                      <div className="text-xs text-indigo-700 mt-1">فيزا: {stats.azaVisa} | عادي: {stats.azaRegular} | زيادة: {stats.azaOnly}</div>
                    </div>
                    <div className="p-4 rounded-xl bg-purple-50 border border-purple-200">
                      <div className="text-xs text-purple-600 font-medium mb-1">🏪 سوني كاشير فيزا</div>
                      <div className="text-2xl font-bold text-purple-800">{stats.sonyTotal}</div>
                      <div className="text-xs text-purple-700 mt-1">مطابق: {stats.matched} | غير متطابق: {stats.unmatchedSony}</div>
                    </div>
                    <div className="p-4 rounded-xl bg-amber-50 border border-amber-200">
                      <div className="text-xs text-amber-600 font-medium mb-1">مشتريات معلقة (لثاني يوم)</div>
                      <div className="text-2xl font-bold text-amber-800">{stats.pendingBank}</div>
                      <div className="text-xs text-amber-700 mt-1">مطابق مع سوني: {stats.pendingMatched}</div>
                    </div>
                  </div>

                  {autoMatchResults.length > 0 && (
                    <div className="p-4 rounded-xl bg-green-50 border border-green-200">
                      <div className="flex items-center justify-between flex-wrap gap-2">
                        <div>
                          <div className="text-sm font-bold text-green-800">🤖 مطابقات تلقائية مقترحة</div>
                          <div className="text-xs text-green-700 mt-0.5">{autoMatchResults.length} مطابقة (سوني ↔ أستاذ فيزا)</div>
                        </div>
                        <button onClick={handleAcceptAllAuto}
                          className="flex items-center gap-1.5 px-4 py-2 bg-green-600 text-white rounded-lg text-sm font-medium hover:bg-green-700 transition-colors">
                          <Check className="w-4 h-4"/>تأكيد الكل
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* Visa Bank tab */}
              {tab === "bank" && (
                <>
                <div className="flex flex-wrap items-center justify-between gap-2 border-b bg-amber-50/70 px-4 py-3">
                  <p className="text-xs text-amber-900">مشتريات عمولة تجار تُستبعد من المطابقة العادية ويمكن ترحيلها للمعلّقات حتى مطابقتها لاحقاً مع تفويض سوني.</p>
                  <button type="button" onClick={handleMoveVisaBankToPending} disabled={!visaBankRows.some(bank => !pendingVisaBank.some(item => item.bankId === bank.id))}
                    className="rounded-lg bg-amber-600 px-3 py-2 text-xs font-medium text-white disabled:opacity-40">
                    نقل مشتريات التجار للمعلّقات ({visaBankRows.filter(bank => !pendingVisaBank.some(item => item.bankId === bank.id)).length})
                  </button>
                </div>
                <table className="w-full text-sm">
                  <thead><tr className="bg-muted text-xs">
                    <th className="px-3 py-2.5 text-right">البيان</th>
                    <th className="px-3 py-2.5 text-right">المبلغ</th>
                    <th className="px-3 py-2.5 text-right">النوع</th>
                    <th className="px-3 py-2.5 text-right">التاريخ</th>
                  </tr></thead>
                  <tbody>
                    {visaBankRows.map(b => (
                      <tr key={b.id} className="border-t hover:bg-muted/20 transition-colors">
                        <td className="px-3 py-2.5 font-medium text-teal-700">{b.description}</td>
                        <td className="px-3 py-2.5 font-mono font-bold text-teal-700">{fmtNum(b.rawAmount)}</td>
                        <td className={`px-3 py-2.5 text-xs font-semibold ${b.type === "مدفوع" ? "text-red-600" : "text-green-600"}`}>{b.type}</td>
                        <td className="px-3 py-2.5 text-xs text-muted-foreground">{b.date}</td>
                      </tr>
                    ))}
                    {!visaBankRows.length && <tr><td colSpan={4} className="py-10 text-center text-muted-foreground text-xs">لا توجد حوالات "مشتريات عمولة تجار". ارفع كشف البنك أولاً.</td></tr>}
                  </tbody>
                </table>
                </>
              )}

              {/* Aza tab */}
              {tab === "aza" && (
                <div>
                  <div className="flex items-center gap-2 px-4 py-2.5 bg-indigo-50/60 border-b border-indigo-100 flex-wrap">
                    <span className="text-xs font-medium text-indigo-800">أستاذ عادي: {azaRegularRows.length}</span>
                    <span className="text-xs text-indigo-600">| أستاذ فيزا: {azaVisaRows.length}</span>
                    {azaOnlyRows.length > 0 && (
                      <span className="text-xs px-2 py-0.5 rounded bg-amber-100 text-amber-700 border border-amber-300">{azaOnlyRows.length} زيادة بالأستاذ (مش باليومي)</span>
                    )}
                  </div>
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 text-right">الاسم</th>
                      <th className="px-3 py-2.5 text-right">رقم الفيزا</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                      <th className="px-3 py-2.5 text-right">النوع</th>
                      <th className="px-3 py-2.5 text-right">نوع الحساب</th>
                      <th className="px-3 py-2.5 text-right">التاريخ</th>
                    </tr></thead>
                    <tbody>
                      {parsedAza.map(a => (
                        <tr key={a.id} className={`border-t transition-colors ${a.visaNumber ? "bg-teal-50/40 hover:bg-teal-50" : "hover:bg-muted/20"}`}>
                          <td className="px-3 py-2.5 font-medium">
                            {a.name}
                            {azaOnlyRows.some(o => o.id === a.id) && (
                              <span className="mr-1.5 text-[9px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 border border-amber-300">زيادة</span>
                            )}
                          </td>
                          <td className="px-3 py-2.5">
                            {a.visaNumber ? <span className="font-mono font-bold text-teal-700 bg-teal-50 px-2 py-0.5 rounded border border-teal-200">{a.visaNumber}</span> : "—"}
                          </td>
                          <td className="px-3 py-2.5 font-mono font-bold">{fmtNum(a.amount)}</td>
                          <td className={`px-3 py-2.5 text-xs font-semibold ${a.type === "مدفوع" ? "text-red-600" : "text-green-600"}`}>{a.type}</td>
                          <td className="px-3 py-2.5 text-xs text-muted-foreground">{a.accountType || "—"}</td>
                          <td className="px-3 py-2.5 text-xs text-muted-foreground">{a.date}</td>
                        </tr>
                      ))}
                      {!parsedAza.length && <tr><td colSpan={6} className="py-10 text-center text-muted-foreground text-xs">لا توجد بيانات. ارفع كشف الأستاذ أولاً.</td></tr>}
                    </tbody>
                  </table>
                </div>
              )}

              {/* Sony tab */}
              {tab === "sony" && (
                <table className="w-full text-sm">
                  <thead><tr className="bg-muted text-xs">
                    <th className="px-3 py-2.5 text-right">الاسم</th>
                    <th className="px-3 py-2.5 text-right">رقم الفيزا</th>
                    <th className="px-3 py-2.5 text-right">رقم التفويض</th>
                    <th className="px-3 py-2.5 text-right">المبلغ</th>
                    <th className="px-3 py-2.5 text-right">التاريخ</th>
                    <th className="px-3 py-2.5 text-right">الحالة</th>
                  </tr></thead>
                  <tbody>
                    {parsedSony.map(s => {
                      const isMatched = matchedSonyIds.has(s.id);
                      return (
                        <tr key={s.id} className={`border-t transition-colors ${isMatched ? "bg-green-50/50" : "hover:bg-muted/20"}`}>
                          <td className="px-3 py-2.5 font-medium">{s.name}</td>
                          <td className="px-3 py-2.5">
                            {s.visaNumber ? <span className="font-mono font-bold text-teal-700 bg-teal-50 px-2 py-0.5 rounded border border-teal-200">{s.visaNumber}</span> : "—"}
                          </td>
                          <td className="px-3 py-2.5 text-xs font-mono">{s.authNumber || "—"}</td>
                          <td className="px-3 py-2.5 font-mono font-bold">{fmtNum(s.amount)}</td>
                          <td className="px-3 py-2.5 text-xs text-muted-foreground">{s.date}</td>
                          <td className="px-3 py-2.5">
                            {isMatched ? (
                              <span className="text-green-600 bg-green-50 px-1.5 py-0.5 rounded text-xs border border-green-200 flex items-center gap-0.5 w-fit">
                                <Check className="w-2.5 h-2.5"/>مطابق
                              </span>
                            ) : (
                              <span className="text-amber-600 bg-amber-50 px-1.5 py-0.5 rounded text-xs border border-amber-200">بانتظار</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                    {!parsedSony.length && <tr><td colSpan={6} className="py-10 text-center text-muted-foreground text-xs">لا توجد بيانات. ارفع ملف سوني كاشير فيزا.</td></tr>}
                  </tbody>
                </table>
              )}

              {/* Matched tab — سوني ↔ أستاذ فيزا */}
              {tab === "matched" && (
                <div>
                  {autoMatchResults.length > 0 && (
                    <div className="flex items-center gap-3 px-4 py-2.5 bg-green-50/60 border-b border-green-100 flex-wrap">
                      <span className="text-xs text-green-800">{autoMatchResults.length} مطابقة تلقائية مقترحة</span>
                      <button onClick={handleAcceptAllAuto}
                        className="mr-auto flex items-center gap-1.5 px-3 py-1.5 bg-green-600 text-white rounded-lg text-xs font-medium hover:bg-green-700 transition-colors">
                        <Check className="w-3.5 h-3.5"/>تأكيد الكل
                      </button>
                    </div>
                  )}
                  {/* Auto-match suggestions */}
                  {autoMatchResults.map((r, i) => {
                    const key = `auto-${r.sony.id}-${r.aza.id}`;
                    const isExpanded = expandedKey === key;
                    return (
                      <React.Fragment key={i}>
                        <div className={`px-4 py-3 border-b transition-colors ${isExpanded ? "bg-blue-50" : "bg-blue-50/30 hover:bg-blue-50/60"}`}>
                          <div className="flex items-center gap-3 flex-wrap">
                            <div className="flex-1 min-w-0">
                              <div className="flex items-center gap-2 flex-wrap">
                                <span className="font-medium text-sm">{r.sony.name}</span>
                                {r.sony.visaNumber && <span className="text-[10px] px-1.5 py-0.5 rounded bg-teal-100 text-teal-700 border border-teal-300 font-mono">{r.sony.visaNumber}</span>}
                                {r.visaMatch && <span className="text-[10px] px-1.5 py-0.5 rounded bg-green-100 text-green-700 border border-green-300">✓ فيزا مطابق</span>}
                                <span className="text-muted-foreground text-xs">↔</span>
                                <span className="text-sm text-muted-foreground">{r.aza.name}</span>
                                <span className="font-mono font-bold text-blue-700 text-sm">{fmtNum(r.aza.amount)}</span>
                                {r.amountDiff > 0.01 && <span className="text-[10px] text-orange-600">فرق: {fmtNum(r.amountDiff)}</span>}
                                <span className={`text-[10px] ${r.score * 100 >= 90 ? "text-green-600" : "text-amber-600"}`}>({(r.score * 100).toFixed(0)}%)</span>
                              </div>
                            </div>
                            <div className="flex items-center gap-1.5 shrink-0">
                              <button onClick={() => handleAcceptMatch(r.sony, r.aza, r.amountDiff > 0.01 ? `فرق مبلغ: ${fmtNum(r.amountDiff)}` : undefined)}
                                className="flex items-center gap-1 px-3 py-1.5 bg-green-600 text-white rounded-lg text-xs hover:bg-green-700 transition-colors">
                                <Check className="w-3.5 h-3.5"/>تأكيد
                              </button>
                              <button onClick={() => handleRejectMatch(r.sony.id, r.aza.id)}
                                className="flex items-center gap-1 px-3 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs hover:bg-red-100 transition-colors">
                                <X className="w-3.5 h-3.5"/>رفض
                              </button>
                              <button onClick={() => setExpandedKey(isExpanded ? null : key)}
                                className="p-1 text-muted-foreground hover:text-foreground">
                                <span className="text-base font-bold">⋮</span>
                              </button>
                            </div>
                          </div>
                          {isExpanded && (
                            <div className="mt-3 pt-3 border-t border-blue-200 text-xs space-y-1">
                              <div><strong>سوني:</strong> {r.sony.rawName} | مبلغ: {fmtNum(r.sony.amount)} | فيزا: {r.sony.visaNumber || "—"} | تفويض: {r.sony.authNumber || "—"}</div>
                              <div><strong>أستاذ:</strong> {r.aza.rawName} | مبلغ: {fmtNum(r.aza.amount)} | فيزا: {r.aza.visaNumber || "—"} | حساب: {r.aza.accountType || "—"}</div>
                              <div><strong>نوع المطابقة:</strong> {r.matchType} | <strong>التشابه:</strong> {(r.score * 100).toFixed(0)}%</div>
                            </div>
                          )}
                        </div>
                      </React.Fragment>
                    );
                  })}

                  {/* Confirmed matches */}
                  {matchedVisa.length > 0 && (
                    <div className="px-4 py-2 bg-green-50/40 border-b">
                      <span className="text-xs font-bold text-green-800">مطابقات مؤكدة ({matchedVisa.length})</span>
                    </div>
                  )}
                  {matchedVisa.map(m => (
                    <div key={m.id} className="px-4 py-3 border-b hover:bg-green-50/30 transition-colors">
                      <div className="flex items-center gap-3 flex-wrap">
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2 flex-wrap">
                            <Shield className="w-3.5 h-3.5 text-green-600"/>
                            <span className="font-medium text-sm">{m.sonyName}</span>
                            {m.visaNumber && <span className="text-[10px] px-1.5 py-0.5 rounded bg-teal-100 text-teal-700 border border-teal-300 font-mono">{m.visaNumber}</span>}
                            {m.authNumber && <span className="text-[10px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 border border-purple-300 font-mono">تفويض: {m.authNumber}</span>}
                            <span className="text-muted-foreground text-xs">↔</span>
                            <span className="text-sm text-muted-foreground">{m.azaName}</span>
                            <span className="font-mono font-bold text-green-700 text-sm">{fmtNum(m.amount)}</span>
                            {m.note && <span className="text-[10px] text-orange-600">{m.note}</span>}
                          </div>
                        </div>
                        <button onClick={() => handleUnmatch(m)}
                          className="flex items-center gap-1 px-2.5 py-1.5 bg-amber-50 text-amber-600 border border-amber-200 rounded-lg text-xs hover:bg-amber-100 transition-colors">
                          <X className="w-3.5 h-3.5"/>إلغاء
                        </button>
                      </div>
                    </div>
                  ))}

                  {!autoMatchResults.length && !matchedVisa.length && (
                    <div className="py-10 text-center text-muted-foreground text-xs">
                      لا توجد مطابقات بعد. ارفع ملف سوني كاشير فيزا وكشف الأستاذ لتظهر المطابقات التلقائية.
                    </div>
                  )}
                </div>
              )}

              {/* Pending Bank tab — مشتريات معلقة لثاني يوم */}
              {tab === "pendingBank" && (
                <div>
                  <div className="px-4 py-3 bg-amber-50/60 border-b border-amber-100">
                    <p className="text-xs text-amber-800">
                      هذه الحركات «مشتريات عمولة تجار» محفوظة للمراجعة لاحقاً. رقم التفويض هو الجزء الأوسط من البيان البنكي، ويُطابق مع رقم التفويض في صف سوني.
                    </p>
                    {!!pendingVisaBank.length && <button type="button" onClick={() => setStagePage("stageC")}
                      className="mt-2 rounded-lg bg-indigo-600 px-3 py-2 text-xs font-medium text-white hover:bg-indigo-700">
                      مطابقة هذه المعلّقات عبر C ثم D ({pendingVisaBank.length})
                    </button>}
                    {pendingMatchResults.length > 0 && (
                      <p className="text-xs text-green-700 mt-1">{pendingMatchResults.length} مطابقة محتملة مع سوني مؤكد!</p>
                    )}
                  </div>
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 text-right">بيان البنك</th>
                      <th className="px-3 py-2.5 text-right">رقم التفويض</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                      <th className="px-3 py-2.5 text-right">التاريخ</th>
                      <th className="px-3 py-2.5 text-right">مطابقة محتملة</th>
                      <th className="px-3 py-2.5 text-right">إجراء</th>
                    </tr></thead>
                    <tbody>
                      {pendingVisaBank.map(p => {
                        const match = pendingMatchResults.find(r => r.pending.id === p.id);
                        return (
                          <tr key={p.id} className="border-t hover:bg-muted/20 transition-colors">
                            <td className="px-3 py-2.5 font-medium text-teal-700">{p.bankDesc}</td>
                            <td className="px-3 py-2.5 font-mono">{p.authNumber || "—"}</td>
                            <td className="px-3 py-2.5 font-mono font-bold text-amber-700">{fmtNum(p.amount)}</td>
                            <td className="px-3 py-2.5 text-xs text-muted-foreground">{p.date}</td>
                            <td className="px-3 py-2.5 text-xs">
                              {match ? (
                                <span className="text-green-700 bg-green-50 px-1.5 py-0.5 rounded border border-green-200">
                                  {match.matched.sonyName} — {match.reason}
                                </span>
                              ) : (
                                <span className="text-muted-foreground">لا توجد</span>
                              )}
                            </td>
                            <td className="px-3 py-2.5">
                              <button onClick={() => setPendingVisaBank(prev => prev.filter(x => x.id !== p.id))}
                                className="flex items-center gap-1 px-2.5 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs hover:bg-red-100 transition-colors">
                                <Trash2 className="w-3.5 h-3.5"/>حذف
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                      {!pendingVisaBank.length && <tr><td colSpan={6} className="py-10 text-center text-muted-foreground text-xs">لا توجد حوالات معلقة. فرز كشف البنك وانقل «مشتريات عمولة تجار» للمعلّقات.</td></tr>}
                    </tbody>
                  </table>
                </div>
              )}

              {/* Unmatched tab */}
              {tab === "unmatched" && (
                <div className="p-4 space-y-4">
                  <div>
                    <h3 className="font-semibold text-sm mb-2">سوني كاشير فيزا غير متطابق ({unmatchedSony.length})</h3>
                    <table className="w-full text-sm border rounded-lg overflow-hidden">
                      <thead><tr className="bg-muted text-xs">
                        <th className="px-3 py-2.5 text-right">الاسم</th>
                        <th className="px-3 py-2.5 text-right">رقم الفيزا</th>
                        <th className="px-3 py-2.5 text-right">رقم التفويض</th>
                        <th className="px-3 py-2.5 text-right">المبلغ</th>
                      </tr></thead>
                      <tbody>
                        {unmatchedSony.map(s => (
                          <tr key={s.id} className="border-t hover:bg-muted/20">
                            <td className="px-3 py-2.5 font-medium">{s.name}</td>
                            <td className="px-3 py-2.5">{s.visaNumber ? <span className="font-mono text-teal-700">{s.visaNumber}</span> : "—"}</td>
                            <td className="px-3 py-2.5 text-xs font-mono">{s.authNumber || "—"}</td>
                            <td className="px-3 py-2.5 font-mono font-bold">{fmtNum(s.amount)}</td>
                          </tr>
                        ))}
                        {!unmatchedSony.length && <tr><td colSpan={4} className="py-6 text-center text-muted-foreground text-xs">كل سوني مطابق ✓</td></tr>}
                      </tbody>
                    </table>
                  </div>
                  <div>
                    <h3 className="font-semibold text-sm mb-2">أستاذ فيزا غير متطابق ({unmatchedAzaVisa.length})</h3>
                    <table className="w-full text-sm border rounded-lg overflow-hidden">
                      <thead><tr className="bg-muted text-xs">
                        <th className="px-3 py-2.5 text-right">الاسم</th>
                        <th className="px-3 py-2.5 text-right">رقم الفيزا</th>
                        <th className="px-3 py-2.5 text-right">المبلغ</th>
                        <th className="px-3 py-2.5 text-right">نوع الحساب</th>
                      </tr></thead>
                      <tbody>
                        {unmatchedAzaVisa.map(a => (
                          <tr key={a.id} className="border-t hover:bg-muted/20">
                            <td className="px-3 py-2.5 font-medium">{a.name}</td>
                            <td className="px-3 py-2.5">{a.visaNumber ? <span className="font-mono text-teal-700">{a.visaNumber}</span> : "—"}</td>
                            <td className="px-3 py-2.5 font-mono font-bold">{fmtNum(a.amount)}</td>
                            <td className="px-3 py-2.5 text-xs text-muted-foreground">{a.accountType || "—"}</td>
                          </tr>
                        ))}
                        {!unmatchedAzaVisa.length && <tr><td colSpan={4} className="py-6 text-center text-muted-foreground text-xs">كل أستاذ فيزا مطابق ✓</td></tr>}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

            </div>
          </div>
        )}
      </div>
    </div>
  );
}
