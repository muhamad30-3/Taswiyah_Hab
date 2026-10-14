"use client"
/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unused-vars */
import React, { Suspense, lazy, useState, useCallback, useMemo, useEffect, useRef, useDeferredValue } from "react";
import * as XLSX from "xlsx";
import type { StageAResult, StageASummary, StageDBankInput } from "./Reconciliation2";
import { BatchCutoffTool, ClearingTool, SplitTool, syntheticTransactionId, transactionKey, type BankSplit, type ClearingGroup, type CutoffBatch, type CutoffRowRef, type SplitAllocation, type ToolBankRow, type ToolCashierRow } from "./reconciliation-tools";
import { Upload, FileSpreadsheet, Download, ChevronDown, X, Trash2, TriangleAlert as AlertTriangle, Check, Search, Link2, Link2Off, ChevronLeft, ChevronRight, CreditCard, Sparkles, Info, Save, Shield, FolderOpen, FolderPlus, RotateCcw, FolderMinus, GripVertical, Landmark, Users, BarChart3, ArrowUpRight, Plus, Star, MessageSquare } from "lucide-react";
import { cleanStatementText, removeStatementBalances, type StatementCell } from "../lib/account-statement-processor";
import { centralStorageDelete, centralStorageEnabled, centralStorageGet, centralStorageSet, isCentralAdmin, uploadStoredFile } from "../lib/central-storage";
import { clearStageASourceFiles, getStageASourceFiles } from "../lib/stage-a-source-files";

const Reconciliation2 = lazy(() => import("./Reconciliation2"));

// ─── Excel helpers ────────────────────────────────────────────────────────────
function readFileBuf(f: File): Promise<ArrayBuffer> {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(e.target!.result as ArrayBuffer);
    r.onerror = () => rej(new Error("فشل قراءة الملف"));
    r.readAsArrayBuffer(f);
  });
}
function cleanStatementBalanceData(headers: string[], rows: Record<string, unknown>[]) {
  return removeStatementBalances(headers, rows);
}
function normalizeColumnName(value: string): string {
  return value.toLowerCase().normalize("NFKC").replace(/[\u064b-\u065f\u0670]/g, "").replace(/[_-]+/g, " ").replace(/[^\p{L}\p{N}]+/gu, " ").replace(/\s+/g, " ").trim();
}
function cleanJawwalPayLedger(headers: string[], rows: Record<string, unknown>[]) {
  const entryHeader = headers.find(header => /رقم\s*القيد/.test(normalizeColumnName(header)));
  const debitHeader = headers.find(header => /^(?:ال)?مدين$|^debit$/.test(normalizeColumnName(header)));
  const creditHeader = headers.find(header => /^(?:ال)?دائن$|^credit$/.test(normalizeColumnName(header)));
  const isLedger = !!entryHeader && !!debitHeader && !!creditHeader &&
    headers.some(header => /رصيد|balance/i.test(header));
  const outputHeaders = isLedger ? headers.filter(header => !/رصيد|balance/i.test(header)) : headers;
  if (!isLedger) return { headers: outputHeaders, rows };
  const outputRows = rows.filter(row => {
    return isJawwalPayLedgerMovement(headers, row);
  }).map(row => Object.fromEntries(outputHeaders.map(header => [header, cleanStatementText((row[header] ?? "") as StatementCell)])));
  return { headers: outputHeaders, rows: outputRows };
}
function isJawwalPayLedgerMovement(headers: string[], row: Record<string, unknown>) {
  const entryHeader = headers.find(header => /رقم\s*القيد/.test(normalizeColumnName(header)));
  const debitHeader = headers.find(header => /^(?:ال)?مدين$|^debit$/.test(normalizeColumnName(header)));
  const creditHeader = headers.find(header => /^(?:ال)?دائن$|^credit$/.test(normalizeColumnName(header)));
  const isLedger = !!entryHeader && !!debitHeader && !!creditHeader &&
    headers.some(header => /رصيد|balance/i.test(header));
  if (!isLedger) return true;
  const entryNumber = String(row[entryHeader] ?? "").trim().replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)));
  const debit = Math.abs(toNum(row[debitHeader]));
  const credit = Math.abs(toNum(row[creditHeader]));
  return /^\d+$/.test(entryNumber) && (debit > 0 || credit > 0);
}
function downloadCleanedSheet(fileName: string, headers: string[], rows: Record<string, unknown>[]) {
  const cleaned = cleanStatementBalanceData(headers, rows);
  if (!cleaned.headers.length || !cleaned.rows.length) return;
  const worksheet = XLSX.utils.aoa_to_sheet([cleaned.headers, ...cleaned.rows.map((row) => cleaned.headers.map((header) => row[header] ?? ""))]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "الكشف المنظف");
  const baseName = fileName.replace(/\.[^.]+$/, "") || "كشف";
  XLSX.writeFile(workbook, `${baseName}-منظف.xlsx`);
}
function fmtDate(v: unknown): string {
  if (!v) return "";
  if (v instanceof Date) return `${String(v.getDate()).padStart(2, "0")}/${String(v.getMonth() + 1).padStart(2, "0")}/${v.getFullYear()}`;
  if (typeof v === "number" && Number.isFinite(v)) {
    const date = new Date(Date.UTC(1899, 11, 30) + v * 86_400_000);
    return `${String(date.getUTCDate()).padStart(2, "0")}/${String(date.getUTCMonth() + 1).padStart(2, "0")}/${date.getUTCFullYear()}`;
  }
  const normalized = String(v).replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, digit => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)))
    .replace(/[\u200e\u200f\u202a-\u202e]/g, "").trim();
  const dateParts = normalized.match(/^(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})/);
  if (dateParts) {
    const [, first, second, third] = dateParts;
    if (first.length === 4) return `${third.padStart(2, "0")}/${second.padStart(2, "0")}/${first}`;
    return `${first.padStart(2, "0")}/${second.padStart(2, "0")}/${third}`;
  }
  return normalized;
}
function toNum(v: unknown): number {
  if (v === "" || v == null) return 0;
  if (typeof v === "number") return v;
  return parseFloat(String(v).replace(/,/g, "")) || 0;
}
function fmtNum(n: number) {
  return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

interface MatchReviewItem {
  id: string;
  savedMatch?: SavedMatch;
  cashier?: CashierRow;
  bank?: BankRow;
  title?: string;
  detail?: string;
  addedAt: string;
}

function resolveDebitCredit(
  row: Record<string, unknown>,
  debitCol: string,
  creditCol: string,
  swap: boolean
): { debit: number; credit: number; rawAmount: number; type: "مدفوع" | "مستلم" } {
  const debitRaw = Math.abs(toNum(debitCol ? row[debitCol] : 0));
  const creditRaw = Math.abs(toNum(creditCol ? row[creditCol] : 0));
  const debit = swap ? creditRaw : debitRaw;
  const credit = swap ? debitRaw : creditRaw;

  let rawAmount = 0;
  let type: "مدفوع" | "مستلم" = "مستلم";

  if (debit > 0 && credit === 0) { rawAmount = debit; type = "مستلم"; }
  else if (credit > 0 && debit === 0) { rawAmount = credit; type = "مدفوع"; }
  else if (debit > 0 && credit > 0) { rawAmount = debit; type = "مستلم"; }

  return { debit, credit, rawAmount, type };
}

// ─── تخزين دائم ────────────────────────────────────────────────────────────────
interface AppStorage {
  get(key: string): Promise<{ value: string } | null>;
  set(key: string, value: string): Promise<void>;
  delete(key: string): Promise<void>;
  list(prefix: string): Promise<{ keys: string[] }>;
}

interface StorageWindow extends Window {
  storage?: AppStorage;
}

async function storageGet<T>(key: string, fallback: T): Promise<T> {
  if (centralStorageEnabled) {
    const remote = await centralStorageGet<T>(key);
    if (remote !== undefined) return remote;
    if (isCentralAdmin()) {
      try {
        const legacyValue = localStorage.getItem(key);
        if (legacyValue) {
          const parsed = JSON.parse(legacyValue) as T;
          await centralStorageSet(key, parsed);
          return parsed;
        }
      } catch (error) {
        console.error(`تعذرت ترقية بيانات «${key}» إلى قاعدة البيانات المركزية:`, error);
        throw error;
      }
    }
    return fallback;
  }
  try {
    const w = window as StorageWindow;
    if (w?.storage?.get) {
      const res = await w.storage.get(key);
      return res ? (JSON.parse(res.value) as T) : fallback;
    }
  } catch { /* المفتاح غير موجود أو خطأ - نكمل على localStorage */ }
  try {
    const raw = localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch {
    return fallback;
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
    const w = window as StorageWindow;
    if (w?.storage?.set) {
      await w.storage.set(key, JSON.stringify(value));
      return;
    }
  } catch (e) { console.error("storage.set فشل، سيتم استخدام localStorage:", e); }
  try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) { console.error("فشل الحفظ نهائياً:", e); }
}
async function storageDelete(key: string): Promise<void> {
  if (centralStorageEnabled) {
    await centralStorageDelete(key);
    return;
  }
  try {
    const w = window as StorageWindow;
    if (w?.storage?.delete) { await w.storage.delete(key); return; }
  } catch { /* تجاهل */ }
  try { localStorage.removeItem(key); } catch { /* تجاهل */ }
}

function normalizeIdSet(value: unknown): Set<number> {
  if (value instanceof Set) {
    return new Set(Array.from(value).map(toNum).filter(Number.isFinite));
  }
  if (Array.isArray(value)) {
    return new Set(value.map(toNum).filter(Number.isFinite));
  }
  if (value && typeof value === "object") {
    // Older sessions could contain a JSON-serialized Set, which becomes an
    // object instead of an iterable array. Object keys preserve any numeric ids.
    return new Set(Object.keys(value).map(toNum).filter(Number.isFinite));
  }
  return new Set();
}

function normalizeSpecialKeySet(value: unknown): Set<string> {
  if (Array.isArray(value)) {
    return new Set(value.map(item => typeof item === "string"
      ? item
      : Number.isFinite(Number(item)) ? `legacy:${Number(item)}` : ""
    ).filter(Boolean));
  }
  return new Set();
}

// ─── Visa detection ────────────────────────────────────────────────────────────
const VISA_NUM_RE = new RegExp("[/-]\\s*(\\d{4})(?!\\d)");
function extractVisaNumber(rawName: string): string | null {
  const m = rawName.match(VISA_NUM_RE);
  return m ? m[1] : null;
}
function extractVisaName(rawName: string): string {
  return rawName.replace(VISA_NUM_RE, "").replace(/\s+/g, " ").trim();
}

// ─── Flags ────────────────────────────────────────────────────────────────────
const ARITHMETIC_RE = /(\d+(?:\.\d+)?)(\+(\d+(?:\.\d+)?))+/;
function evalSum(expr: string): number {
  return expr.split("+").reduce((acc, p) => acc + parseFloat(p.trim()), 0);
}
function parseCashierName(raw: string): { name: string; notes: string } {
  let notes = "";
  let name = raw.trim();
  const arithmeticMatch = name.match(ARITHMETIC_RE);
  if (arithmeticMatch) { notes += " " + arithmeticMatch[0]; name = name.replace(ARITHMETIC_RE, ""); }
  return { name: name.replace(/\s+/g, " ").trim(), notes: notes.trim() };
}
function parseNotes(notes: string) {
  const match = notes.match(ARITHMETIC_RE);
  const splitExpr = match ? match[0] : null;
  const matchAmount = splitExpr ? evalSum(splitExpr) : null;
  return { splitExpr, matchAmount };
}

function isMahmoudWalletCashier(row: Pick<CashierRow, "name" | "rawName">): boolean {
  const text = `${row.rawName} ${row.name}`.normalize("NFKC").replace(/\s+/g, " ");
  return /محفظة[\s/\\:،-]*(?:محمود|شخصية)/i.test(text) ||
    /mahmoud[\s/-]*wallet|wallet[\s/-]*mahmoud/i.test(text);
}

function manualEditIssue(note?: string): string | undefined {
  const trimmed = note?.trim();
  return trimmed && /(مبلغ|قيمة|مالي|تاريخ|نوع الحساب|مدين|دائن|خصم)/i.test(trimmed) ? trimmed : undefined;
}

function isTextEntryTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  if (target.isContentEditable || target.tagName === "TEXTAREA" || target.tagName === "SELECT") return true;
  return target instanceof HTMLInputElement &&
    !["checkbox", "radio", "button", "submit"].includes(target.type);
}

// ─── Name matching ────────────────────────────────────────────────────────────
const AR_EN: Record<string, string> = {
  "ا":"a","أ":"a","إ":"a","آ":"a","ى":"a","ة":"h","ب":"b","ت":"t","ث":"th",
  "ج":"j","ح":"h","خ":"kh","د":"d","ذ":"th","ر":"r","ز":"z","س":"s","ش":"sh",
  "ص":"s","ض":"d","ط":"t","ظ":"z","ع":"a","غ":"g","ف":"f","ق":"q","ك":"k",
  "ل":"l","م":"m","ن":"n","ه":"h","و":"w","ي":"i","ئ":"i","ء":"","ّ":"",
  "َ":"","ُ":"","ِ":"","ْ":"","ً":"","ٌ":"","ٍ":"",
};
function normName(s: string) {
  const c = s.toLowerCase()
    .replace(/mahmoud/g, "mhmd")
    .replace(/raefet/g, "raft")
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
  if (a.length > b.length) [a, b] = [b, a];
  let previous = Array.from({ length: a.length + 1 }, (_, index) => index);
  let current = new Array<number>(a.length + 1);
  for (let row = 1; row <= b.length; row++) {
    current[0] = row;
    for (let column = 1; column <= a.length; column++) {
      current[column] = a[column - 1] === b[row - 1]
        ? previous[column - 1]
        : 1 + Math.min(previous[column - 1], current[column - 1], previous[column]);
    }
    [previous, current] = [current, previous];
  }
  return previous[a.length];
}
function levenshteinWithin(a: string, b: string, limit: number): boolean {
  if (Math.abs(a.length - b.length) > limit) return false;
  let previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  let current = new Array<number>(b.length + 1);
  for (let row = 1; row <= a.length; row++) {
    current[0] = row;
    let rowMinimum = current[0];
    for (let column = 1; column <= b.length; column++) {
      current[column] = a[row - 1] === b[column - 1]
        ? previous[column - 1]
        : 1 + Math.min(previous[column - 1], current[column - 1], previous[column]);
      rowMinimum = Math.min(rowMinimum, current[column]);
    }
    if (rowMinimum > limit) return false;
    [previous, current] = [current, previous];
  }
  return previous[b.length] <= limit;
}
function nameSimilarityNormalized(na: string, nb: string): number {
  if (!na || !nb) return 0;
  if (na === nb) return 1;
  return 1 - levenshtein(na, nb) / Math.max(na.length, nb.length);
}
function nameSim(a: string, b: string): number {
  const na=normName(a), nb=normName(b);
  return nameSimilarityNormalized(na, nb);
}

function wordTypoMatch(a: string, b: string): boolean {
  if (!a || !b) return false;
  if (a === b) return true;
  const maxLen = Math.max(a.length, b.length);
  const threshold = maxLen <= 4 ? 1 : maxLen <= 7 ? 2 : Math.ceil(maxLen * 0.3);
  return levenshteinWithin(a, b, threshold);
}

function nameTokens(s: string): string[] {
  return normName(s).split(" ").filter(t => t.length > 1 && t !== "al");
}

function normalizePlatformRefSeed(accountType: string): string {
  const value = (accountType || "")
    .normalize("NFKC")
    .replace(/[\u064B-\u065F\u0670\u0640]/g, "")
    .replace(/[أإآ]/g, "ا")
    .toLowerCase();

  if (/بنك\s*فلسطين|bank\s*of\s*palestine|palestine/.test(value)) return "BP";
  if (/بنك\s*القدس|bank\s*of\s*jerusalem|jerusalem|quds/.test(value)) return "BQ";
  if (/بنك\s*الامريكي|bank\s*of\s*america|america|امريكا/.test(value)) return "BA";
  if (/بنك\s*الاسترالي|bank\s*of\s*australia|australia|استراليا/.test(value)) return "BAU";
  if (/محفظة\s*(محمود|شخصية)|mahmoud|محمود|personal\s*wallet/.test(value)) return "WM";
  if (/محفظة|wallet|بال\s*بي|pal\s*pay|palpay/.test(value)) return "WP";
  if (/جوال\s*بي|jawwal\s*pay|jawwal|جوال/.test(value)) return "WJ";
  if (/فيزا|visa/.test(value)) return "VI";

  const words = value
    .split("")
    .map(character => AR_EN[character] ?? character)
    .join("")
    .replace(/[^a-z\s]/g, " ")
    .split(/\s+/)
    .filter(part => Boolean(part) && !["bank", "bnk", "of", "the", "al"].includes(part));
  const isBank = /بنك|\bbank\b/.test(value);
  const initials = words.map(part => part[0].toUpperCase());
  const letters = isBank ? `B${initials[0] ?? "X"}` : initials.slice(0, 2).join("");

  return letters || "XX";
}

function generateDefaultReference(accountType: string, rowKind: "invoice" | "transfer", rowIndex: number): string {
  const prefix = normalizePlatformRefSeed(accountType);
  const kind = rowKind === "invoice" ? "R" : "T";
  const ordinal = Math.max(1, Math.abs(rowIndex) || 1);
  return `${prefix}${kind}${ordinal}`;
}

// ─── Types ────────────────────────────────────────────────────────────────────
export interface BankRow {
  id: number; date: string; description: string;
  debit: number; credit: number; rawAmount: number;
  totalAmount?: number;
  movementType?: string;
  type: "مدفوع" | "مستلم";
  accountType: string;
  ref: string;
  defaultRef?: string;
  orig: Record<string,unknown>;
  _fromHeld?: boolean;
  _fileSessionId?: number;
}
export interface CashierRow {
  id: number; rawName: string; name: string; notes: string;
  splitExpr: string|null;
  debit: number; credit: number; amount: number; matchAmount: number;
  type: "مدفوع" | "مستلم";
  accountType: string;
  userId?: string;
  registrationTime?: string;
  invoiceTotal?: number;
  cashAmount?: number;
  checkAmount?: number;
  debtAmount?: number;
  discountAmount?: number;
  cardAmount?: number;
  customerNumber?: string;
  invoiceNumber?: string;
  category?: string;
  ref: string;
  defaultRef?: string;
  date: string; orig: Record<string,unknown>;
  _fromHeld?: boolean;
  _fileSessionId?: number;
}

function defaultReferenceForRow(row: BankRow | CashierRow): string {
  if (row.defaultRef) return row.defaultRef;
  const isBank = "rawAmount" in row;
  const idRanges = isBank
    ? [0, 1_000_000, 4_000_000, 10_000_000]
    : [0, 2_000_000, 3_000_000, 5_000_000, 20_000_000];
  const offset = idRanges.filter(start => row.id >= start).pop() ?? 0;
  const rowIndex = offset >= 10_000_000 ? row.id % 100_000 : row.id - offset;
  return generateDefaultReference(row.accountType, isBank ? "transfer" : "invoice", rowIndex + 1);
}

function DefaultReferenceButton({ row, onOpen }: {
  row: BankRow | CashierRow;
  onOpen: (row: BankRow | CashierRow) => void;
}) {
  const reference = defaultReferenceForRow(row);
  return (
    <button
      type="button"
      onClick={event => { event.stopPropagation(); onOpen(row); }}
      title="اعرض أماكن وجود هذه الحركة"
      className="inline-flex items-center rounded border border-indigo-200 bg-indigo-50 px-1.5 py-0.5 font-mono text-[10px] font-semibold text-indigo-700 hover:bg-indigo-100"
    >
      {reference}
    </button>
  );
}

function wasSpecialCashierRejected(rejected: Set<string>, row: CashierRow): boolean {
  return rejected.has(transactionKey(row)) ||
    ((row._fileSessionId ?? 0) === 0 && rejected.has(`legacy:${row.id}`));
}

type PlatformRole = "invoice" | "transfer" | "other";
interface PlatformSource {
  fileName: string;
  headers: string[];
  rows: Record<string, unknown>[];
  map: { date: string; description: string; debit: string; credit: string; ref: string; userId?: string };
  role: PlatformRole;
  otherName: string;
  otherRole: "invoice" | "transfer";
  fileSessionId: number;
}
interface CustomPlatform {
  id: number;
  name: string;
  transfer: PlatformSource | null;
  invoice: PlatformSource | null;
}
function CustomPlatformNameInput({ value, onCommit }: { value: string; onCommit: (value: string) => void }) {
  const [draft, setDraft] = useState(value);
  useEffect(() => setDraft(value), [value]);
  const commit = () => {
    const trimmed = draft.trim();
    if (trimmed !== value) onCommit(trimmed);
  };
  return (
    <input
      value={draft}
      placeholder="مثال: بنك القدس"
      onChange={event => setDraft(event.target.value)}
      onBlur={commit}
      onKeyDown={event => { if (event.key === "Enter") event.currentTarget.blur(); }}
      className="mt-1 w-full rounded-lg border bg-slate-50 px-3 py-2 text-sm"
    />
  );
}
interface FileSnapshot {
  label: string;
  headers: string[];
  rows: Record<string, unknown>[];
  savedAt: string;
  fileCount?: number;
}
interface ManualMatchGroup {
  id: string;
  banks: BankRow[];
  cashiers: CashierRow[];
  note?: string;
}

// ─── Saved matches ──────────────────────────────────────────────────────────
interface SavedMatch {
  id: string;
  cashierId: number;
  bankId: number;
  cashierSessionId?: number;
  bankSessionId?: number;
  cashierDefaultRef?: string;
  bankDefaultRef?: string;
  cashierName: string;
  bankDesc: string;
  cashierIdentity?: string;
  bankIdentity?: string;
  cashierDate?: string;
  bankDate?: string;
  cashierAmount?: number;
  cashierRawName?: string;
  amount: number;
  type: "مدفوع" | "مستلم";
  date: string;
  savedAt: string;
  note?: string;
  isAmountDiff?: boolean;
  isPhoneDiff?: boolean;
  isNameDiff?: boolean;
  isManual?: boolean;
  isAccountTypeDiff?: boolean;
  isDateDiff?: boolean;
  bankAccountType?: string;
  cashierAccountType?: string;
  sourceGroupId?: string;
  matchScore?: number;
  editorNotes?: string;
}

function dedupeSavedMatches(matches: SavedMatch[]): SavedMatch[] {
  const seen = new Set<string>();
  return matches.filter(match => {
    if (seen.has(match.id)) return false;
    seen.add(match.id);
    return true;
  });
}

function normalizeMovementIdentityText(value: string): string {
  return value.normalize("NFKC")
    .replace(/[\u064B-\u065F\u0670\u0640]/g, "")
    .replace(/[أإآ]/g, "ا")
    .toLocaleLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "");
}

function normalizeMovementIdentityDate(value: string): string {
  const normalized = value.normalize("NFKC")
    .replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, digit => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)))
    .trim();
  const parts = normalized.match(/^(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})/);
  if (!parts) return normalizeMovementIdentityText(normalized);
  const [, first, second, third] = parts;
  const [year, month, day] = first.length === 4 ? [first, second, third] : [third, second, first];
  return `${year.padStart(4, "0")}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
}

function actualMovementReference(reference: string, description: string): string {
  const explicit = reference.trim();
  if (explicit) return normalizeMovementIdentityText(explicit);
  const match = description.match(
    /(?:رقم\s*(?:الحركة|المعاملة|القيد|المرجع)|(?:transaction|movement|reference)\s*(?:number|no|id)?)\s*[:#：/-]?\s*([a-z\d-]{5,})|\/(\d{8,})\s*\/0\b|\b(pacs008-[a-z\d-]+|pibc[a-z\d-]+)\b/i,
  );
  return normalizeMovementIdentityText(match?.[1] ?? match?.[2] ?? match?.[3] ?? "");
}

function movementIdentity(row: BankRow | CashierRow): string {
  const isBank = "rawAmount" in row;
  const description = isBank ? row.description : row.rawName;
  const reference = actualMovementReference(row.ref, description);
  return JSON.stringify([
    normalizeMovementIdentityText(row.accountType),
    normalizeMovementIdentityDate(row.date),
    row.type,
    Math.round((isBank ? row.rawAmount : row.amount) * 100),
    reference ? `ref:${reference}` : `description:${normalizeMovementIdentityText(description)}`,
  ]);
}

function savedBankSnapshotMatches(match: SavedMatch, row: BankRow): boolean {
  return (!match.bankDesc || row.description === match.bankDesc) &&
    Math.abs(row.rawAmount - match.amount) <= 0.01 &&
    row.type === match.type &&
    (!match.bankAccountType || row.accountType === match.bankAccountType) &&
    (!match.bankDate || normalizeMovementIdentityDate(row.date) === normalizeMovementIdentityDate(match.bankDate));
}

function savedCashierSnapshotMatches(match: SavedMatch, row: CashierRow): boolean {
  return row.name === match.cashierName &&
    row.type === match.type &&
    (!match.cashierAccountType || row.accountType === match.cashierAccountType) &&
    (!match.cashierDate || normalizeMovementIdentityDate(row.date) === normalizeMovementIdentityDate(match.cashierDate)) &&
    (match.cashierRawName == null || row.rawName === match.cashierRawName) &&
    (match.cashierAmount == null || Math.abs(row.amount - match.cashierAmount) <= 0.01);
}

function savedBankCandidates(match: SavedMatch, rows: BankRow[]): BankRow[] {
  if (match.bankIdentity) {
    return rows.filter(row => movementIdentity(row) === match.bankIdentity);
  }
  const reference = actualMovementReference("", match.bankDesc);
  if (reference) {
    const referenceMatches = rows.filter(row => actualMovementReference(row.ref, row.description) === reference);
    if (referenceMatches.length) return referenceMatches;
  }
  const description = normalizeMovementIdentityText(match.bankDesc);
  const descriptionWithoutNumbers = normalizeMovementIdentityText(match.bankDesc.replace(/\d+/g, ""));
  return rows.filter(row =>
    row.type === match.type &&
    Math.abs(row.rawAmount - match.amount) <= 0.01 &&
    (!match.bankAccountType || row.accountType === match.bankAccountType) &&
    (normalizeMovementIdentityText(row.description) === description ||
      normalizeMovementIdentityText(row.description.replace(/\d+/g, "")) === descriptionWithoutNumbers)
  );
}

function savedCashierCandidates(match: SavedMatch, rows: CashierRow[]): CashierRow[] {
  if (match.cashierIdentity) {
    return rows.filter(row => movementIdentity(row) === match.cashierIdentity);
  }
  return rows.filter(row =>
    row.name === match.cashierName &&
    row.type === match.type &&
    (!match.cashierAccountType || row.accountType === match.cashierAccountType) &&
    (match.cashierDate == null || normalizeMovementIdentityDate(row.date) === normalizeMovementIdentityDate(match.cashierDate)) &&
    (match.cashierRawName == null || row.rawName === match.cashierRawName) &&
    (match.cashierAmount == null || Math.abs(row.amount - match.cashierAmount) <= 0.01)
  );
}

function resolveSavedCashierRow(match: SavedMatch, rows: CashierRow[]): CashierRow | undefined {
  if (match.cashierIdentity) {
    const identityMatches = savedCashierCandidates(match, rows);
    return identityMatches.length === 1 ? identityMatches[0] : undefined;
  }

  let candidates = rows.filter(row => row.id === match.cashierId && savedCashierSnapshotMatches(match, row));
  if (match.cashierSessionId != null) {
    candidates = candidates.filter(row => (row._fileSessionId ?? 0) === match.cashierSessionId);
  }
  return candidates.length === 1 ? candidates[0] : undefined;
}

function resolveSavedBankRow(match: SavedMatch, rows: BankRow[]): BankRow | undefined {
  if (match.bankIdentity) {
    const identityMatches = savedBankCandidates(match, rows);
    return identityMatches.length === 1 ? identityMatches[0] : undefined;
  }

  let candidates = savedBankCandidates(match, rows).filter(row => savedBankSnapshotMatches(match, row));
  if (match.bankSessionId != null) {
    const sameSessionAndRow = candidates.filter(row =>
      row.id === match.bankId && (row._fileSessionId ?? 0) === match.bankSessionId
    );
    if (sameSessionAndRow.length === 1) return sameSessionAndRow[0];
  } else {
    const sameLegacyRow = candidates.filter(row => row.id === match.bankId);
    if (sameLegacyRow.length === 1) return sameLegacyRow[0];
  }
  return candidates.length === 1 ? candidates[0] : undefined;
}

function matchesPhoneAndAmount(cashier: CashierRow, bank: BankRow): boolean {
  const cashierPhone = extractMobileNumber(`${cashier.rawName} ${cashier.name} ${cashier.ref}`);
  const bankPhone = extractMobileNumber(`${bank.description} ${bank.ref}`);
  return !!cashierPhone && cashierPhone === bankPhone &&
    Math.abs(bank.rawAmount - cashier.matchAmount) <= 0.01;
}

function movementPhones(cashier: CashierRow, bank: BankRow) {
  return {
    cashierPhone: extractMobileNumber(`${cashier.rawName} ${cashier.name} ${cashier.ref}`),
    bankPhone: extractMobileNumber(`${bank.description} ${bank.ref}`),
  };
}

function hasPhoneDifference(cashier: CashierRow, bank: BankRow): boolean {
  const { cashierPhone, bankPhone } = movementPhones(cashier, bank);
  return !!bankPhone && cashierPhone !== bankPhone;
}

function bankSplitAllocationGroups(split: BankSplit): Map<string, SplitAllocation[]> {
  const groups = new Map<string, SplitAllocation[]>();
  split.allocations.forEach(allocation => {
    if (allocation.kind !== "company" || allocation.cashierId == null) return;
    const key = `${allocation.cashierSessionId ?? 0}:${allocation.cashierId}`;
    groups.set(key, [...(groups.get(key) ?? []), allocation]);
  });
  return groups;
}

function bankSplitVirtualId(split: BankSplit, cashierKey: string, allocations: SplitAllocation[]): number {
  if (allocations.length === 1) return syntheticTransactionId(`${split.id}:${allocations[0].id}`);
  return syntheticTransactionId(`${split.id}:cashier:${cashierKey}`);
}

// ─── Visa items ──────────────────────────────────────────────────────────────
interface VisaItem {
  id: string;
  cashierId: number;
  cashierSessionId?: number;
  defaultRef?: string;
  rawName: string;
  name: string;
  visaNumber: string;
  amount: number;
  type: "مدفوع" | "مستلم";
  date: string;
  movedAt: string;
  note?: string;
  source?: "pending" | "unmatched" | "saved" | "manual";
}

interface HeldItem {
  id: string;
  kind: "bank" | "cashier";
  refId: number;
  fileSessionId: number;
  data: BankRow | CashierRow;
  heldAt: string;
  note?: string;
}

function dedupeHeldItems(items: HeldItem[]): HeldItem[] {
  const seen = new Set<string>();
  return items.filter((item) => {
    const key = `${item.kind}:${item.fileSessionId}:${item.refId}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function getNonReconciliationBankReason(row: BankRow): string | null {
  const description = row.description
    .normalize("NFKC")
    .replace(/[\u064B-\u065F\u0670\u0640]/g, "")
    .replace(/[أإآ]/g, "ا")
    .toLowerCase();
  if (/تحويل\s+الى\s+حساب\s+(?:ال)?شيكل/.test(description) ||
      /transfer\s+to\s+account\b.*\b(?:shiekel|shekel|sheqel|nis|ils)\b/i.test(description)) {
    return "تحويل إلى حساب الشيكل";
  }
  if (/عمولة\s+(?:ادارة|اداره)\s+حساب/.test(description) ||
      /\baccount\s+management\s+fee\b|\baccount\s+fee\b/i.test(description)) {
    return "عمولة إدارة الحساب";
  }
  if (/نقاط\s*بيع/.test(description) && /عمول|صرف/.test(description)) {
    return "عمليات نقاط البيع والعمولات";
  }
  if ((/مشتريات/.test(description) && /عمول/.test(description) && /تجار|بطاق/.test(description)) ||
      /merchant.*(?:purchase|commission)|(?:purchase|commission).*merchant/.test(description)) {
    return "عمولات المشتريات والتجار";
  }
  if (/تحويل\s*(الكتروني|إلكتروني)?\s*موبايل/.test(description)) {
    const parties = description.match(/من\s+(.+?)\s+الى\s+(.+?)(?:[,/؛]|$)/);
    const normalizeParty = (party: string) => party
      .replace(/حساب|رقم|بنكي|محفظة|تحويل|الكتروني|موبايل|\d+/g, "")
      .replace(/[\s.,/-]/g, "");
    if (parties?.[1] && normalizeParty(parties[1]) === normalizeParty(parties[2])) {
      return "تحويل بين حسابات صاحب الحساب نفسه";
    }
  }
  return null;
}

function isMerchantPurchaseDescription(description: string): boolean {
  const normalized = description.normalize("NFKC").replace(/[\u064B-\u065F\u0670\u0640]/g, "").toLowerCase();
  return (/مشتريات/.test(normalized) && /عمول/.test(normalized) && /تجار|بطاق/.test(normalized)) ||
    /merchant.*(?:purchase|commission)|(?:purchase|commission).*merchant/.test(normalized);
}

function sameBankMovement(a: BankRow, b: BankRow): boolean {
  return a.date === b.date && a.description === b.description && a.type === b.type &&
    a.debit === b.debit && a.credit === b.credit && a.rawAmount === b.rawAmount &&
    a.accountType === b.accountType && a.ref === b.ref;
}

function sameCashierMovement(a: CashierRow, b: CashierRow): boolean {
  return a.date === b.date && a.name === b.name && a.rawName === b.rawName &&
    a.type === b.type && a.amount === b.amount && a.matchAmount === b.matchAmount &&
    a.debit === b.debit && a.credit === b.credit && a.accountType === b.accountType;
}

function normAccountType(s: string): string {
  return (s || "").toString().trim().toLowerCase().replace(/\s+/g, " ");
}
function accountTypesDiffer(bankType: string, cashierType: string): boolean {
  const a = normAccountType(bankType), b = normAccountType(cashierType);
  if (!a || !b) return false;
  return a !== b;
}
function transactionDatesDiffer(bankDate: string, cashierDate: string): boolean {
  const candidates = (value: string): Set<string> => {
    const normalized = String(value || "").normalize("NFKC")
      .replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
      .replace(/[۰-۹]/g, digit => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)))
      .replace(/[\u200e\u200f\u202a-\u202e]/g, "")
      .trim();
    const parts = normalized.match(/^(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})/);
    if (!parts) return new Set(normalized ? [normalized.replace(/\s+/g, "")] : []);
    const [, first, second, third] = parts;
    const year = first.length === 4 ? first : third;
    const a = Number(first.length === 4 ? second : first);
    const b = Number(first.length === 4 ? third : second);
    if (first.length === 4) return new Set([`${year}-${String(a).padStart(2, "0")}-${String(b).padStart(2, "0")}`]);
    const dates = new Set<string>();
    if (a >= 1 && a <= 31 && b >= 1 && b <= 12) dates.add(`${year}-${String(b).padStart(2, "0")}-${String(a).padStart(2, "0")}`);
    if (b >= 1 && b <= 31 && a >= 1 && a <= 12) dates.add(`${year}-${String(a).padStart(2, "0")}-${String(b).padStart(2, "0")}`);
    return dates;
  };
  const bank = candidates(bankDate);
  const cashier = candidates(cashierDate);
  if (!bank.size || !cashier.size) return false;
  return !Array.from(bank).some(date => cashier.has(date));
}

// ─── Advanced match check ───────────────────────────────────────────────────
function advancedMatchCheckSingle(cashierName: string, bankDesc: string): {
  isMatch: boolean;
  isApprox: boolean;
  matchType: "exact" | "firstSecond" | "fourthName" | "typo" | "none";
  matchedTokens?: string[];
} {
  return advancedMatchCheckTokens(nameTokens(cashierName), nameTokens(bankDesc));
}

function advancedMatchCheckTokens(cT: string[], bT: string[]): {
  isMatch: boolean;
  isApprox: boolean;
  matchType: "exact" | "firstSecond" | "fourthName" | "typo" | "none";
  matchedTokens?: string[];
} {
  if (!cT.length || !bT.length) return { isMatch: false, isApprox: false, matchType: "none" };

  const squished = (arr: string[]) => arr.slice().sort().join("");
  if (squished(cT) === squished(bT)) {
    return { isMatch: true, isApprox: false, matchType: "exact" };
  }

  if (cT.length >= 2 && bT.length >= 2) {
    const firstMatch = wordTypoMatch(cT[0], bT[0]);
    const lastMatch = wordTypoMatch(cT[cT.length - 1], bT[bT.length - 1]);
    if (firstMatch && lastMatch) {
      return {
        isMatch: true,
        isApprox: cT[0] !== bT[0] || cT[cT.length - 1] !== bT[bT.length - 1],
        matchType: "firstSecond",
        matchedTokens: [cT[0], cT[cT.length - 1]]
      };
    }
  }

  if (cT.length >= 2 && bT.length >= 2) {
    const firstMatch = wordTypoMatch(cT[0], bT[0]);
    const secondMatch = wordTypoMatch(cT[1], bT[1]);
    if (firstMatch && secondMatch) {
      const exactFirst = cT[0] === bT[0];
      const exactSecond = cT[1] === bT[1];
      return {
        isMatch: true,
        isApprox: !(exactFirst && exactSecond),
        matchType: "firstSecond",
        matchedTokens: [cT[0], cT[1]]
      };
    }
  }

  if (cT.length >= 4 && bT.length >= 4) {
    if (wordTypoMatch(cT[3], bT[3])) {
      return {
        isMatch: true,
        isApprox: cT[3] !== bT[3],
        matchType: "fourthName",
        matchedTokens: [cT[3]]
      };
    }
  }

  if (cT.length === 1 || bT.length === 1) {
    const shortSide = cT.length === 1 ? cT : bT;
    const longSide = cT.length === 1 ? bT : cT;
    if (wordTypoMatch(shortSide[0], longSide[0])) {
      return {
        isMatch: true,
        isApprox: shortSide[0] !== longSide[0],
        matchType: "typo",
        matchedTokens: [shortSide[0]]
      };
    }
  }

  for (const ct of cT) {
    for (const bt of bT) {
      if (wordTypoMatch(ct, bt)) {
        return {
          isMatch: true,
          isApprox: true,
          matchType: "typo",
          matchedTokens: [ct, bt]
        };
      }
    }
  }

  return { isMatch: false, isApprox: false, matchType: "none" };
}

function advancedMatchCheck(cashierName: string, bankDesc: string): {
  isMatch: boolean;
  isApprox: boolean;
  matchType: "exact" | "firstSecond" | "fourthName" | "typo" | "none";
  matchedTokens?: string[];
} {
  const bankSegments = bankDesc.split("/").map(segment => segment.trim()).filter(Boolean);
  if (bankSegments.length <= 1) return advancedMatchCheckSingle(cashierName, bankDesc);
  const results = bankSegments.map(segment => advancedMatchCheckSingle(cashierName, segment));
  return results
    .filter(result => result.isMatch)
    .sort((a, b) => Number(a.isApprox) - Number(b.isApprox))[0] ||
    { isMatch: false, isApprox: false, matchType: "none" };
}

function smartNameKey(value: string): string {
  const tokens = nameTokens(value);
  return tokens.length >= 2 ? `${tokens[0]} ${tokens[tokens.length - 1]}` : tokens.join(" ");
}

function extractMobileNumber(value: string): string | null {
  const asciiDigits = value.replace(/[٠-٩۰-۹]/g, digit => {
    const code = digit.charCodeAt(0);
    return String(code >= 0x06f0 ? code - 0x06f0 : code - 0x0660);
  });
  const matches = asciiDigits.match(/(?:^|[^\d])((?:(?:00970|970)\D*)?0?\D*5(?:[\s()./-]*\d){8})(?=$|[^\d])/g) ?? [];
  const normalized = [...new Set(matches.map(match => {
    const candidate = match.replace(/\D/g, "");
    const local = candidate.startsWith("00970") ? candidate.slice(5)
      : candidate.startsWith("970") ? candidate.slice(3)
      : candidate;
    const phone = local.startsWith("0") ? local : `0${local}`;
    return /^05\d{8}$/.test(phone) ? phone : "";
  }).filter(Boolean))];
  return normalized.length === 1 ? normalized[0] : null;
}

function jawwalPhoneDigits(value: string): Array<{ digits: string; partial: boolean }> {
  const fullPhone = extractMobileNumber(value);
  if (fullPhone) return [{ digits: fullPhone, partial: false }];
  const asciiDigits = value.replace(/[٠-٩۰-۹]/g, digit => {
    const code = digit.charCodeAt(0);
    return String(code >= 0x06f0 ? code - 0x06f0 : code - 0x0660);
  });
  const fragments = [...new Set([...asciiDigits.matchAll(/(?:\/|-|\()\s*(\d{2,4})(?=$|[\s/),.-])/g)]
    .map(match => match[1]))];
  return fragments.length === 1 ? [{ digits: fragments[0], partial: true }] : [];
}

function jawwalPhoneSuffixes(value: string): string[] {
  return jawwalPhoneDigits(value).flatMap(({ digits, partial }) =>
    partial ? [digits] : [digits.slice(-4), digits.slice(-3), digits.slice(-2)]
  );
}

function jawwalPhoneSuffixMatches(cashierText: string, bankText: string): boolean {
  const cashierPhones = jawwalPhoneDigits(cashierText);
  const bankPhones = jawwalPhoneDigits(bankText);
  return cashierPhones.some(cashierPhone => bankPhones.some(bankPhone => {
    if (cashierPhone.partial && bankPhone.partial) return cashierPhone.digits === bankPhone.digits;
    if (cashierPhone.partial) return bankPhone.digits.endsWith(cashierPhone.digits);
    if (bankPhone.partial) return cashierPhone.digits.endsWith(bankPhone.digits);
    return cashierPhone.digits === bankPhone.digits;
  }));
}

function isJawwalAccount(accountType: string, configuredName = "جوال بي"): boolean {
  return accountType.trim().toLocaleLowerCase() === configuredName.trim().toLocaleLowerCase() ||
    /جوال\s*بي|jawwal\s*pay/i.test(accountType);
}

function uniqueTransactionRows<T extends { id: number; _fileSessionId?: number }>(rows: T[]): T[] {
  const unique = new Map<string, T>();
  rows.forEach(row => {
    const key = transactionKey(row);
    if (!unique.has(key)) unique.set(key, row);
  });
  return [...unique.values()];
}

function transactionPairKey(cashier: CashierRow, bank: BankRow): string {
  return `${transactionKey(cashier)}-${transactionKey(bank)}`;
}

function wasPairRejected(rejected: Set<string>, cashier: CashierRow, bank: BankRow): boolean {
  if (rejected.has(transactionPairKey(cashier, bank))) return true;
  return (cashier._fileSessionId ?? 0) === 0 && (bank._fileSessionId ?? 0) === 0
    && rejected.has(`${cashier.id}-${bank.id}`);
}

function smartAliasTarget(cashierName: string, bankDescription: string): string {
  return bankDescription.split("/")
    .map(segment => segment.trim())
    .filter(Boolean)
    .sort((a, b) => nameSim(cashierName, b) - nameSim(cashierName, a))[0] || bankDescription;
}

interface AmountIndexEntry<T> {
  row: T;
  amount: number;
  order: number;
}

function indexByAmountAndType<T>(rows: T[], getAmount: (row: T) => number, getType: (row: T) => string) {
  const index = new Map<string, AmountIndexEntry<T>[]>();
  rows.forEach((row, order) => {
    const type = getType(row);
    const entries = index.get(type) ?? [];
    entries.push({ row, amount: getAmount(row), order });
    index.set(type, entries);
  });
  index.forEach(entries => entries.sort((a, b) => a.amount - b.amount || a.order - b.order));
  return index;
}

function amountRange<T>(entries: AmountIndexEntry<T>[], target: number, tolerance: number): AmountIndexEntry<T>[] {
  const roundingMargin = Number.EPSILON * Math.max(1, Math.abs(target), Math.abs(tolerance)) * 4;
  let low = 0;
  let high = entries.length;
  const minimum = target - tolerance - roundingMargin;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (entries[middle].amount < minimum) low = middle + 1;
    else high = middle;
  }
  const start = low;
  low = start;
  high = entries.length;
  const maximum = target + tolerance + roundingMargin;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (entries[middle].amount <= maximum) low = middle + 1;
    else high = middle;
  }
  return entries.slice(start, low).sort((a, b) => a.order - b.order);
}

function hasAmountWithin<T>(entries: AmountIndexEntry<T>[], target: number, tolerance: number): boolean {
  const roundingMargin = Number.EPSILON * Math.max(1, Math.abs(target), Math.abs(tolerance)) * 4;
  let low = 0;
  let high = entries.length;
  const minimum = target - tolerance - roundingMargin;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (entries[middle].amount < minimum) low = middle + 1;
    else high = middle;
  }
  while (low < entries.length && entries[low].amount <= target + tolerance + roundingMargin) {
    if (Math.abs(entries[low].amount - target) <= tolerance) return true;
    low += 1;
  }
  return false;
}

interface ReconcilePerformanceMetrics {
  possiblePairs: number;
  amountCandidates: number;
  nameChecks: number;
}

// ─── Reconcile ───────────────────────────────────────────��──────────────────
function reconcile(
  bank: BankRow[],
  cashier: CashierRow[],
  manualGroups: ManualMatchGroup[],
  savedMatches: SavedMatch[],
  rejectedPairs: Set<string>,
  visaCashierIds: Set<string>,
  jawwalPayCashierIds: Set<string>,
  mahmoudWalletCashierIds: Set<string>,
  amountTolerancePercent: number = 0.5,
  aliases: Record<string, string> = {},
  jawwalAccountType = "جوال بي",
  performanceMetrics?: ReconcilePerformanceMetrics,
): MatchResult[] {
  bank = uniqueTransactionRows(bank);
  cashier = uniqueTransactionRows(cashier);
  if (performanceMetrics) {
    performanceMetrics.possiblePairs = bank.length * cashier.length;
    performanceMetrics.amountCandidates = 0;
    performanceMetrics.nameChecks = 0;
  }
  const results: MatchResult[] = [];
  const bankKey = (row: BankRow) => transactionKey(row);
  const cashierKey = (row: CashierRow) => transactionKey(row);
  const usedBank = new Set<string>();
  const usedCashier = new Set<string>();

  const isJawwalRow = (row: BankRow | CashierRow) => isJawwalAccount(row.accountType, jawwalAccountType);
  const eligibleSavedMatches = savedMatches.flatMap(match => {
    const bankRow = resolveSavedBankRow(match, bank);
    const cashierRow = resolveSavedCashierRow(match, cashier);
    if (!bankRow || !cashierRow) return [];
    if (match.isManual) return [{ match, bankRow, cashierRow }];
    const jawwalMatch = isJawwalRow(bankRow) || isJawwalRow(cashierRow);
    if (!jawwalMatch) return [{ match, bankRow, cashierRow }];
    const isValid = matchesPhoneAndAmount(cashierRow, bankRow);
    return isValid ? [{ match, bankRow, cashierRow }] : [];
  });
  const claimedCashiers = new Set<string>();
  const claimedBanks = new Set<string>();
  const uniqueSavedMatches = eligibleSavedMatches.filter(({ cashierRow, bankRow }) => {
    const cashierSourceKey = cashierKey(cashierRow);
    const bankSourceKey = bankKey(bankRow);
    if (claimedCashiers.has(cashierSourceKey) || claimedBanks.has(bankSourceKey)) return false;
    claimedCashiers.add(cashierSourceKey);
    claimedBanks.add(bankSourceKey);
    return true;
  });
  const savedKeys = new Set(uniqueSavedMatches.map(({ cashierRow, bankRow }) => `${cashierKey(cashierRow)}-${bankKey(bankRow)}`));
  const savedCashierIds = new Set(uniqueSavedMatches.map(({ cashierRow }) => cashierKey(cashierRow)));
  const savedBankIds = new Set(uniqueSavedMatches.map(({ bankRow }) => bankKey(bankRow)));

  uniqueSavedMatches.forEach(({ match, bankRow, cashierRow }) => {
    results.push({ type: "saved", bank: bankRow, cashier: cashierRow, savedMatch: match });
    usedBank.add(bankKey(bankRow));
    usedCashier.add(cashierKey(cashierRow));
  });

  const jawwalBanksByPhone = new Map<string, BankRow[]>();
  bank.forEach(row => {
    if (usedBank.has(bankKey(row)) || savedBankIds.has(bankKey(row))) return;
    const phone = extractMobileNumber(`${row.description} ${row.ref}`);
    if (!phone) return;
    jawwalBanksByPhone.set(phone, [...(jawwalBanksByPhone.get(phone) ?? []), row]);
  });
  const jawwalCashiersByPhone = new Map<string, CashierRow[]>();
  cashier.forEach(row => {
    if (usedCashier.has(cashierKey(row)) || savedCashierIds.has(cashierKey(row))) return;
    const phone = extractMobileNumber(`${row.rawName} ${row.name} ${row.ref}`);
    if (!phone) return;
    jawwalCashiersByPhone.set(phone, [...(jawwalCashiersByPhone.get(phone) ?? []), row]);
  });
  jawwalCashiersByPhone.forEach((cashiers, phone) => {
    const banks = jawwalBanksByPhone.get(phone);
    if (!banks?.length) return;
    const possiblePairs = cashiers.flatMap(cashier => banks
      .filter(bankRow =>
        (isJawwalRow(cashier) || isJawwalRow(bankRow)) &&
        !usedBank.has(bankKey(bankRow)) &&
        !wasPairRejected(rejectedPairs, cashier, bankRow) &&
        matchesPhoneAndAmount(cashier, bankRow)
      )
      .map(bankRow => ({ cashier, bank: bankRow, amountDiff: Math.abs(bankRow.rawAmount - cashier.matchAmount) })))
      .sort((a, b) => a.amountDiff - b.amountDiff);
    possiblePairs.forEach(({ cashier: cashierRow, bank: bankRow }) => {
      if (usedCashier.has(cashierKey(cashierRow)) || usedBank.has(bankKey(bankRow))) return;
      results.push({
        type: "pending",
        bank: bankRow,
        cashier: cashierRow,
        isApprox: false,
        isExactAmount: true,
        matchType: "phone",
        amountDiff: 0,
        matchScore: 100,
        accountTypeDiff: accountTypesDiffer(bankRow.accountType, cashierRow.accountType),
      });
      usedCashier.add(cashierKey(cashierRow));
      usedBank.add(bankKey(bankRow));
    });

    const remainingCashiers = cashiers.filter(row => !usedCashier.has(cashierKey(row)));
    const remainingBanks = banks.filter(row => !usedBank.has(bankKey(row)));
    const differentAmountPairs = remainingCashiers.flatMap(cashierRow => remainingBanks
      .filter(bankRow =>
        (isJawwalRow(cashierRow) || isJawwalRow(bankRow)) &&
        !wasPairRejected(rejectedPairs, cashierRow, bankRow)
      )
      .map(bankRow => ({
        cashier: cashierRow,
        bank: bankRow,
        amountDiff: Math.abs(bankRow.rawAmount - cashierRow.matchAmount),
      })))
      .sort((a, b) => a.amountDiff - b.amountDiff);
    differentAmountPairs.forEach(({ cashier: cashierRow, bank: bankRow, amountDiff }) => {
      if (usedCashier.has(cashierKey(cashierRow)) || usedBank.has(bankKey(bankRow))) return;
      results.push({
        type: "pending",
        bank: bankRow,
        cashier: cashierRow,
        isApprox: false,
        isExactAmount: amountDiff <= 0.01,
        matchType: "phone",
        amountDiff,
        matchScore: 100,
        accountTypeDiff: accountTypesDiffer(bankRow.accountType, cashierRow.accountType),
      });
      usedCashier.add(cashierKey(cashierRow));
      usedBank.add(bankKey(bankRow));
    });
  });

  for (const suffixLength of [4, 3, 2]) {
    const cashierGroups = new Map<string, CashierRow[]>();
    cashier.forEach(row => {
      if (usedCashier.has(cashierKey(row)) || savedCashierIds.has(cashierKey(row))) return;
      jawwalPhoneDigits(`${row.rawName} ${row.name}`).forEach(({ digits, partial }) => {
        if (partial && digits.length !== suffixLength) return;
        if (!partial && digits.length < suffixLength) return;
        const suffix = digits.slice(-suffixLength);
        cashierGroups.set(suffix, [...(cashierGroups.get(suffix) ?? []), row]);
      });
    });
    const bankGroups = new Map<string, BankRow[]>();
    bank.forEach(row => {
      if (usedBank.has(bankKey(row)) || savedBankIds.has(bankKey(row))) return;
      jawwalPhoneDigits(row.description).forEach(({ digits, partial }) => {
        if (partial && digits.length !== suffixLength) return;
        if (!partial && digits.length < suffixLength) return;
        const suffix = digits.slice(-suffixLength);
        bankGroups.set(suffix, [...(bankGroups.get(suffix) ?? []), row]);
      });
    });
    cashierGroups.forEach((cashierRows, suffix) => {
      const bankRows = bankGroups.get(suffix);
      if (cashierRows.length !== 1 || bankRows?.length !== 1) return;
      const cashierRow = cashierRows[0];
      const bankRow = bankRows[0];
      if (!(isJawwalRow(cashierRow) || isJawwalRow(bankRow)) ||
          !jawwalPhoneSuffixMatches(`${cashierRow.rawName} ${cashierRow.name}`, bankRow.description) ||
          Math.abs(bankRow.rawAmount - cashierRow.matchAmount) > 0.01 ||
          wasPairRejected(rejectedPairs, cashierRow, bankRow)) return;
      results.push({
        type: "pending",
        bank: bankRow,
        cashier: cashierRow,
        isApprox: false,
        isExactAmount: true,
        matchType: "phone",
        amountDiff: 0,
        matchScore: 100,
        accountTypeDiff: accountTypesDiffer(bankRow.accountType, cashierRow.accountType),
      });
      usedCashier.add(cashierKey(cashierRow));
      usedBank.add(bankKey(bankRow));
    });
  }

  cashier.forEach(c => {
    if (usedCashier.has(cashierKey(c)) || savedCashierIds.has(cashierKey(c))) return;
    if (jawwalPayCashierIds.has(cashierKey(c))) {
      results.push({ type: "jawwalPay", cashier: c });
      usedCashier.add(cashierKey(c));
      return;
    }
    if (mahmoudWalletCashierIds.has(cashierKey(c))) {
      results.push({ type: "mahmoudWallet", cashier: c });
      usedCashier.add(cashierKey(c));
      return;
    }
    if (visaCashierIds.has(cashierKey(c))) {
      results.push({ type:"visa", cashier:c });
      usedCashier.add(cashierKey(c));
    }
  });

  const candidates: Array<{ cashier: CashierRow; bank: BankRow; score: number; isApprox: boolean; matchType: string }> = [];
  const bankAmountIndexes = indexByAmountAndType(bank, row => row.rawAmount, row => row.type);
  const cashierMatchProfiles = new Map(cashier.map(row => [row, {
    tokens: nameTokens(row.name),
    normalized: normName(row.name),
    aliasKey: smartNameKey(row.name),
  }]));
  const bankMatchProfiles = new Map(bank.map(row => [row, {
    normalized: normName(row.description),
    segments: row.description.split("/").map(segment => segment.trim()).filter(Boolean).map(segment => ({
      tokens: nameTokens(segment),
      aliasKey: smartNameKey(segment),
    })),
  }]));
  cashier.forEach(c => {
    if (usedCashier.has(cashierKey(c)) || savedCashierIds.has(cashierKey(c)) || visaCashierIds.has(cashierKey(c)) || jawwalPayCashierIds.has(cashierKey(c)) || mahmoudWalletCashierIds.has(cashierKey(c))) return;
    const cashierProfile = cashierMatchProfiles.get(c)!;
    const amountTolerance = Math.max(0.01, c.matchAmount * amountTolerancePercent / 100);
    const amountCandidates = amountRange(bankAmountIndexes.get(c.type) ?? [], c.matchAmount, amountTolerance);
    if (performanceMetrics) performanceMetrics.amountCandidates += amountCandidates.length;
    amountCandidates.forEach(({ row: b }) => {
      if (usedBank.has(bankKey(b)) || savedBankIds.has(bankKey(b))) return;
      if (isJawwalRow(c) || isJawwalRow(b)) return;
      if (wasPairRejected(rejectedPairs, c, b) || savedKeys.has(`${cashierKey(c)}-${bankKey(b)}`)) return;
      if (Math.abs(b.rawAmount - c.matchAmount) > amountTolerance) return;
      if (performanceMetrics) performanceMetrics.nameChecks += 1;
      const bankProfile = bankMatchProfiles.get(b)!;
      const aliasTarget = aliases[cashierProfile.aliasKey];
      const aliasMatch = aliasTarget && bankProfile.segments.some(segment => segment.aliasKey === aliasTarget);
      const segmentMatches = aliasMatch ? [] : bankProfile.segments
        .map(segment => advancedMatchCheckTokens(cashierProfile.tokens, segment.tokens))
        .filter(result => result.isMatch)
        .sort((a, b) => Number(a.isApprox) - Number(b.isApprox));
      const ms = aliasMatch
        ? { isMatch: true, isApprox: false, matchType: "exact" as const }
        : segmentMatches[0] ?? { isMatch: false, isApprox: false, matchType: "none" as const };
      const similarity = nameSimilarityNormalized(cashierProfile.normalized, bankProfile.normalized);
      if (!ms.isMatch && similarity < 0.25) return;
      const priority = ms.isMatch ? (ms.matchType === "exact" ? 4 : ms.matchType === "firstSecond" ? 3 : ms.matchType === "fourthName" ? 2 : 1) : 0;
      candidates.push({ cashier: c, bank: b, score: priority + similarity, isApprox: !ms.isMatch || ms.isApprox, matchType: ms.isMatch ? ms.matchType : "closest-name" });
    });
  });
  candidates.sort((a, b) => b.score - a.score);
  candidates.forEach(candidate => {
    if (usedCashier.has(cashierKey(candidate.cashier)) || usedBank.has(bankKey(candidate.bank))) return;
    results.push({
      type: "pending",
      bank: candidate.bank,
      cashier: candidate.cashier,
      isApprox: candidate.isApprox,
      isExactAmount: true,
      matchType: candidate.matchType,
      amountDiff: Math.abs(candidate.bank.rawAmount - candidate.cashier.matchAmount),
      matchScore: Math.round(Math.min(1, candidate.score / 4.5) * 100),
      accountTypeDiff: accountTypesDiffer(candidate.bank.accountType, candidate.cashier.accountType)
    });
    usedBank.add(bankKey(candidate.bank));
    usedCashier.add(cashierKey(candidate.cashier));
  });

  const remC = cashier.filter(c => !isJawwalRow(c) && !usedCashier.has(cashierKey(c)) && !savedCashierIds.has(cashierKey(c)) && !visaCashierIds.has(cashierKey(c)) && !jawwalPayCashierIds.has(cashierKey(c)) && !mahmoudWalletCashierIds.has(cashierKey(c)));
  const remB = bank.filter(b => !isJawwalRow(b) && !usedBank.has(bankKey(b)) && !savedBankIds.has(bankKey(b)));

  const getBase = (name:string) => normName(name).split(" ").filter(t=>t.length>2&&t!=="al").slice(0,2).join(" ");
  const groups: Record<string,{cashiers:CashierRow[];banks:BankRow[]}> = {};
  const groupsByFirstToken = new Map<string, string[]>();
  remC.forEach(c => {
    const k = getBase(c.name);
    if (!k) return;
    if (!groups[k]) {
      groups[k] = {cashiers:[], banks:[]};
      const firstToken = k.split(" ")[0];
      groupsByFirstToken.set(firstToken, [...(groupsByFirstToken.get(firstToken) ?? []), k]);
    }
    groups[k].cashiers.push(c);
  });
  const groupOrder = new Map(Object.keys(groups).map((key, index) => [key, index]));
  const groupKeysByTrigram = new Map<string, string[]>();
  groupsByFirstToken.forEach((keys, firstToken) => {
    const grams = new Set<string>();
    for (let start = 0; start <= firstToken.length - 3; start++) grams.add(firstToken.slice(start, start + 3));
    grams.forEach(gram => {
      const indexedKeys = groupKeysByTrigram.get(gram) ?? [];
      keys.forEach(key => indexedKeys.push(key));
      groupKeysByTrigram.set(gram, indexedKeys);
    });
  });
  remB.forEach(b => {
    const normalizedBankName = normName(b.description);
    const candidateGroups = new Set<string>();
    const words = normalizedBankName.split(" ");
    words.forEach(word => {
      const grams = new Set<string>();
      for (let start = 0; start <= word.length - 3; start++) grams.add(word.slice(start, start + 3));
      grams.forEach(gram => groupKeysByTrigram.get(gram)?.forEach(key => candidateGroups.add(key)));
    });
    const orderedCandidates = [...candidateGroups].sort((a, b) => groupOrder.get(a)! - groupOrder.get(b)!);
    for (const k of orderedCandidates) {
      const kT = k.split(" ");
      if (kT.every(kt => normalizedBankName.includes(kt))) {
        groups[k].banks.push(b);
        break;
      }
    }
  });

  Object.values(groups).forEach(g => {
    if (!g.cashiers.length || !g.banks.length) return;

    const types = new Set(g.cashiers.map(c => c.type));
    if (types.size > 1) return;

    const totC = g.cashiers.reduce((s,c) => s + c.matchAmount, 0);
    const totB = g.banks.reduce((s,b) => s + b.rawAmount, 0);
    if (Math.abs(totC - totB) > 0.05) return;

    const pairCount = Math.min(g.cashiers.length, g.banks.length);
    for (let i = 0; i < pairCount; i++) {
      const c = g.cashiers[i];
      const b = g.banks[i];
      if (usedCashier.has(cashierKey(c)) || usedBank.has(bankKey(b))) continue;
      if (!wasPairRejected(rejectedPairs, c, b) && !savedKeys.has(`${cashierKey(c)}-${bankKey(b)}`)) {
        results.push({
          type: "pending",
          bank: b,
          cashier: c,
          isApprox: false,
          isBundled: true,
          isExactAmount: true,
          matchType: "bundle",
          amountDiff: 0
        });
        usedCashier.add(cashierKey(c));
        usedBank.add(bankKey(b));
      }
    }
  });

  const remainingExactBankIndex = indexByAmountAndType(
    bank.filter(b => !isJawwalRow(b) && !usedBank.has(bankKey(b)) && !savedBankIds.has(bankKey(b))),
    row => row.rawAmount,
    row => row.type,
  );
  const remainingExactCashierIndex = indexByAmountAndType(
    cashier.filter(c => !isJawwalRow(c) && !usedCashier.has(cashierKey(c)) && !savedCashierIds.has(cashierKey(c)) && !visaCashierIds.has(cashierKey(c)) && !jawwalPayCashierIds.has(cashierKey(c)) && !mahmoudWalletCashierIds.has(cashierKey(c))),
    row => row.matchAmount,
    row => row.type,
  );

  cashier.forEach(c => {
    if (usedCashier.has(cashierKey(c)) || savedCashierIds.has(cashierKey(c)) || visaCashierIds.has(cashierKey(c)) || jawwalPayCashierIds.has(cashierKey(c)) || mahmoudWalletCashierIds.has(cashierKey(c))) return;
    const hasAmt = hasAmountWithin(remainingExactBankIndex.get(c.type) ?? [], c.matchAmount, 0.01);
    results.push({
      type:"unmatchedCashier",
      cashier:c,
      reason: hasAmt ? "اختلاف في الاسم" : "غير موجودة في البنك"
    });
  });

  bank.forEach(b => {
    if (usedBank.has(bankKey(b)) || savedBankIds.has(bankKey(b))) return;
    const hasAmt = hasAmountWithin(remainingExactCashierIndex.get(b.type) ?? [], b.rawAmount, 0.01);
    results.push({
      type:"unmatchedBank",
      bank:b,
      reason: hasAmt ? "اختلاف في الاسم" : "الحوالة غير موجودة في الكاشير"
    });
  });

  return results;
}

// ─── MatchResult type ────────────────────────────────────────────────────────
type MatchResult =
  | { type:"saved"; bank:BankRow; cashier:CashierRow; savedMatch:SavedMatch }
  | { type:"pending"; bank:BankRow; cashier:CashierRow; isApprox:boolean; isBundled?:boolean; isExactAmount:boolean; matchType?:string; amountDiff?:number; accountTypeDiff?:boolean; matchScore?:number }
  | { type:"manualGroup"; group:ManualMatchGroup }
  | { type:"visa"; cashier:CashierRow }
  | { type:"jawwalPay"; cashier:CashierRow }
  | { type:"mahmoudWallet"; cashier:CashierRow }
  | { type:"unmatchedCashier"; cashier:CashierRow; reason:"اختلاف في الاسم"|"غير موجودة في البنك" }
  | { type:"unmatchedBank"; bank:BankRow; reason:"اختلاف في الاسم"|"الحوالة غير موجودة في الكاشير" };

type TabId = "saved"|"pending"|"manual"|"visa"|"jawwalPay"|"mahmoudWallet"|"held"|"review"|"nonReconciliationBank"|"uCashier"|"uBank";
type PageId = "upload" | "stageBChoice" | "main" | "manual" | "assist" | "report" | "recon2" | "recon2StageA" | "recon2Visa" | "splitTool" | "batchCutoff" | "clearing";

interface ReferenceLocation {
  id: string;
  label: string;
  detail: string;
  page: PageId;
  tab?: TabId;
}

interface MatchUndoSnapshot {
  manualGroups: ManualMatchGroup[];
  savedMatches: SavedMatch[];
  rejectedPairs: Set<string>;
  nameAliases: Record<string, string>;
  selectedPendingKeys: Set<string>;
}

type UndoAction =
  | { type: "match"; snapshot: MatchUndoSnapshot }
  | { type: "custom"; label: string; undo: () => void };

function scoreCandidate(
  cashierRow: CashierRow,
  b: BankRow,
  claimedNames: Set<string>,
  aliases: Record<string, string> = {}
): { score:number; matchType:string; reason:string } | null {
  const aliasMatch = aliases[smartNameKey(cashierRow.name)] && b.description.split("/").some(segment => smartNameKey(segment) === aliases[smartNameKey(cashierRow.name)]);
  const ms = aliasMatch ? { isMatch: true, isApprox: false, matchType: "exact" } : advancedMatchCheck(cashierRow.name, b.description);
  const diff = Math.abs(b.rawAmount - cashierRow.matchAmount);
  const amtMatch = diff <= 0.01;

  if (ms.isMatch && amtMatch) {
    return { score: 0.95, matchType: ms.matchType, reason: "الاسم والمبلغ متطابقان" };
  }
  if (ms.isMatch && !amtMatch) {
    const score = 0.5 + nameSim(cashierRow.name, b.description) * 0.4;
    return { score, matchType: ms.matchType, reason: "الاسم متطابق، المبلغ مختلف" };
  }
  if (!ms.isMatch && amtMatch) {
    const bankNameKey = normName(b.description);
    if (bankNameKey && claimedNames.has(bankNameKey)) return null;
    return { score: 0.2, matchType: "amount_only", reason: "المبلغ متطابق فقط (الاسم مختلف)" };
  }
  return null;
}

// ─── Suggestions ──────────────────────────────────────────────────────────────
function getSuggestions(
  cashierRow: CashierRow,
  availableBank: BankRow[],
  rejectedPairs: Set<string>,
  savedKeys: Set<string>,
  claimedNames: Set<string>,
  ownerMap: Map<string, {cashierKey:string; score:number}>,
  aliases: Record<string, string> = {}
): Array<{bank:BankRow; score:number; matchType:string; amountDiff:number; reason:string}> {
  return availableBank
    .filter(b => {
      const pairKey = transactionPairKey(cashierRow, b);
      if (cashierRow.type !== b.type) return false;
      if (wasPairRejected(rejectedPairs, cashierRow, b) || savedKeys.has(pairKey)) return false;
      if (isJawwalAccount(cashierRow.accountType) || isJawwalAccount(b.accountType)) {
        const cashierText = `${cashierRow.rawName} ${cashierRow.name}`;
        const sameSuffixRows = availableBank.filter(candidate =>
          (isJawwalAccount(candidate.accountType) || isJawwalAccount(cashierRow.accountType)) &&
          jawwalPhoneSuffixMatches(cashierText, candidate.description)
        );
        return sameSuffixRows.length === 1 && sameSuffixRows[0] === b &&
          jawwalPhoneSuffixMatches(cashierText, b.description) &&
          Math.abs(b.rawAmount - cashierRow.matchAmount) <= 0.01;
      }
      return true;
    })
    .map(b => {
      if (isJawwalAccount(cashierRow.accountType) || isJawwalAccount(b.accountType)) {
        return {
          bank: b, score: 1, matchType: "phone",
          amountDiff: Math.abs(b.rawAmount - cashierRow.matchAmount),
          reason: "رقم الجوال والمبلغ متطابقان",
        };
      }
      const sc = scoreCandidate(cashierRow, b, claimedNames, aliases);
      if (!sc) return null;
      const owner = ownerMap.get(transactionKey(b));
      if (owner && owner.cashierKey !== transactionKey(cashierRow)) return null;
      const diff = Math.abs(b.rawAmount - cashierRow.matchAmount);
      return { bank: b, score: sc.score, matchType: sc.matchType, amountDiff: diff, reason: sc.reason };
    })
    .filter((x): x is {bank:BankRow; score:number; matchType:string; amountDiff:number; reason:string} => x !== null)
    .sort((a, b) => b.score - a.score)
    .slice(0, 5);
}

// ─── Export ───────────────────────────────────────────────────────────────────
function sanitizeSheetName(name: string, used: Set<string>): string {
  const clean = name.replace(/[\\/?*[\]:]/g, "-").slice(0, 31).trim() || "شيت";
  let final = clean;
  let i = 2;
  while (used.has(final)) {
    const suffix = ` ${i}`;
    final = clean.slice(0, 31 - suffix.length) + suffix;
    i++;
  }
  used.add(final);
  return final;
}

function matchChannelLabel(bank: BankRow, cashier: CashierRow): string {
  const bankChannel = bank.accountType || "غير محدد";
  const cashierChannel = cashier.accountType || "غير محدد";
  return bankChannel === cashierChannel ? bankChannel : `${cashierChannel} ← ${bankChannel}`;
}

function matchedChannelRows(results: MatchResult[]) {
  return results
    .filter((row): row is Extract<MatchResult, { type: "saved" }> => row.type === "saved")
    .map(row => ({
      ...row,
      channel: matchChannelLabel(row.bank, row.cashier),
      bankNet: row.bank.debit - row.bank.credit,
      cashierNet: row.cashier.debit - row.cashier.credit,
    }));
}

function financialIssueRows(results: MatchResult[]) {
  return matchedChannelRows(results)
    .map(row => ({
      ...row,
      hasAmountDifference: Math.abs(row.bank.rawAmount - row.cashier.matchAmount) > 0.01,
      hasAccountTypeDifference: accountTypesDiffer(row.bank.accountType, row.cashier.accountType),
      hasDateDifference: transactionDatesDiffer(row.bank.date, row.cashier.date),
      manualEditIssue: manualEditIssue(row.savedMatch.editorNotes),
    }))
    .filter(row => row.hasAmountDifference || row.hasAccountTypeDifference || row.hasDateDifference || !!row.manualEditIssue);
}

function doExport(
  results: MatchResult[],
  savedMatches: SavedMatch[],
  visaItems: VisaItem[],
  heldItems: HeldItem[],
  bankRows: BankRow[] = [],
  cashierRows: CashierRow[] = [],
  resumeData?: ResumeData,
  nonReconciliationBankRows: BankRow[] = [],
  bankSplits: BankSplit[] = [],
  cutoffBatches: CutoffBatch[] = [],
  clearingGroups: ClearingGroup[] = [],
  sourceCashierRows: CashierRow[] = [],
  sourceBankRows: BankRow[] = [],
  rowNotes: Record<string, string> = {},
  stageAInvoicesParam: StageAResult[] = [],
  stageADatabaseSummaryParam?: StageASummary,
) {
  const stageADatabaseSummary: StageASummary = stageADatabaseSummaryParam ?? { databaseInvoiceCount: 0, databaseInvoiceTotal: 0, databaseCashTotal: 0, databaseDebtTotal: 0, databaseCardTotal: 0 };
  const wb = XLSX.utils.book_new();
  const usedSheetNames = new Set<string>();
  const cleanReportNote = (value: unknown) => String(value ?? "")
    .replace(/(?:الرقم بالفاتورة|الرقم في الفاتورة)[^.!؛\n]*(?:مختلف|$)[^.!؛\n]*/g, "اختلاف رقم الجوال")
    .replace(/(?:تعديل الرقم إلى|الرقم الموجود في الحوالة)[^.!؛\n]*/g, "")
    .replace(/غير موجود/g, "—")
    .replace(/\s+/g, " ")
    .trim();
  const add=(name:string,hdrs:string[],rows:unknown[][])=>{
    const ws=XLSX.utils.aoa_to_sheet([hdrs,...rows]);
    const widths = hdrs.map((header, columnIndex) => {
      const longestValue = rows.reduce((longest, row) => {
        const value = row[columnIndex];
        return Math.max(longest, value == null ? 0 : String(value).length);
      }, 0);
      return { wch: Math.min(46, Math.max(12, String(header).length + 3, longestValue + 2)) };
    });
    ws["!cols"] = widths;
    ws["!autofilter"] = { ref: ws["!ref"] || `A1:${XLSX.utils.encode_col(Math.max(0, hdrs.length - 1))}1` };
    ws["!rows"] = [{ hpt: 26 }];
    if (ws["!ref"]) {
      const range = XLSX.utils.decode_range(ws["!ref"]);
      for (let rowIndex = 1; rowIndex <= range.e.r; rowIndex += 1) {
        for (let columnIndex = 0; columnIndex <= range.e.c; columnIndex += 1) {
          const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
          const cell = ws[cellAddress];
          const header = hdrs[columnIndex] || "";
          if (cell && typeof cell.v === "number" &&
              /مبلغ|الإجمالي|الفرق|مدين|دائن|صافي|قيمة|المفروض|الواصل|حصة/.test(header) &&
              !/عدد|معرّف|معرف|رقم/.test(header)) {
            cell.z = '#,##0.00;[Red]-#,##0.00;–';
          }
        }
      }
    }
    XLSX.utils.book_append_sheet(wb,ws,sanitizeSheetName(name, usedSheetNames));
  };

  const bankTotal = bankRows.reduce((sum, row) => sum + row.rawAmount, 0);
  const cashierTotal = cashierRows.reduce((sum, row) => sum + row.matchAmount, 0);
  const matchedResults = results.filter((row): row is Extract<MatchResult, { type: "saved" }> => row.type === "saved");
  const matchedBankTotal = matchedResults.reduce((sum, row) => sum + row.bank.rawAmount, 0);
  const matchedCashierTotal = matchedResults.reduce((sum, row) => sum + row.cashier.matchAmount, 0);
  const categorySummary = [
    ["إجمالي الكشوف", bankTotal, cashierTotal, bankTotal - cashierTotal, bankRows.length, cashierRows.length],
    ["المطابقات المؤكدة", matchedBankTotal, matchedCashierTotal, matchedBankTotal - matchedCashierTotal, matchedResults.length, matchedResults.length],
    ["فيزا", 0, results.filter(row => row.type === "visa").reduce((sum, row) => sum + row.cashier.matchAmount, 0), 0, 0, results.filter(row => row.type === "visa").length],
    ["محفظة شخصية", 0, results.filter(row => row.type === "mahmoudWallet").reduce((sum, row) => sum + row.cashier.matchAmount, 0), 0, 0, results.filter(row => row.type === "mahmoudWallet").length],
    ["جوال بي", 0, results.filter(row => row.type === "jawwalPay").reduce((sum, row) => sum + row.cashier.matchAmount, 0), 0, 0, results.filter(row => row.type === "jawwalPay").length],
    ["فواتير بلا حوالات", 0, results.filter(row => row.type === "unmatchedCashier").reduce((sum, row) => sum + row.cashier.matchAmount, 0), 0, 0, results.filter(row => row.type === "unmatchedCashier").length],
    ["حوالات بلا فواتير", results.filter(row => row.type === "unmatchedBank").reduce((sum, row) => sum + row.bank.rawAmount, 0), 0, 0, results.filter(row => row.type === "unmatchedBank").length, 0],
    ["معلقات بنك", heldItems.filter(item => item.kind === "bank").reduce((sum, item) => sum + (item.data as BankRow).rawAmount, 0), 0, 0, heldItems.filter(item => item.kind === "bank").length, 0],
    ["معلقات كاشير", 0, heldItems.filter(item => item.kind === "cashier").reduce((sum, item) => sum + (item.data as CashierRow).matchAmount, 0), 0, 0, heldItems.filter(item => item.kind === "cashier").length],
  ];
  categorySummary.push(
    ["قاعدة البيانات · إجمالي الفواتير", 0, stageADatabaseSummary.databaseInvoiceTotal, 0, 0, stageADatabaseSummary.databaseInvoiceCount],
    ["قاعدة البيانات · مدفوع نقداً", 0, stageADatabaseSummary.databaseCashTotal, 0, 0, stageADatabaseSummary.databaseInvoiceCount],
    ["قاعدة البيانات · المتبقي على الزبائن", 0, stageADatabaseSummary.databaseDebtTotal, 0, 0, stageADatabaseSummary.databaseInvoiceCount],
    ["قاعدة البيانات · المدفوع بالتطبيقات/البطاقات", 0, stageADatabaseSummary.databaseCardTotal, 0, 0, stageADatabaseSummary.databaseInvoiceCount],
    ["قاعدة البيانات · الخصومات", 0, stageADatabaseSummary.databaseDiscountTotal ?? 0, 0, 0, (stageADatabaseSummary.databaseInvoices ?? []).filter(invoice => invoice.discountAmount > 0).length],
    ["قاعدة البيانات · المطابق / غير المطابق", 0, 0, 0, stageAInvoicesParam.filter(invoice => invoice.matched).length, stageAInvoicesParam.filter(invoice => !invoice.matched).length],
  );
  add("ملخص التسوية",
    ["البند","من البنك","من الكاشير","الفرق (بنك - كاشير)","عدد البنك","عدد الكاشير"],
    categorySummary.map(row => row.map((value, index) => typeof value === "number" && index > 0 && index < 4 ? fmtNum(value) : value)));
  const unmatchedFinancialRows = [
    ["فواتير بلا حوالات · مدين", results.filter(r => r.type === "unmatchedCashier" && r.cashier.debit > 0).length, results.filter(r => r.type === "unmatchedCashier").reduce((sum, r) => sum + r.cashier.debit, 0)],
    ["فواتير بلا حوالات · دائن", results.filter(r => r.type === "unmatchedCashier" && r.cashier.credit > 0).length, results.filter(r => r.type === "unmatchedCashier").reduce((sum, r) => sum + r.cashier.credit, 0)],
    ["حوالات بلا فواتير · مدين", results.filter(r => r.type === "unmatchedBank" && r.bank.debit > 0).length, results.filter(r => r.type === "unmatchedBank").reduce((sum, r) => sum + r.bank.debit, 0)],
    ["حوالات بلا فواتير · دائن", results.filter(r => r.type === "unmatchedBank" && r.bank.credit > 0).length, results.filter(r => r.type === "unmatchedBank").reduce((sum, r) => sum + r.bank.credit, 0)],
  ];
  add("غير المطابقات مدين ودائن", ["الفئة","العدد","الإجمالي بدون طرح الدائن"], unmatchedFinancialRows);
  add("فروقات فواتير المرحلة A",
    ["المنصة", "حالة مطابقة قاعدة البيانات", "بيان الفاتورة", "رقم الفاتورة", "رقم المستخدم", "رقم الزبون", "التاريخ", "الوقت", "نوع الحركة", "مبلغ الحركة", "إجمالي الفاتورة", "نقدي", "شيك", "مدفوع بالكارد / تطبيق", "المتبقي", "الخصم", "ملاحظة"],
    stageAInvoicesParam.map(invoice => [
      invoice.category || invoice.source, invoice.matched ? "مطابق" : "غير مطابق", invoice.originalName, invoice.invoiceNumber || "—",
      invoice.userId || "—", invoice.customerNumber || "—", invoice.invoiceDate || "—", invoice.registrationTime || "—", invoice.type,
      invoice.amount, invoice.invoiceTotal, invoice.cashAmount, invoice.checkAmount, invoice.databaseAmount, invoice.debtAmount, invoice.discountAmount,
      invoice.issue || invoice.categoryIssue || "—",
    ]));
  if (stageADatabaseSummary.databaseInvoiceCount > 0) {
    add("ملخص فواتير قاعدة البيانات",
      ["البند", "القيمة"],
      [
        ["عدد الفواتير", stageADatabaseSummary.databaseInvoiceCount],
        ["إجمالي الفواتير", stageADatabaseSummary.databaseInvoiceTotal],
        ["المبلغ النقدي", stageADatabaseSummary.databaseCashTotal],
        ["المدفوع بالشيك", (stageADatabaseSummary.databaseInvoices ?? []).reduce((sum, invoice) => sum + invoice.checkAmount, 0)],
        ["مجموع الخصومات", stageADatabaseSummary.databaseDiscountTotal ?? 0],
        ["المتبقي على الزبائن", stageADatabaseSummary.databaseDebtTotal],
        ["المدفوع بالتطبيقات/البطاقات", stageADatabaseSummary.databaseCardTotal],
      ]);
    add("فواتير قاعدة البيانات المرحلة",
      ["المنصة", "البيان", "اسم قاعدة البيانات", "رقم الفاتورة", "نوع الحركة", "مبلغ الحركة", "مبلغ التطبيق", "إجمالي الفاتورة", "نقدي", "شيك", "دين", "الخصم", "رقم المستخدم", "التاريخ", "الساعة", "حالة مطابقة قاعدة البيانات"],
      stageAInvoicesParam.map(invoice => [
        invoice.source, invoice.originalName, invoice.databaseName || "—", invoice.invoiceNumber || "—",
        invoice.type, invoice.amount, invoice.databaseAmount, invoice.invoiceTotal, invoice.cashAmount, invoice.checkAmount, invoice.debtAmount, invoice.discountAmount,
        invoice.userId || "—", invoice.invoiceDate || "—", invoice.registrationTime || "—", invoice.matched ? "مطابقة قاعدة البيانات" : "فاتورة مرحّلة دون سجل قاعدة بيانات",
      ]));
    add("فواتير مرحّلة بلا سجل قاعدة البيانات",
      ["المنصة", "البيان", "نوع الحركة", "مبلغ الحركة", "تاريخ الفاتورة", "الوقت"],
      stageAInvoicesParam.filter(invoice => !invoice.matched).map(invoice => [
        invoice.source, invoice.originalName, invoice.type, invoice.amount, invoice.invoiceDate || "—", invoice.registrationTime || "—",
      ]));
    const databaseInvoices = stageADatabaseSummary.databaseInvoices ?? [];
    add("كل فواتير قاعدة البيانات",
      ["رقم الفاتورة", "التاريخ", "الوقت", "اسم الزبون", "رقم الزبون", "رقم المستخدم", "إجمالي الفاتورة", "نقدي", "شيك", "تطبيق/بطاقة", "دين", "الخصم"],
      databaseInvoices.map(invoice => [invoice.invoiceNumber, invoice.invoiceDate, invoice.registrationTime, invoice.name, invoice.customerNumber, invoice.userId, invoice.invoiceTotal, invoice.cashAmount, invoice.checkAmount, invoice.cardAmount, invoice.debtAmount, invoice.discountAmount]));
    add("فواتير فيها خصم",
      ["رقم الفاتورة", "التاريخ", "الوقت", "اسم الزبون", "رقم الزبون", "رقم المستخدم", "إجمالي الفاتورة", "نقدي", "شيك", "تطبيق/بطاقة", "دين", "الخصم"],
      databaseInvoices.filter(invoice => invoice.discountAmount > 0).map(invoice => [invoice.invoiceNumber, invoice.invoiceDate, invoice.registrationTime, invoice.name, invoice.customerNumber, invoice.userId, invoice.invoiceTotal, invoice.cashAmount, invoice.checkAmount, invoice.cardAmount, invoice.debtAmount, invoice.discountAmount]));
  }
  add("معلقات البنك غير التابعة للمطابقة",
    ["التاريخ","البيان","النوع","مدين","دائن","المبلغ","نوع الحساب"],
    nonReconciliationBankRows.map(row => [row.date, row.description, row.type, row.debit, row.credit, row.rawAmount, row.accountType || "—"]));
  const sourceCashierByKey = new Map(sourceCashierRows.map(row => [transactionKey(row), row]));
  const sourceBankByKey = new Map(sourceBankRows.map(row => [transactionKey(row), row]));
  const splitRows = bankSplits.flatMap(split => split.allocations.map((allocation, index) => {
    const bank = sourceBankByKey.get(`${split.bankSessionId}:${split.bankId}`);
    const cashier = allocation.kind === "company" && allocation.cashierId != null
      ? sourceCashierByKey.get(`${allocation.cashierSessionId ?? 0}:${allocation.cashierId}`)
      : undefined;
    return [
      split.date, split.description, split.originalAmount, bank?.ref ?? "", bank?.accountType ?? "",
      index + 1, allocation.kind === "company" ? "حصة شركتكم" : "حصة خارجية",
      allocation.amount, allocation.note, cashier?.date ?? "", cashier?.name ?? "",
      cashier?.ref ?? "", cashier?.matchAmount ?? "", allocation.cashierId ?? "",
      allocation.cashierSessionId ?? "", split.savedAt,
    ];
  }));
  add("تجزئة الحوالات", ["تاريخ الحوالة","بيان الحوالة الأصلية","قيمة الأصل","مرجع الحوالة","حساب الحوالة","رقم الحصة","التخصيص","مبلغ الحصة","ملاحظة التخصيص","تاريخ الفاتورة","اسم الفاتورة أو السند","مرجع الفاتورة","مبلغ الفاتورة","معرّف الفاتورة","جلسة الفاتورة","تاريخ التجزئة"], splitRows);
  const cutoffRows = cutoffBatches.flatMap(batch => batch.sourceRows.map((source, index) => {
    const target = sourceCashierByKey.get(`${batch.aggregateSessionId}:${batch.aggregateId}`);
    const detail = sourceCashierByKey.get(`${source.sessionId}:${source.id}`);
    return [
      batch.dateFrom, batch.dateTo, batch.graceDays, target?.date ?? "", target?.name ?? "",
      target?.ref ?? "", target?.matchAmount ?? "", batch.aggregateId, batch.aggregateSessionId,
      index + 1, detail?.date ?? "", detail?.name ?? "", detail?.ref ?? "",
      detail?.type ?? "", detail?.matchAmount ?? "", source.id, source.sessionId,
      batch.total, batch.note, batch.savedAt,
    ];
  }));
  add("تجميع فترة التعطل", ["من تاريخ","إلى تاريخ","نافذة التداخل بالأيام","تاريخ سند الإجمالي","اسم سند الإجمالي","مرجع سند الإجمالي","قيمة سند الإجمالي","معرّف سند الإجمالي","جلسة سند الإجمالي","رقم الفاتورة","تاريخ الفاتورة الفعلية","اسم الفاتورة الفعلية","مرجع الفاتورة","الاتجاه","قيمة الفاتورة","معرّف الفاتورة الفعلية","جلسة الفاتورة","إجمالي المجموعة","ملاحظة تدقيقية","تاريخ الحفظ"], cutoffRows);
  const clearingRows = clearingGroups.flatMap(group => [
    ...group.debitRows.map((reference, index) => {
      const row = sourceCashierByKey.get(`${reference.sessionId}:${reference.id}`);
      return ["مدين", group.id, row?.date ?? "", row?.name ?? "", row?.ref ?? "", row?.debit ?? "", reference.id, reference.sessionId, index + 1, group.total, group.note, group.savedAt];
    }),
    ...group.creditRows.map((reference, index) => {
      const row = sourceCashierByKey.get(`${reference.sessionId}:${reference.id}`);
      return ["دائن", group.id, row?.date ?? "", row?.name ?? "", row?.ref ?? "", row?.credit ?? "", reference.id, reference.sessionId, index + 1, group.total, group.note, group.savedAt];
    }),
  ]);
  add("تسكير السندات", ["الطرف","معرّف مجموعة التسكير","التاريخ","بيان الحركة","المرجع","مبلغ الحركة","معرّف الحركة","جلسة الحركة","رقم الحركة في الطرف","المبلغ المتساوي","ملاحظة تدقيقية","تاريخ الحفظ"], clearingRows);

  const savedRows = results.filter(r=>r.type==="saved").map((r:any) => {
    const sm = r.savedMatch as SavedMatch;
    return [
      r.cashier.name, r.cashier.userId || "—", r.bank.description,
      sm.type, r.bank.rawAmount, r.cashier.matchAmount, r.cashier.discountAmount ?? 0,
      r.cashier.date, r.cashier.registrationTime || "—",
      sm.isAmountDiff ? "نعم" : "لا", sm.isPhoneDiff ? "نعم" : "لا", sm.isNameDiff ? "نعم" : "لا",
      r.bank.accountType || "—", r.cashier.accountType || "—",
      sm.isAccountTypeDiff ? "نعم" : "لا", transactionDatesDiffer(r.bank.date, r.cashier.date) ? "نعم" : "لا", sm.isManual ? "يدوي" : "تلقائي",
      cleanReportNote(sm.note), sm.matchScore ?? "", cleanReportNote(sm.editorNotes),
      r.bank.date, r.cashier.date, sm.savedAt
    ];
  });
  add("✅ المطابقات (المؤكدة)",
    ["الفاتورة","رقم المستخدم","بيان الحوالة","النوع","مبلغ البنك","مبلغ الفاتورة","الخصم","تاريخ الفاتورة","وقت الفاتورة","اختلاف مبلغ","اختلاف جوال","اختلاف اسم","نوع حساب البنك","نوع حساب الكاشير","اختلاف نوع حساب","اختلاف التاريخ","نوع المطابقة","ملاحظة","درجة التطابق","ملاحظات المحرر","ت.البنك","ت.الكاشير","تاريخ الحفظ"],
    savedRows);

  const matchedDetails = matchedChannelRows(results);
  add("تفصيل المطابقات والقنوات",
    ["القناة في المطابقة","فاتورة الكاشير","رقم المستخدم","تاريخ الفاتورة","وقت الفاتورة","خصم الفاتورة","الحوالة","كاشير مدين","كاشير دائن","كاشير صافي","بنك مدين","بنك دائن","بنك صافي","فرق المطابقة","نوع المطابقة","تاريخ الحفظ"],
    matchedDetails.map(row => [
      row.channel, row.cashier.name, row.cashier.userId || "—", row.cashier.date, row.cashier.registrationTime || "—", row.cashier.discountAmount ?? 0, row.bank.description,
      row.cashier.debit, row.cashier.credit, row.cashierNet,
      row.bank.debit, row.bank.credit, row.bankNet,
      row.bankNet - row.cashierNet, row.savedMatch.isManual ? "يدوي" : "تلقائي", row.savedMatch.savedAt,
    ]));
  const matchedChannelTotals = Array.from(new Set(matchedDetails.map(row => row.channel))).map(channel => {
    const rows = matchedDetails.filter(row => row.channel === channel);
    const sum = (pick: (row: typeof rows[number]) => number) => rows.reduce((total, row) => total + pick(row), 0);
    return [channel, rows.length, sum(row => row.bank.debit), sum(row => row.bank.credit), sum(row => row.bankNet), sum(row => row.cashier.debit), sum(row => row.cashier.credit), sum(row => row.cashierNet), sum(row => row.bankNet - row.cashierNet)];
  });
  add("إجماليات المطابقة حسب القناة",
    ["القناة","عدد المطابقات","البنك مدين","البنك دائن","البنك صافي","الفواتير مدين","الفواتير دائن","الفواتير صافي","فرق القناة"],
    matchedChannelTotals);
  const financialIssues = financialIssueRows(results);
  const amountFinancialIssues = financialIssues.filter(row => row.hasAmountDifference);
  add("الفواتير ذات المشاكل المالية",
    ["القناة","الفاتورة","رقم المستخدم","الحوالة","مبلغ الفاتورة","الواصل الفعلي بالبنك","فرق الفاتورة عن البنك","نوع المشكلة","الحالة","الفاتورة مدين","الفاتورة دائن","البنك مدين","البنك دائن","تاريخ البنك","تاريخ الفاتورة"],
    financialIssues.map(row => [
      row.channel, row.cashier.name, row.cashier.userId || "—", row.bank.description,
      row.cashier.matchAmount, row.bank.rawAmount,
      row.cashier.matchAmount - row.bank.rawAmount,
      cleanReportNote(row.manualEditIssue) || [
        row.hasAmountDifference ? "اختلاف مبلغ" : "",
        row.hasAccountTypeDifference ? "اختلاف نوع الحساب" : "",
        row.hasDateDifference ? "اختلاف التاريخ" : "",
      ].filter(Boolean).join(" + "),
      row.manualEditIssue || row.hasDateDifference || row.hasAccountTypeDifference ? "تحتاج مراجعة" : "البنك هو الصحيح", row.cashier.debit, row.cashier.credit,
      row.bank.debit, row.bank.credit, row.bank.date, row.cashier.date,
    ]));
  const financialExpected = amountFinancialIssues.reduce((sum, row) => sum + row.cashier.matchAmount, 0);
  const financialActual = amountFinancialIssues.reduce((sum, row) => sum + row.bank.rawAmount, 0);
  add("إجمالي المشاكل المالية",
    ["عدد الفواتير ذات المشاكل","عدد فروقات المبلغ","عدد فروقات نوع الحساب","عدد فروقات التاريخ","إجمالي المفروض حسب فواتير فرق المبلغ","إجمالي الواصل الفعلي المرتبط بها","فرق المبلغ"],
    [[financialIssues.length, amountFinancialIssues.length, financialIssues.filter(row => row.hasAccountTypeDifference).length, financialIssues.filter(row => row.hasDateDifference).length, financialExpected, financialActual, financialExpected - financialActual]]);

  const needsReview = results.filter(r=>r.type==="saved").filter((r:any) => {
    const sm = r.savedMatch as SavedMatch;
    return sm.isAmountDiff || accountTypesDiffer(r.bank.accountType, r.cashier.accountType) ||
      transactionDatesDiffer(r.bank.date, r.cashier.date) || !!manualEditIssue(sm.editorNotes);
  });
  const amountDiffRows = needsReview.map((r:any) => {
    const sm = r.savedMatch as SavedMatch;
    const diff = r.bank.rawAmount - r.cashier.matchAmount;
    const issues = cleanReportNote(manualEditIssue(sm.editorNotes)) || [
      sm.isAmountDiff ? "الفاتورة مختلفة عن الواصل البنكي" : null,
      accountTypesDiffer(r.bank.accountType, r.cashier.accountType) ? "فرق نوع حساب" : null,
      transactionDatesDiffer(r.bank.date, r.cashier.date) ? "اختلاف تاريخ الفاتورة والحوالة" : null,
    ].filter(Boolean).join(" + ");
    return [r.cashier.name, r.cashier.userId || "—", r.bank.description,
      sm.type, r.bank.rawAmount, r.cashier.matchAmount,
      sm.isAmountDiff ? fmtNum(-diff) : "0", r.bank.accountType || "—", r.cashier.accountType || "—",
      r.bank.date, r.cashier.date, issues, sm.isManual ? "يدوي" : "تلقائي", cleanReportNote(sm.note)];
  });
  add("⚠️ فروقات تحتاج تدقيق",
    ["الفاتورة","رقم المستخدم","بيان الحوالة","النوع","مبلغ البنك","مبلغ الفاتورة","الفرق (بنك - فاتورة)","نوع حساب البنك","نوع حساب الفاتورة","تاريخ البنك","تاريخ الفاتورة","نوع المشكلة","نوع المطابقة","ملاحظة"],
    amountDiffRows);

  const pendingRows = results.filter(r=>r.type==="pending").map((r:any) => [
    r.cashier.name, r.cashier.userId || "—", r.bank.description, r.cashier.type, r.bank.rawAmount, r.cashier.amount,
    r.isApprox ? "تقريبي" : r.isExactAmount ? "ممتاز" : "الجوال متطابق مع فرق مبلغ", r.matchType || "عادي", r.amountDiff ? fmtNum(r.amountDiff) : "0",
    r.bank.accountType || "—", r.cashier.accountType || "—", r.accountTypeDiff ? "نعم" : "لا",
    r.isBundled ? "تجميع" : "فردي", r.bank.date, r.cashier.date
  ]);
  add("⏳ مطابقات منتظرة (غير محفوظة)",
    ["البيان (الكاشير)","رقم المستخدم","بيان البنك","النوع","مبلغ البنك","مبلغ الكاشير","الحالة","نوع المطابقة","فرق المبلغ","نوع حساب البنك","نوع حساب الكاشير","اختلاف نوع حساب","نوع التجميع","ت.البنك","ت.الكاشير"],
    pendingRows);

  const manualRows = results
    .filter((r): r is Extract<MatchResult, { type: "saved" }> => r.type === "saved" && Boolean(r.savedMatch?.isManual))
    .map((r) => {
      const sm = r.savedMatch;
      return [r.cashier.name, r.cashier.userId || "—", r.bank.description, sm.type, r.bank.rawAmount, r.cashier.amount, r.bank.date, r.cashier.date, sm.note || ""];
    });
  add("🛠️ مطابقات يدوية (محفوظة)",
    ["البيان (الكاشير)","رقم المستخدم","بيان البنك","النوع","مبلغ البنك","مبلغ الكاشير","ت.البنك","ت.الكاشير","ملاحظة"],
    manualRows);

  const visaRows = visaItems.map(v => [v.name, v.visaNumber, v.type, v.amount, v.date, v.note || "", v.movedAt]);
  add("💳 الفيزا", ["الاسم","رقم الفيزا","النوع","المبلغ","التاريخ","ملاحظة","تاريخ النقل"], visaRows);
  const jawwalRows = results.filter(r => r.type === "jawwalPay").map((r: any) => [
    r.cashier.name, r.cashier.rawName, r.cashier.userId || "—", r.cashier.type, r.cashier.amount,
    r.cashier.accountType || "—", r.cashier.date
  ]);
  add("جوال بي", ["الاسم","البيان الأصلي","رقم المستخدم","نوع الحركة","المبلغ","نوع الحساب","التاريخ"], jawwalRows);
  const mahmoudWalletRows = results.filter(r => r.type === "mahmoudWallet").map((r: any) => [
    r.cashier.name, r.cashier.rawName, r.cashier.userId || "—", r.cashier.type, r.cashier.amount,
    r.cashier.accountType || "—", r.cashier.date
  ]);
  add("محفظة شخصية", ["الاسم","البيان الأصلي","رقم المستخدم","النوع","المبلغ","نوع الحساب","التاريخ"], mahmoudWalletRows);
  const jawwalTransferRows = sourceBankRows
    .filter(row => isJawwalAccount(row.accountType))
    .map(row => [
      row.date, row.ref, row.description, row.movementType || "—", row.type,
      row.totalAmount ?? row.rawAmount, row.debit, row.credit, row.accountType || "—",
    ]);
  add("حوالات جوال بي", ["التاريخ","المرجع","البيان","نوع الحركة","الاتجاه","المبلغ الكلي","مدين بعد التطبيع","دائن بعد التطبيع","نوع الحساب"], jawwalTransferRows);
  const noteRows = [
    ...sourceBankRows.flatMap(row => rowNotes[transactionKey(row)]
      ? [["حوالة", row.description, row.type, row.rawAmount, row.accountType || "—", row.date, rowNotes[transactionKey(row)]]]
      : []),
    ...sourceCashierRows.flatMap(row => rowNotes[transactionKey(row)]
      ? [["فاتورة", row.name, row.type, row.matchAmount, row.accountType || "—", row.date, rowNotes[transactionKey(row)]]]
      : []),
  ];
  add("ملاحظات الحركات", ["نوع الحركة","البيان","الاتجاه","المبلغ","نوع الحساب","التاريخ","الملاحظة"], noteRows);

  const uCashier = results.filter(r=>r.type==="unmatchedCashier").map((r:any) => [
    r.cashier.name, r.cashier.type, r.cashier.amount, r.cashier.accountType || "—", r.cashier.userId || "—", r.cashier.date, r.reason]);
  add("فواتير لم يجد لها حوالات مطابقة", ["البيان","النوع","المبلغ","نوع الحساب المسجل","رقم المستخدم","التاريخ","السبب"], uCashier);

  const uBank = results.filter(r=>r.type==="unmatchedBank").map((r:any) => [
    r.bank.description, r.bank.type, r.bank.rawAmount, r.bank.accountType || "—", r.bank.date, r.reason]);
  add("حوالات لم يجد لها فواتير مسجلة", ["البيان","النوع","المبلغ","نوع الحساب المسجل","نوع الحركة","المبلغ الكلي","التاريخ","السبب"], uBank.map((row, index) => {
    const bank = (results.filter(r => r.type === "unmatchedBank")[index] as any)?.bank as BankRow | undefined;
    return [...row.slice(0, 4), bank?.movementType || "—", bank?.totalAmount ?? bank?.rawAmount ?? "—", row[4], row[5]];
  }));

  const heldRows = heldItems.map(h => {
    const isBank = h.kind === "bank";
    const label = isBank ? (h.data as BankRow).description : (h.data as CashierRow).name;
    const amt = isBank ? (h.data as BankRow).rawAmount : (h.data as CashierRow).amount;
    return [ isBank ? "بنك" : "كاشير", label, h.data.type, amt, h.data.accountType || "—", h.heldAt, h.note || "" ];
  });
  add("🗂️ المعلقات", ["المصدر","البيان","النوع","المبلغ","نوع الحساب","تاريخ التعليق","ملاحظة"], heldRows);

  if (resumeData) {
    try { embedResumeSheet(wb, resumeData); } catch (e) { console.error("تعذّر تضمين بيانات الاستكمال بالملف:", e); }
  }
  XLSX.writeFile(wb,"تقرير_التسوية.xlsx");
}

// ─── استكمال العمل من الملف المُصدَّر ──────────────────────────────────────────
type ResumeData = {
  bankHeaders: string[];
  bankRowsRaw: Record<string,unknown>[];
  bankMap: { date:string; desc:string; debit:string; credit:string; accountType?:string };
  bankSwap: boolean;
  bankFileSessionId?: number;
  cashHeaders: string[];
  cashRowsRaw: Record<string,unknown>[];
  cashMap: { date:string; name:string; debit:string; credit:string; accountType?:string; ref?:string; userId?:string };
  cashSwap: boolean;
  cashFileSessionId?: number;
  manualGroups: ManualMatchGroup[];
  savedMatches: SavedMatch[];
  reviewItems?: MatchReviewItem[];
  rejectedPairs: string[];
  visaItems: VisaItem[];
  heldItems: HeldItem[];
  returnedHeldBank?: BankRow[];
  returnedHeldCashier?: BankRow[] | CashierRow[];
  bankSplits?: BankSplit[];
  cutoffBatches?: CutoffBatch[];
  clearingGroups?: ClearingGroup[];
  nameAliases?: Record<string, string>;
  rejectedSpecialCashierIds?: Array<number | string>;
  mahmoudWalletCashierIds?: string[];
  rowNotes?: Record<string, string>;
  amountTolerancePercent?: number;
  bankFileSnapshots?: FileSnapshot[];
  cashFileSnapshots?: FileSnapshot[];
  walletBankFileSnapshots?: FileSnapshot[];
  walletCashFileSnapshots?: FileSnapshot[];
  jawwalBankFileSnapshots?: FileSnapshot[];
  jawwalCashFileSnapshots?: FileSnapshot[];
  stageAInvoices?: StageAResult[];
  stageADatabaseSummary?: StageASummary;
  walletBankHeaders?: string[];
  walletBankRows?: Record<string, unknown>[];
  walletBankMap?: { date: string; desc: string; debit: string; credit: string; ref: string };
  walletBankSwap?: boolean;
  walletBankFileSessionId?: number;
  walletCashHeaders?: string[];
  walletCashRows?: Record<string, unknown>[];
  walletCashMap?: { date: string; name: string; debit: string; credit: string; ref: string; userId?: string };
  walletCashSwap?: boolean;
  walletCashFileSessionId?: number;
  jawwalBankHeaders?: string[];
  jawwalBankRows?: Record<string, unknown>[];
  jawwalBankMap?: { date: string; desc: string; debit: string; credit: string; ref: string; totalAmount: string; movementAmount: string; direction: string; movementType?: string };
  jawwalBankSwap?: boolean;
  jawwalBankOptions?: { useTotalAmount: boolean; readDebitCredit: boolean };
  jawwalBankFileSessionId?: number;
  jawwalCashHeaders?: string[];
  jawwalCashRows?: Record<string, unknown>[];
  jawwalCashMap?: { date: string; name: string; debit: string; credit: string; ref: string; userId?: string };
  jawwalCashSwap?: boolean;
  jawwalCashFileSessionId?: number;
  platformNames?: { bank: string; wallet: string; jawwal: string };
  customPlatforms?: CustomPlatform[];
};

const RESUME_SHEET_NAME = "⚙️ بيانات_الاستكمال";
const RESUME_CHUNK_SIZE = 25000;

function embedResumeSheet(wb: XLSX.WorkBook, data: ResumeData) {
  const json = JSON.stringify(data);
  const MAX_JSON_SIZE = 50_000_000;
  if (json.length > MAX_JSON_SIZE) {
    throw new Error(`بيانات الاستكمال كبيرة جداً (${json.length.toLocaleString()} حرف) لإضافتها إلى ملف Excel. نزّل الملفات بعد تقليل بيانات الجلسة أو قسّم العمل إلى مشروعين.`);
  }
  const chunks: string[] = [];
  for (let i = 0; i < json.length; i += RESUME_CHUNK_SIZE) {
    chunks.push(json.slice(i, i + RESUME_CHUNK_SIZE));
  }
  const rows: string[][] = [
    ["⚠️ هذا الشيت خاص بالبرنامج - لا تعدّله أو تحذفه. هو اللي بيخلي الملف قابل للاستكمال لاحقاً عبر زر (استيراد ملف سابق)."],
    ...chunks.map(c => [c])
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, RESUME_SHEET_NAME);
}

async function extractResumeData(file: File): Promise<ResumeData> {
  const buf = await readFileBuf(file);
  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(buf, { type: "array", cellDates: true });
  } catch {
    throw new Error("تعذّرت قراءة ملف Excel. تأكد أن الملف سليم وغير محمي بكلمة مرور.");
  }
  const sheetName = wb.SheetNames.find(name => name === RESUME_SHEET_NAME || name.includes("بيانات_الاستكمال") || name.includes("استكمال"));
  if (!sheetName) {
    throw new Error("لم يُعثر على شيت بيانات الاستكمال. استورد ملف Excel الذي نزلته من زر «تصدير Excel» في المنصة، ولا تحذف شيت «بيانات_الاستكمال» منه.");
  }
  const worksheet = wb.Sheets[sheetName];
  if (!worksheet || !worksheet["!ref"]) throw new Error("شيت بيانات الاستكمال فارغ أو غير صالح.");
  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  const chunks: string[] = [];
  for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: range.s.c })];
    if (cell?.v != null) chunks.push(String(cell.v));
  }
  const json = chunks.join("");
  if (!json.trim()) throw new Error("شيت الاستكمال موجود لكنه لا يحتوي على بيانات. أعد تصدير الملف من المنصة.");
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch {
    throw new Error("بيانات الاستكمال في الملف غير مكتملة أو معدّلة. نزّل نسخة جديدة من زر «تصدير Excel» وحاول استيرادها.");
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error("تنسيق بيانات الاستكمال غير صحيح.");
  }
  const data = parsed as Partial<ResumeData>;
  const requiredArrays: Array<keyof ResumeData> = [
    "bankHeaders", "bankRowsRaw", "cashHeaders", "cashRowsRaw",
    "manualGroups", "savedMatches", "rejectedPairs", "visaItems", "heldItems",
  ];
  if (requiredArrays.some(key => !Array.isArray(data[key])) || !data.bankMap || !data.cashMap) {
    throw new Error("ملف الاستكمال لا يحتوي على كل بيانات الكشوف الأساسية المطلوبة.");
  }
  const rowArrayKeys: Array<keyof ResumeData> = [
    "bankRowsRaw", "cashRowsRaw", "walletBankRows", "walletCashRows", "jawwalBankRows", "jawwalCashRows",
  ];
  for (const key of rowArrayKeys) {
    const rows = data[key];
    if (rows !== undefined && (!Array.isArray(rows) || rows.some(row => !row || typeof row !== "object" || Array.isArray(row)))) {
      throw new Error(`بيانات «${String(key)}» في ملف الاستكمال تالفة؛ لم يتم تحميل الملف لتجنب فقد الصفوف.`);
    }
  }
  if (data.customPlatforms !== undefined && (
    !Array.isArray(data.customPlatforms) ||
    data.customPlatforms.some(platform => !platform || typeof platform !== "object" ||
      (["transfer", "invoice"] as const).some(slot => {
        const source = platform[slot];
        return source !== undefined && (
          !source || !Array.isArray(source.headers) || !Array.isArray(source.rows) ||
          source.rows.some(row => !row || typeof row !== "object" || Array.isArray(row))
        );
      }))
  )) {
    throw new Error("بيانات المنصات الإضافية في ملف الاستكمال تالفة؛ لم يتم تحميل الملف لتجنب فقد الصفوف.");
  }
  return {
    ...data,
    bankSplits: data.bankSplits ?? [],
    cutoffBatches: data.cutoffBatches ?? [],
    clearingGroups: data.clearingGroups ?? [],
  } as ResumeData;
}

function autoDetect(headers: string[], hints: string[]): string {
  const normalize=(value:string)=>value.toLowerCase().replace(/[_-]+/g," ").replace(/\s+/g," ").trim();
  const scores=headers.map(h=>({h,s:hints.reduce((score,hint)=>{
    const normalizedHeader=normalize(h), normalizedHint=normalize(hint);
    if (normalizedHeader===normalizedHint) return score+4;
    if (normalizedHint.length>2 && normalizedHeader.includes(normalizedHint)) return score+1;
    return score;
  },0)})).sort((a,b)=>b.s-a.s);
  return scores[0]?.s>0?scores[0].h:"";
}
function autoDetectBankDate(headers: string[]): string {
  const actualDate = headers.find(header => /^(?:actual date|التاريخ الفعلي|تاريخ فعلي|التاريخ الحقيقي|تاريخ حقيقي)$/.test(normalizeColumnName(header)));
  return actualDate ?? autoDetect(headers, HINTS.date);
}
function autoDetectInvoiceName(headers: string[]): string {
  const normalized = (header: string) => normalizeColumnName(header);
  const explicitName = headers.find(header => /^(?:name|customer name|client name|اسم(?: الزبون| العميل| المستخدم)?)$/.test(normalized(header)));
  if (explicitName) return explicitName;
  const invoiceDescription = headers.find(header => /^(?:ال)?بيان$/.test(normalized(header)));
  return invoiceDescription ?? autoDetect(headers, HINTS.name);
}
const HINTS={
  date:   ["date","تاريخ","value","posting"],
  desc:   ["description","narrative","detail","إيضاح","بيان","وصف","narr","payee","particular"],
  debit:  ["debit","مدين","مستلم","مبالغ مستلمة","received","deposit","money in","دخول","إيداع"],
  credit: ["credit","دائن","مدفوع","مبالغ مدفوعة","paid","withdrawal","money out","خروج","سحب"],
  name:   ["name","اسم","customer","client","employee","موظف","عميل","الزبون","بيان","وصف","narrative","detail","إيضاح"],
  userId: ["user id","userid","user number","رقم المستخدم","رقم المستخدم المسجل","معرف المستخدم","معرّف المستخدم","رقم الكاشير"],
  accountType: ["account type","account_type","نوع الحساب","نوع حساب","نوع","حساب","account"],
  ref:    ["reference","ref","transaction","trx","id","reference no","reference number","مرجع","رقم المرجع","رقم الحركة","رقم الحوالة","transaction id","txn","voucher","سند","رقم"],
  totalAmount: ["المبلغ الكلي","total amount","gross amount"],
  movementAmount: ["مبلغ الحركة","transaction amount","amount"],
  direction: ["مدين/دائن","مدين دائن","debit credit","transaction side","نوع القيد"],
};

type MergeRole = "bank" | "cash" | "jawwalBank";

async function parseAndMergeFiles(files: File[], role: MergeRole, previous?: { fileName: string; headers: string[]; rows: Record<string, unknown>[] }) {
  if (!files.length) throw new Error("اختر ملفاً واحداً على الأقل.");
  const buffers = await Promise.all(files.map(readFileBuf));
  const worker = new Worker(new URL("../workers/statement-parser.worker.ts", import.meta.url), { type: "module" });
  try {
    const parsed = await new Promise<{ headers: string[]; rows: Record<string, unknown>[]; fileName: string }>((resolve, reject) => {
      worker.onmessage = event => {
        const response = event.data as
          | { ok: true; headers: string[]; rows: Record<string, unknown>[]; fileName: string }
          | { ok: false; message: string };
        if (response.ok) resolve(response);
        else reject(new Error(response.message));
      };
      worker.onerror = event => reject(new Error(event.message || "تعذّر تشغيل معالج الملفات."));
      worker.onmessageerror = () => reject(new Error("تعذّر استلام نتيجة معالجة الملف."));
      worker.postMessage({
        role,
        baseHeaders: previous?.headers ?? [],
        files: files.map((file, index) => ({ fileName: file.name, buffer: buffers[index] })),
      }, buffers);
    });
    return {
      ...parsed,
      rows: previous ? previous.rows.concat(parsed.rows) : parsed.rows,
      fileName: previous ? `${previous.fileName} + ${parsed.fileName}` : parsed.fileName,
    };
  } finally {
    worker.terminate();
  }
}

// ─── DropZone ─────────────────────────────────────────────────────────────────
function DropZone({ file, fileCount = 1, onFiles, onClear, headers = [], rows = [] }: { file:File|null; fileCount?:number; onFiles:(files:File[], append?:boolean)=>void; onClear:()=>void; headers?:string[]; rows?:Record<string, unknown>[] }) {
  const handleDrop=useCallback((e:React.DragEvent)=>{e.preventDefault();const files=Array.from(e.dataTransfer.files);if(files.length)onFiles(files);},[onFiles]);
  if (file) return (
    <div className="flex items-center justify-between gap-3 p-3.5 bg-blue-500/5 border border-blue-500/25 rounded-lg">
      <div className="flex items-center gap-2.5">
        <FileSpreadsheet className="w-4 h-4 text-blue-600 shrink-0"/>
        <div><div className="text-sm font-medium">{fileCount > 1 ? `${fileCount} ملفات مدمجة` : file.name}</div><div className="text-xs text-muted-foreground">{rows.length.toLocaleString("ar")} صفاً إجمالياً · ${(file.size/1024).toFixed(1)} KB</div></div>
      </div>
      {headers.length > 0 && rows.length > 0 && <button onClick={() => downloadCleanedSheet(file.name, headers, rows)} className="flex items-center gap-1 rounded-md bg-emerald-50 px-2 py-1.5 text-xs font-medium text-emerald-700 hover:bg-emerald-100"><Download className="h-3.5 w-3.5"/>تنزيل Excel</button>}
      <label className="cursor-pointer rounded-md border px-2 py-1.5 text-xs hover:bg-muted">إضافة ملفات للدمج
        <input type="file" multiple accept=".xlsx,.xls,.csv" className="hidden" onChange={event=>{const files=Array.from(event.target.files ?? []);if(files.length)onFiles(files,true);event.currentTarget.value="";}}/>
      </label>
      <button onClick={onClear} aria-label="حذف الملف" title="حذف الملف" className="p-1 hover:bg-muted rounded transition-colors text-muted-foreground"><X className="w-4 h-4"/></button>
    </div>
  );
  return (
    <label onDrop={handleDrop} onDragOver={e=>e.preventDefault()}
      className="flex flex-col items-center gap-2 p-6 border-2 border-dashed border-border rounded-lg cursor-pointer hover:border-blue-500/50 hover:bg-muted/10 transition-all">
      <input type="file" multiple accept=".xlsx,.xls,.csv" className="hidden" onChange={e=>{const files=Array.from(e.target.files ?? []);if(files.length)onFiles(files);e.currentTarget.value="";}}/>
      <Upload className="w-6 h-6 text-muted-foreground"/>
      <span className="text-sm">اسحب ملفاً أو عدة ملفات للكشف أو انقر هنا</span>
      <span className="text-xs text-muted-foreground">سيتم دمج الصفوف في كشف واحد · xlsx · xls · csv</span>
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

// ─── Manual Workbench ──────────────────────────────────────────────────────
type SortMode = "amount_desc" | "amount_asc" | "name";
type MovementTypeFilter = "all" | "مدفوع" | "مستلم";

function RowNoteButton({ note, onSave }: { note?: string; onSave: (note: string) => void }) {
  const [isOpen, setIsOpen] = useState(false);
  const [draft, setDraft] = useState("");
  const openEditor = () => {
    setDraft(note ?? "");
    setIsOpen(true);
  };
  return (
    <div className="relative flex items-center gap-1">
      <button type="button" onClick={openEditor} title={note ? "عرض أو تعديل الملاحظة" : "كتابة ملاحظة"}
        aria-label={note ? "تعديل ملاحظة الحركة" : "كتابة ملاحظة للحركة"}
        className={`rounded-md border p-1.5 ${note ? "border-amber-300 bg-amber-50 text-amber-800" : "border-border text-muted-foreground hover:bg-muted"}`}>
        <MessageSquare className="h-4 w-4" />
      </button>
      {note && <span className="max-w-32 truncate text-[10px] text-amber-800" title={note}>{note}</span>}
      {isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/30 p-4">
          <div className="w-full max-w-md rounded-xl border bg-card p-4 shadow-2xl" onClick={event => event.stopPropagation()}>
            <label className="mb-2 flex items-center gap-1.5 text-sm font-semibold">
              <MessageSquare className="h-4 w-4" /> ملاحظة على الحركة
            </label>
            <textarea autoFocus value={draft} onChange={event => setDraft(event.target.value)}
              placeholder="اكتب ملاحظتك هنا..." rows={4}
              className="w-full resize-y rounded-md border bg-background px-3 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-ring" />
            <div className="mt-3 flex justify-end gap-2">
              <button type="button" onClick={() => setIsOpen(false)}
                className="rounded-md border px-3 py-2 text-xs hover:bg-muted">إلغاء</button>
              <button type="button" onClick={() => { onSave(draft.trim()); setIsOpen(false); }}
                className="flex items-center gap-1 rounded-md bg-blue-600 px-3 py-2 text-xs text-white hover:bg-blue-700">
                <Save className="h-3.5 w-3.5" /> حفظ الملاحظة
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function ListSection({
  title, icon, rows, selectedIds, onToggle, search, onSearch, sort, onSort, onHold, rowNotes, onSaveNote, onOpenReference
}: {
  title:string; icon:string; rows:(BankRow|CashierRow)[]; selectedIds:Set<string>;
  onToggle:(key:string)=>void; search:string; onSearch:(s:string)=>void;
  sort:SortMode; onSort:(s:SortMode)=>void;
  onHold:(row:BankRow|CashierRow)=>void;
  rowNotes: Record<string, string>;
  onSaveNote:(row:BankRow|CashierRow,note:string)=>void;
  onOpenReference:(row:BankRow|CashierRow)=>void;
}) {
  const [typeFilter, setTypeFilter] = useState<MovementTypeFilter>("all");
  const [amountFrom, setAmountFrom] = useState("");
  const [amountTo, setAmountTo] = useState("");
  const [scrollTop, setScrollTop] = useState(0);
  const listRef = useRef<HTMLDivElement>(null);
  const filteredRows = useMemo(() => rows.filter(row => {
    const amount = "rawAmount" in row ? row.rawAmount : row.amount;
    return (typeFilter === "all" || row.type === typeFilter) &&
      (!amountFrom || amount >= Number(amountFrom)) &&
      (!amountTo || amount <= Number(amountTo));
  }), [rows, typeFilter, amountFrom, amountTo]);
  const rowHeight = 64;
  const viewportHeight = 320;
  const firstVisibleRow = Math.max(0, Math.floor(scrollTop / rowHeight) - 5);
  const lastVisibleRow = Math.min(filteredRows.length, Math.ceil((scrollTop + viewportHeight) / rowHeight) + 5);
  const visibleRows = filteredRows.slice(firstVisibleRow, lastVisibleRow);
  useEffect(() => {
    if (listRef.current) listRef.current.scrollTop = 0;
    setScrollTop(0);
  }, [rows, search, typeFilter, amountFrom, amountTo, sort]);
  return (
    <div className="space-y-2">
      <div className="flex items-center gap-2 flex-wrap">
        <span className="font-semibold text-sm">{icon} {title}</span>
        <span className="text-xs text-muted-foreground bg-muted px-2 py-0.5 rounded-full">{filteredRows.length} / {rows.length}</span>
        <div className="flex flex-wrap items-center gap-1.5 mr-auto">
          <div className="relative">
            <Search className="absolute right-2 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-muted-foreground"/>
            <input value={search} onChange={e=>onSearch(e.target.value)} placeholder="بحث بالبيان أو المرجع..."
              className="pl-2 pr-7 py-1.5 text-xs border border-border rounded-lg bg-input-background focus:outline-none w-48 focus:ring-1 focus:ring-ring"/>
          </div>
          <select value={typeFilter} onChange={event => setTypeFilter(event.target.value as MovementTypeFilter)}
            aria-label="تصفية حسب نوع الحركة"
            className="text-xs border border-border rounded-lg px-2 py-1.5 bg-input-background focus:outline-none">
            <option value="all">كل الأنواع</option>
            <option value="مدفوع">مدفوع</option>
            <option value="مستلم">مستلم</option>
          </select>
          <details className="relative">
            <summary className="cursor-pointer list-none rounded-lg border border-border bg-input-background px-3 py-1.5 text-xs">{amountFrom || amountTo ? "نطاق المبلغ · مفعّل" : "نطاق المبلغ"}</summary>
            <div className="absolute left-0 top-full z-30 mt-1 flex w-64 gap-2 rounded-lg border bg-white p-2 shadow-lg">
              <input type="number" min="0" value={amountFrom} onChange={event => setAmountFrom(event.target.value)} placeholder="من مبلغ" aria-label="المبلغ من" className="w-1/2 rounded-lg border border-border bg-input-background px-2 py-1.5 text-xs focus:outline-none" />
              <input type="number" min="0" value={amountTo} onChange={event => setAmountTo(event.target.value)} placeholder="إلى مبلغ" aria-label="المبلغ إلى" className="w-1/2 rounded-lg border border-border bg-input-background px-2 py-1.5 text-xs focus:outline-none" />
            </div>
          </details>
          <select value={sort} onChange={e=>onSort(e.target.value as SortMode)}
            className="text-xs border border-border rounded-lg px-2 py-1.5 bg-input-background focus:outline-none">
            <option value="amount_desc">المبلغ: الأكبر أولاً</option>
            <option value="amount_asc">المبلغ: الأصغر أولاً</option>
            <option value="name">ترتيب: الاسم</option>
          </select>
        </div>
      </div>
      <div className="border border-border rounded-xl overflow-hidden">
        <div ref={listRef} className="h-[320px] overflow-auto" onScroll={event => setScrollTop(event.currentTarget.scrollTop)}>
          <table className="w-full min-w-[1050px] text-sm">
            <thead className="sticky top-0 bg-muted z-10"><tr>
              <th className="px-3 py-2.5 w-8"></th>
              <th className="px-3 py-2.5 text-right text-xs font-medium text-muted-foreground">البيان</th>
              <th className="px-3 py-2.5 text-right text-xs font-medium text-muted-foreground">نوع الحساب</th>
              <th className="px-3 py-2.5 text-right text-xs font-medium text-muted-foreground">رقم المستخدم</th>
              <th className="px-3 py-2.5 text-right text-xs font-medium text-muted-foreground">النوع</th>
              <th className="px-3 py-2.5 text-right text-xs font-medium text-muted-foreground">المبلغ</th>
              <th className="px-3 py-2.5 text-right text-xs font-medium text-muted-foreground">التاريخ</th>
              <th className="px-3 py-2.5 text-right text-xs font-medium text-muted-foreground">تعليق</th>
            </tr></thead>
            <tbody>
              {firstVisibleRow > 0 && <tr aria-hidden="true"><td colSpan={8} style={{ height: firstVisibleRow * rowHeight, padding: 0, border: 0 }} /></tr>}
              {visibleRows.map(r=>{
                const isBank="rawAmount" in r;
                const key=transactionKey(r), selected=selectedIds.has(key);
                const amt=isBank?(r as BankRow).rawAmount:(r as CashierRow).amount;
                const label=isBank?(r as BankRow).description:(r as CashierRow).name;
                const type=isBank?(r as BankRow).type:(r as CashierRow).type;
                const date=isBank?(r as BankRow).date:(r as CashierRow).date;
                const typeColor = type === "مدفوع" ? "text-red-600" : "text-green-600";
                const isFromHeld = !!(r as any)._fromHeld;
                return (
                  <tr key={key} data-default-reference={defaultReferenceForRow(r)} onClick={()=>onToggle(key)}
                  className={`h-16 border-t border-border cursor-pointer transition-colors ${selected?"bg-blue-50 border-r-2 border-r-blue-500":isFromHeld?"bg-purple-50/60 hover:bg-purple-50":"hover:bg-muted/30"}`}>
                    <td className="px-3 py-2.5">
                      <div className={`w-4 h-4 rounded border-2 flex items-center justify-center transition-colors ${selected?"bg-blue-600 border-blue-600":"border-border"}`}>
                        {selected&&<Check className="w-2.5 h-2.5 text-white"/>}
                      </div>
                    </td>
                    <td className="max-w-[560px] whitespace-normal break-words px-3 py-2.5 font-medium leading-5 [overflow-wrap:anywhere]">
                      {label}
                      <span className="mr-2 align-middle"><DefaultReferenceButton row={r} onOpen={onOpenReference} /></span>
                      {isFromHeld && (
                        <span className="mr-1.5 text-[9px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 border border-purple-300 align-middle">📦 من ملف سابق</span>
                      )}
                    </td>
                    <td className="px-3 py-2.5 text-xs">{r.accountType || "—"}</td>
                    <td className="px-3 py-2.5 font-mono text-xs">{isBank ? "—" : (r as CashierRow).userId || "—"}</td>
                    <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{type}</td>
                    <td className={`px-3 py-2.5 font-mono font-semibold ${isBank?"text-blue-700":"text-green-700"}`}>{fmtNum(amt)}</td>
                    <td className="px-3 py-2.5 text-xs text-muted-foreground">{date}</td>
                    <td className="px-3 py-2.5">
                      <div className="flex items-center gap-1" onClick={event => event.stopPropagation()}>
                        <RowNoteButton note={rowNotes[key]} onSave={note => onSaveNote(r, note)} />
                        <button
                          onClick={(e)=>{ e.stopPropagation(); onHold(r); }}
                          title="علّق هذا العنصر لمطابقته لاحقاً"
                          className="flex items-center gap-1 px-2 py-1 rounded-lg text-[10px] font-medium bg-purple-50 text-purple-700 border border-purple-200 hover:bg-purple-100 transition-colors">
                          <CreditCard className="w-3 h-3"/>علّق
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
              {filteredRows.length > lastVisibleRow && <tr aria-hidden="true"><td colSpan={8} style={{ height: (filteredRows.length - lastVisibleRow) * rowHeight, padding: 0, border: 0 }} /></tr>}
              {filteredRows.length===0&&<tr><td colSpan={8} className="px-3 py-8 text-center text-muted-foreground text-xs">لا توجد سجلات تطابق عوامل التصفية</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function ManualWorkbench({
  bankRows, cashierRows, manualGroups, onAddGroup, onRemoveGroup, onBack,
  savedCashierIds, savedBankIds, onHoldBank, onHoldCashier, onToggleReview, isInReview, rowNotes, onSaveNote, onOpenReference, focusReference
}: {
  bankRows: BankRow[]; cashierRows: CashierRow[];
  manualGroups: ManualMatchGroup[];
  onAddGroup: (g:ManualMatchGroup)=>boolean;
  onRemoveGroup: (id:string)=>void;
  onBack: ()=>void;
  savedCashierIds: Set<string>;
  savedBankIds: Set<string>;
  onHoldBank: (b:BankRow)=>void;
  onHoldCashier: (c:CashierRow)=>void;
  onToggleReview: (cashier: CashierRow, bank: BankRow, detail?: string) => void;
  isInReview: (cashier: CashierRow, bank: BankRow) => boolean;
  rowNotes: Record<string, string>;
  onSaveNote: (row: BankRow | CashierRow, note: string) => void;
  onOpenReference: (row: BankRow | CashierRow) => void;
  focusReference?: string | null;
}) {
  const [bankSearch, setBankSearch] = useState("");
  const [cashierSearch, setCashierSearch] = useState("");
  const [bankSort, setBankSort] = useState<SortMode>("amount_desc");
  const [cashierSort, setCashierSort] = useState<SortMode>("amount_desc");
  const [selectedBankIds, setSelectedBankIds] = useState<Set<string>>(new Set());
  const [selectedCashierIds, setSelectedCashierIds] = useState<Set<string>>(new Set());
  const [noteInput, setNoteInput] = useState("");
  const selectionUndoRef = useRef<Array<{ banks: Set<string>; cashiers: Set<string> }>>([]);
  useEffect(() => {
    if (!focusReference) return;
    if (bankRows.some(row => defaultReferenceForRow(row) === focusReference)) setBankSearch(focusReference);
    if (cashierRows.some(row => defaultReferenceForRow(row) === focusReference)) setCashierSearch(focusReference);
  }, [focusReference, bankRows, cashierRows]);
  const rememberSelection = () => {
    selectionUndoRef.current.push({ banks: new Set(selectedBankIds), cashiers: new Set(selectedCashierIds) });
    if (selectionUndoRef.current.length > 50) selectionUndoRef.current.shift();
  };

  const matchedBankIds = useMemo(()=>new Set(manualGroups.flatMap(g=>g.banks.map(transactionKey))),[manualGroups]);
  const matchedCashierIds = useMemo(()=>new Set(manualGroups.flatMap(g=>g.cashiers.map(transactionKey))),[manualGroups]);

  const filteredBanks = useMemo(()=>{
    let rows=uniqueTransactionRows(bankRows).filter(b => !matchedBankIds.has(transactionKey(b)) && !savedBankIds.has(transactionKey(b)));
    if (bankSearch) { const s=bankSearch.toLowerCase(); rows=rows.filter(b=>b.description.toLowerCase().includes(s)||String(b.rawAmount).includes(s)||defaultReferenceForRow(b).toLowerCase().includes(s)); }
    rows=[...rows].sort((a,b)=>{
      if (bankSort==="name") return a.description.localeCompare(b.description,"ar");
      return bankSort==="amount_asc" ? a.rawAmount-b.rawAmount : b.rawAmount-a.rawAmount;
    });
    return rows;
  },[bankRows,bankSearch,bankSort,matchedBankIds,savedBankIds]);

  const filteredCashiers = useMemo(()=>{
    let rows=uniqueTransactionRows(cashierRows).filter(c => !matchedCashierIds.has(transactionKey(c)) && !savedCashierIds.has(transactionKey(c)));
    if (cashierSearch) { const s=cashierSearch.toLowerCase(); rows=rows.filter(c=>c.name.toLowerCase().includes(s)||c.rawName.toLowerCase().includes(s)||String(c.amount).includes(s)||defaultReferenceForRow(c).toLowerCase().includes(s)); }
    rows=[...rows].sort((a,b)=>{
      if (cashierSort==="name") return a.name.localeCompare(b.name,"ar");
      return cashierSort==="amount_asc" ? a.amount-b.amount : b.amount-a.amount;
    });
    return rows;
  },[cashierRows,cashierSearch,cashierSort,matchedCashierIds,savedCashierIds]);

  const toggleBank=(key:string)=>{
    rememberSelection();
    setSelectedBankIds(p=>{
    const n=new Set(p);
    if (n.has(key)) n.delete(key); else n.add(key);
    return n;
    });
  };
  const toggleCashier=(key:string)=>{
    rememberSelection();
    setSelectedCashierIds(p=>{
    const n=new Set(p);
    if (n.has(key)) n.delete(key); else n.add(key);
    return n;
    });
  };

  const handleMatch=()=>{
    if (!selectedBankIds.size||!selectedCashierIds.size) return;
    const banks=uniqueTransactionRows(bankRows).filter(b=>selectedBankIds.has(transactionKey(b)));
    const cashiers=uniqueTransactionRows(cashierRows).filter(c=>selectedCashierIds.has(transactionKey(c)));
    if (banks.length !== 1 || cashiers.length !== 1) return;
    if (!onAddGroup({ id:`grp-${Date.now()}`, banks, cashiers, note: noteInput.trim() || undefined })) return;
    selectionUndoRef.current = [];
    setSelectedBankIds(new Set());
    setSelectedCashierIds(new Set());
    setNoteInput("");
  };

  const selectedBanks=uniqueTransactionRows(bankRows).filter(b=>selectedBankIds.has(transactionKey(b)));
  const selectedCashiers=uniqueTransactionRows(cashierRows).filter(c=>selectedCashierIds.has(transactionKey(c)));
  const selBTotal=selectedBanks.reduce((s,b)=>s+b.rawAmount,0);
  const selCTotal=selectedCashiers.reduce((s,c)=>s+c.amount,0);
  const amtMatch=selectedBankIds.size>0&&selectedCashierIds.size>0&&Math.abs(selBTotal-selCTotal)<=0.01;
  const selectedBank=selectedBanks.length === 1 ? selectedBanks[0] : undefined;
  const selectedCashier=selectedCashiers.length === 1 ? selectedCashiers[0] : undefined;
  const selectedPhones = selectedBank && selectedCashier ? movementPhones(selectedCashier, selectedBank) : null;
  const phoneDifference = !!selectedPhones &&
    !!selectedPhones.bankPhone &&
    selectedPhones.cashierPhone !== selectedPhones.bankPhone;
  const amountDifference = !!selectedBank && !!selectedCashier &&
    Math.abs(selectedBank.rawAmount - selectedCashier.matchAmount) > 0.01;
  const dateDifference = !!selectedBank && !!selectedCashier &&
    transactionDatesDiffer(selectedBank.date, selectedCashier.date);
  const accountTypeDifference = !!selectedBank && !!selectedCashier &&
    accountTypesDiffer(selectedBank.accountType, selectedCashier.accountType);
  const oneToOneSelection=selectedBanks.length === 1 && selectedCashiers.length === 1;
  const canSaveSelection=oneToOneSelection;

  useEffect(()=>{
    if (selectedBankIds.size>0 && selectedCashierIds.size>0 && !amtMatch && !noteInput) {
      setNoteInput(`اختلاف مبلغ: ${fmtNum(Math.abs(selBTotal-selCTotal))}`);
    }
  }, [selectedBankIds.size, selectedCashierIds.size, amtMatch, noteInput, selBTotal, selCTotal]);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (!(event.ctrlKey || event.metaKey) || event.code !== "KeyZ" || event.shiftKey) return;
      if (isTextEntryTarget(event.target)) return;
      const previous = selectionUndoRef.current.pop();
      if (!previous) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      setSelectedBankIds(previous.banks);
      setSelectedCashierIds(previous.cashiers);
    };
    window.addEventListener("keydown", onKeyDown, true);
    return () => window.removeEventListener("keydown", onKeyDown, true);
  }, []);

  return (
    <div dir="rtl" className="min-h-screen bg-background text-foreground">
      <div className="sticky top-0 z-20 bg-card border-b border-border px-6 py-3.5 flex items-center gap-3 shadow-sm">
        <button onClick={onBack} className="flex items-center gap-1.5 text-sm text-muted-foreground hover:text-foreground transition-colors">
          <ChevronLeft className="w-4 h-4"/>رجوع
        </button>
        <div className="w-px h-4 bg-border"/>
        <h1 className="font-bold text-sm">المطابقة اليدوية المتقدمة</h1>
        {selectedBankIds.size>0||selectedCashierIds.size>0?(
          <div className="mr-auto flex items-center gap-3 flex-wrap">
            <div className="text-xs">
              <span className="text-blue-700 font-semibold">{selectedBankIds.size} بنك</span>
              <span className="text-muted-foreground mx-1">↔</span>
              <span className="text-green-700 font-semibold">{selectedCashierIds.size} كاشير</span>
              {selectedBankIds.size>0&&selectedCashierIds.size>0&&(
                <span className={`mr-2 px-1.5 py-0.5 rounded text-xs font-medium ${amtMatch?"bg-green-100 text-green-700":"bg-amber-100 text-amber-700"}`}>
                  {amtMatch?"المبالغ متطابقة ✓":`فرق: ${fmtNum(Math.abs(selBTotal-selCTotal))}`}
                </span>
              )}
            </div>
            {selectedBankIds.size>0&&selectedCashierIds.size>0&&(
              <div className="flex items-center gap-1">
                <input value={noteInput} onChange={e=>setNoteInput(e.target.value)}
                  placeholder="سبب الربط (اختياري): اختلاف مبلغ، تحويل من الزوجة..."
                  className="text-xs border border-border rounded-lg px-2.5 py-1.5 bg-input-background focus:outline-none w-64 focus:ring-1 focus:ring-ring"/>
                <button type="button" onClick={() => {
                  const cashier = selectedCashiers[0];
                  const bank = selectedBanks[0];
                  if (cashier && bank) onToggleReview(cashier, bank, "مطابقة يدوية للمراجعة");
                }}
                  disabled={!oneToOneSelection}
                  aria-label="إضافة المطابقة المحددة إلى قائمة المراجعة"
                  title="نسخ الفاتورة والحوالة إلى قائمة المراجعة دون حفظ المطابقة"
                  className="rounded-lg border border-amber-200 p-2 text-amber-700 hover:bg-amber-50 disabled:opacity-40">
                  <Star className={`h-4 w-4 ${selectedCashiers[0] && selectedBanks[0] && isInReview(selectedCashiers[0], selectedBanks[0]) ? "fill-current" : ""}`} />
                </button>
              </div>
            )}
            <button onClick={handleMatch} disabled={!canSaveSelection}
              className="flex items-center gap-1.5 px-4 py-2 bg-green-600 text-white rounded-lg text-sm font-medium hover:bg-green-700 disabled:opacity-40 transition-colors">
              <Save className="w-4 h-4"/>{phoneDifference || amountDifference || dateDifference || accountTypeDifference ? "حفظ مع فرق للمراجعة" : "حفظ المطابقة (تؤكد فوراً)"}
            </button>
            <button onClick={()=>{rememberSelection();setSelectedBankIds(new Set());setSelectedCashierIds(new Set());setNoteInput("");}}
              className="p-2 text-muted-foreground hover:bg-muted rounded-lg transition-colors">
              <X className="w-4 h-4"/>
            </button>
          </div>
        ):(
          <div className="mr-auto text-xs text-muted-foreground">اختر من كلا القائمتين ثم اضغط "حفظ المطابقة"</div>
        )}
      </div>

      <div className="mx-auto w-full max-w-screen-2xl px-4 py-6 sm:px-6 2xl:px-10 space-y-6">
        {bankRows.length===0&&cashierRows.length===0&&(
          <div className="text-center py-16 text-muted-foreground">
            <Info className="w-10 h-10 mx-auto mb-3 opacity-40"/>
            <p>لا توجد بيانات محملة. ارجع وأضف ملفات البنك والكاشير أولاً.</p>
          </div>
        )}

        <ListSection title="كشف البنك" icon="" rows={filteredBanks} selectedIds={selectedBankIds} 
          onToggle={toggleBank} search={bankSearch} onSearch={setBankSearch}
          sort={bankSort} onSort={setBankSort}
          rowNotes={rowNotes} onSaveNote={onSaveNote}
          onHold={(row)=>{           onHoldBank(row as BankRow); setSelectedBankIds(p=>{const n=new Set(p);n.delete(transactionKey(row));return n;}); }}
          onOpenReference={onOpenReference}/>

        <ListSection title="كشف الكاشير" icon="" rows={filteredCashiers} selectedIds={selectedCashierIds}
          onToggle={toggleCashier} search={cashierSearch} onSearch={setCashierSearch}
          sort={cashierSort} onSort={setCashierSort}
          rowNotes={rowNotes} onSaveNote={onSaveNote}
          onHold={(row)=>{           onHoldCashier(row as CashierRow); setSelectedCashierIds(p=>{const n=new Set(p);n.delete(transactionKey(row));return n;}); }}
          onOpenReference={onOpenReference}/>

        {(selectedBankIds.size > 1 || selectedCashierIds.size > 1) && (
          <p className="rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 text-xs text-amber-800">
            المطابقة اليدوية تمنع تكرار الحركة، لذلك اختر حوالة واحدة وفاتورة واحدة فقط. لتوزيع حوالة على عدة فواتير استخدم أداة التجزئة.
          </p>
        )}
        {oneToOneSelection && (phoneDifference || amountDifference || dateDifference || accountTypeDifference) && (
          <p className="rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 text-xs leading-6 text-amber-900">
            ستحفظ هذه كمطابقة يدوية وتظهر في تقرير الفروقات للمراجعة.
            {phoneDifference && ` الجوال الوارد في الفاتورة: ${selectedPhones?.cashierPhone ?? "غير موجود"} · الجوال الوارد في الحوالة، وهو المعتمد: ${selectedPhones?.bankPhone ?? "غير موجود"}.`}
            {amountDifference && ` فرق المبلغ: ${fmtNum(Math.abs(selectedBank!.rawAmount - selectedCashier!.matchAmount))} ₪.`}
            {dateDifference && ` تاريخ الفاتورة (${selectedCashier!.date}) مختلف عن تاريخ الحوالة (${selectedBank!.date})؛ يجب أن يكونا في اليوم نفسه للصندوق.`}
            {accountTypeDifference && ` نوع الحساب مختلف: الحوالة (${selectedBank!.accountType}) والفاتورة (${selectedCashier!.accountType}).`}
            {phoneDifference && selectedPhones?.bankPhone && " سيُحفظ الربط اليدوي وتظهر ملاحظة تعديل الجوال في الفروقات؛ لا تتغير بيانات المصدر."}
          </p>
        )}

        {(selectedBankIds.size>0&&selectedCashierIds.size>0)&&(
          <button onClick={handleMatch} disabled={!canSaveSelection}
            className="w-full py-3.5 rounded-xl font-medium flex items-center justify-center gap-2 text-sm transition-colors bg-green-600 hover:bg-green-700 text-white">
            <Save className="w-5 h-5"/>
            حفظ {selectedBankIds.size} حوالة بنكية ↔ {selectedCashierIds.size} فاتورة كاشier (تؤكد فوراً)
            {!amtMatch&&<span className="text-xs opacity-80">(مجاميع مختلفة)</span>}
          </button>
        )}

        {manualGroups.length>0&&(
          <div className="space-y-3">
            <h2 className="font-semibold text-sm">المجموعات المربوطة يدوياً (سيتم حفظها فوراً)</h2>
            <div className="text-xs text-muted-foreground">المطابقة اليدوية تذهب للمؤكدة مباشرة</div>
            <div className="space-y-2">
              {manualGroups.map(g=>{
                const gDiff = g.banks.reduce((s,b)=>s+b.rawAmount,0) - g.cashiers.reduce((s,c)=>s+c.amount,0);
                const hasDiff = Math.abs(gDiff) > 0.01;
                return (
                <div key={g.id} className="border border-green-200 bg-green-50/60 rounded-xl p-4">
                  <div className="flex items-start gap-3">
                    <div className="flex-1 grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <div className="text-xs font-medium text-muted-foreground mb-2">البنك ({g.banks.length})</div>
                        {g.banks.map(b=>(
                          <div key={transactionKey(b)} className="flex items-center gap-2 text-sm mb-1">
                            <span className="font-medium truncate">{b.description}</span>
                            <span className={`text-xs font-semibold ${b.type === "مدفوع" ? "text-red-600" : "text-green-600"}`}>{b.type}</span>
                            <span className="text-blue-700 font-mono text-xs shrink-0">{fmtNum(b.rawAmount)}</span>
                          </div>
                        ))}
                        <div className="text-xs text-blue-700 font-semibold mt-1">المجموع: {fmtNum(g.banks.reduce((s,b)=>s+b.rawAmount,0))}</div>
                      </div>
                      <div>
                        <div className="text-xs font-medium text-muted-foreground mb-2">الكاشير ({g.cashiers.length})</div>
                        {g.cashiers.map(c=>(
                          <div key={transactionKey(c)} className="flex items-center gap-2 text-sm mb-1">
                            <span className="font-medium">{c.name}</span>
                            <span className={`text-xs font-semibold ${c.type === "مدفوع" ? "text-red-600" : "text-green-600"}`}>{c.type}</span>
                            <span className="text-green-700 font-mono text-xs shrink-0">{fmtNum(c.amount)}</span>
                          </div>
                        ))}
                        <div className="text-xs text-green-700 font-semibold mt-1">المجموع: {fmtNum(g.cashiers.reduce((s,c)=>s+c.amount,0))}</div>
                      </div>
                    </div>
                    <button onClick={()=>onRemoveGroup(g.id)} className="p-1.5 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors shrink-0">
                      <Trash2 className="w-4 h-4"/>
                    </button>
                  </div>
                  <div className="mt-3 pt-3 border-t border-green-200/70 flex items-center gap-2 flex-wrap">
                    <span className="text-xs font-medium px-2 py-1 rounded-lg bg-green-100 text-green-800 flex items-center gap-1">
                      <Check className="w-3 h-3"/>سيتم حفظها في المؤكدة
                    </span>
                    {(hasDiff||g.note)&&(
                      <>
                        {hasDiff&&(
                          <span className="text-xs font-medium px-2 py-1 rounded-lg bg-orange-100 text-orange-800 flex items-center gap-1">
                            <AlertTriangle className="w-3 h-3"/>فرق مبلغ {fmtNum(Math.abs(gDiff))}
                          </span>
                        )}
                        {g.note&&(
                          <span className="text-xs px-2 py-1 rounded-lg bg-white border border-amber-300 text-amber-800">📝 {g.note}</span>
                        )}
                      </>
                    )}
                  </div>
                </div>
                );
              })}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── مساعد المطابقة ────────────────────────────────────────────────────────
function MatchAssistant({
  cashierRows, bankRows, rejectedPairs, onMatch, onHold, onHoldBank, onReject, onBack, onOpenReference, assistantSkipRef
}: {
  cashierRows: CashierRow[];
  bankRows: BankRow[];
  rejectedPairs: Set<string>;
  onMatch: (c: CashierRow, b: BankRow) => void;
  onHold: (c: CashierRow) => void;
  onHoldBank: (b: BankRow) => void;
  onReject: (c: CashierRow, b: BankRow) => void;
  onBack: () => void;
  onOpenReference: (row: BankRow | CashierRow) => void;
  assistantSkipRef: React.MutableRefObject<((row: CashierRow) => boolean) | null>;
}) {
  const [skipped, setSkipped] = useState<Set<string>>(new Set());
  const [search, setSearch] = useState("");
  const deferredSearch = useDeferredValue(search);

  const queue = useMemo(
    () => cashierRows.filter(c => !skipped.has(transactionKey(c))),
    [cashierRows, skipped]
  );

  const current = queue[0] ?? null;
  const skipCurrent = (row = current) => {
    if (!row || !current || transactionKey(row) !== transactionKey(current)) return false;
    const currentKey = transactionKey(current);
    setSkipped(previous => new Set(previous).add(currentKey));
    return true;
  };

  useEffect(() => {
    assistantSkipRef.current = row => skipCurrent(row);
    return () => { assistantSkipRef.current = null; };
  }, [assistantSkipRef, current]);

  useEffect(() => { setSearch(""); }, [current?.id]);

  const candidates = useMemo(() => {
    if (!current) return [];
    let list = bankRows.filter(b => {
      if (b.type !== current.type) return false;
      return !wasPairRejected(rejectedPairs, current, b);
    });
    if (deferredSearch.trim()) {
      const s = deferredSearch.trim().toLowerCase();
      list = list.filter(b => b.description.toLowerCase().includes(s) || String(b.rawAmount).includes(s) || defaultReferenceForRow(b).toLowerCase().includes(s));
    }
    return list
      .map(b => ({
        bank: b,
        diff: Math.abs(b.rawAmount - current.matchAmount),
        exact: Math.abs(b.rawAmount - current.matchAmount) <= 0.01,
        accountDiff: accountTypesDiffer(b.accountType, current.accountType),
        sim: nameSim(current.name, b.description)
      }))
      .sort((a, b) => a.diff - b.diff)
      .slice(0, 40);
  }, [current, bankRows, deferredSearch, rejectedPairs]);

  return (
    <div dir="rtl" className="min-h-screen bg-background text-foreground">
      <div className="sticky top-0 z-20 bg-card border-b border-border px-6 py-3.5 flex items-center gap-3 shadow-sm flex-wrap">
        <button onClick={onBack} className="flex items-center gap-1.5 text-sm text-muted-foreground hover:text-foreground transition-colors">
          <ChevronLeft className="w-4 h-4"/>رجوع
        </button>
        <div className="w-px h-4 bg-border"/>
        <h1 className="font-bold text-sm">مساعد المطابقة</h1>
        <span className="text-xs text-muted-foreground bg-muted px-2 py-0.5 rounded-full">{queue.length} فاتورة متبقية</span>
      </div>

      <div className="max-w-4xl mx-auto px-6 py-6 space-y-5">
        {!current ? (
          <div className="text-center py-20 text-muted-foreground">
            <Check className="w-12 h-12 mx-auto mb-3 text-green-500"/>
            <p className="text-sm font-medium">{skipped.size ? `تم تخطي ${skipped.size} فاتورة. لا توجد فواتير أخرى في قائمة المساعد.` : "ممتاز! ما في فواتير كاشير قيد الانتظار حالياً 🎉"}</p>
            {skipped.size > 0 && <button type="button" onClick={onBack} className="mt-4 rounded-lg bg-indigo-600 px-4 py-2 text-sm font-medium text-white hover:bg-indigo-700">العودة إلى التسوية</button>}
          </div>
        ) : (
          <>
            <div className="border-2 border-indigo-300 bg-indigo-50/50 rounded-2xl p-5">
              <div className="flex items-center justify-between gap-3 flex-wrap">
                <div>
                  <div className="text-xs text-indigo-600 font-medium mb-1">فاتورة الكاشير الحالية</div>
                  <div className="flex items-center gap-2 text-lg font-bold">{current.name}<DefaultReferenceButton row={current} onOpen={onOpenReference} /></div>
                  <div className="flex items-center gap-2 mt-1.5 flex-wrap">
                    <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${current.type === "مدفوع" ? "bg-red-100 text-red-700" : "bg-green-100 text-green-700"}`}>{current.type}</span>
                    <span className="text-lg font-mono font-bold text-indigo-700">{fmtNum(current.matchAmount)}</span>
                    {current.accountType && (
                      <span className="text-xs px-2 py-0.5 rounded-full bg-purple-100 text-purple-700 border border-purple-200">{current.accountType}</span>
                    )}
                    <span className="text-xs text-muted-foreground">{current.date}</span>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button onClick={() => skipCurrent()}
                    className="flex items-center gap-1.5 px-3 py-2 bg-card border border-border rounded-lg text-xs font-medium hover:bg-muted/40 transition-colors">
                    ⏭️ تخطي الفاتورة
                  </button>
                  <button onClick={() => onHold(current)}
                    className="flex items-center gap-1.5 px-3 py-2 bg-purple-50 text-purple-700 border border-purple-200 rounded-lg text-xs font-medium hover:bg-purple-100 transition-colors">
                    <CreditCard className="w-3.5 h-3.5"/>نقل للمعلقات
                  </button>
                </div>
              </div>
            </div>

            <div className="relative">
              <Search className="absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground"/>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="بحث بالبيان أو المبلغ داخل الحوالات..."
                className="w-full pl-3 pr-9 py-2.5 text-sm border border-border rounded-lg bg-input-background focus:outline-none focus:ring-1 focus:ring-ring"/>
            </div>

            <div className="space-y-2">
              <p className="text-xs font-medium text-muted-foreground">اضغط على بيان الحوالة لمطابقتها فوراً، أو استخدم "مش هاي"/"علّق" لإبعاد حوالة غير مرتبطة:</p>
              <div className="border border-border rounded-xl overflow-hidden divide-y divide-border">
                {candidates.map(({bank, diff, exact, accountDiff, sim}) => (
                  <div key={transactionKey(bank)} data-default-reference={`${defaultReferenceForRow(current)}|${defaultReferenceForRow(bank)}`} className="w-full px-4 py-3 hover:bg-indigo-50/40 transition-colors flex items-center gap-3">
                    <button onClick={() => onMatch(current, bank)} className="flex-1 min-w-0 text-right">
                      <div className="font-medium truncate">{bank.description}</div>
                      <div className="flex items-center gap-2 mt-1 flex-wrap">
                        <span className="font-mono font-bold text-blue-700">{fmtNum(bank.rawAmount)}</span>
                        {exact ? (
                          <span className="text-[10px] px-1.5 py-0.5 rounded bg-green-100 text-green-700 border border-green-300">✓ نفس المبلغ</span>
                        ) : (
                          <span className="text-[10px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 border border-amber-300">فرق مبلغ: {fmtNum(diff)}</span>
                        )}
                        {bank.accountType && (
                          <span className={`text-[10px] px-1.5 py-0.5 rounded border ${accountDiff ? "bg-orange-100 text-orange-700 border-orange-300" : "bg-purple-50 text-purple-700 border-purple-200"}`}>
                            {bank.accountType}{accountDiff ? " (مختلف)" : ""}
                          </span>
                        )}
                        <span className="text-[10px] text-muted-foreground">تشابه اسم: {(sim*100).toFixed(0)}٪</span>
                        <span className="text-[10px] text-muted-foreground">{bank.date}</span>
                      </div>
                    </button>
                    <div className="flex items-center gap-1 shrink-0">
                      <DefaultReferenceButton row={bank} onOpen={onOpenReference} />
                      <button onClick={() => onReject(current, bank)}
                        title="مش هاي — أبعدها عن هذه الفاتورة بس خلّيها متاحة لغيرها"
                        className="px-2 py-1.5 bg-gray-50 text-gray-500 border border-gray-200 rounded-lg text-[10px] font-medium hover:bg-gray-100 transition-colors">
                        ✕ مش هاي
                      </button>
                      <button onClick={() => onHoldBank(bank)}
                        title="علّق هذه الحوالة كلياً"
                        className="px-2 py-1.5 bg-purple-50 text-purple-700 border border-purple-200 rounded-lg text-[10px] font-medium hover:bg-purple-100 transition-colors flex items-center gap-1">
                        <CreditCard className="w-3 h-3"/>علّق
                      </button>
                    </div>
                  </div>
                ))}
                {!candidates.length && (
                  <p className="px-4 py-8 text-center text-xs text-muted-foreground">لا توجد حوالات بنفس نوع الحركة، أو كلها استُبعدت بالبحث/الرفض.</p>
                )}
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ─── Projects ──────────────────────────────────────────────────────────────
interface SavedProject {
  id: string;
  name: string;
  savedAt: string;
  bankHeaders: string[];
  bankRowsRaw: Record<string,unknown>[];
  bankMap: { date:string; desc:string; debit:string; credit:string; accountType?:string };
  bankSwap: boolean;
  bankFileSessionId?: number;
  cashHeaders: string[];
  cashRowsRaw: Record<string,unknown>[];
  cashMap: { date:string; name:string; debit:string; credit:string; accountType?:string; ref?:string; userId?:string };
  cashSwap: boolean;
  cashFileSessionId?: number;
  manualGroups: ManualMatchGroup[];
  savedMatches: SavedMatch[];
  reviewItems?: MatchReviewItem[];
  rejectedPairs: string[];
  visaItems: VisaItem[];
  heldItems: HeldItem[];
  returnedHeldBank?: BankRow[];
  returnedHeldCashier?: CashierRow[];
  bankSplits?: BankSplit[];
  cutoffBatches?: CutoffBatch[];
  clearingGroups?: ClearingGroup[];
  nameAliases?: Record<string, string>;
  rejectedSpecialCashierIds?: Array<number | string>;
  mahmoudWalletCashierIds?: string[];
  rowNotes?: Record<string, string>;
  amountTolerancePercent?: number;
  bankFileSnapshots?: FileSnapshot[];
  cashFileSnapshots?: FileSnapshot[];
  walletBankFileSnapshots?: FileSnapshot[];
  walletCashFileSnapshots?: FileSnapshot[];
  jawwalBankFileSnapshots?: FileSnapshot[];
  jawwalCashFileSnapshots?: FileSnapshot[];
  stageAInvoices?: StageAResult[];
  stageADatabaseSummary?: StageASummary;
  stageASession?: Record<string, unknown> | null;
  stageASourceFiles?: Array<{ id: string; stage: "stageA" | "stageB" | "stageC"; source: string; fileName: string; mimeType: string }>;
  walletBankHeaders?: string[];
  walletBankRows?: Record<string, unknown>[];
  walletBankMap?: { date: string; desc: string; debit: string; credit: string; ref: string };
  walletBankSwap?: boolean;
  walletCashHeaders?: string[];
  walletCashRows?: Record<string, unknown>[];
  walletCashMap?: { date: string; name: string; debit: string; credit: string; ref: string; userId?: string };
  walletCashSwap?: boolean;
  jawwalBankHeaders?: string[];
  jawwalBankRows?: Record<string, unknown>[];
  jawwalBankMap?: { date: string; desc: string; debit: string; credit: string; ref: string; totalAmount: string; movementAmount: string; direction: string; movementType?: string };
  jawwalBankSwap?: boolean;
  jawwalCashHeaders?: string[];
  jawwalCashRows?: Record<string, unknown>[];
  jawwalCashMap?: { date: string; name: string; debit: string; credit: string; ref: string; userId?: string };
  jawwalCashSwap?: boolean;
  platformNames?: { bank: string; wallet: string; jawwal: string };
  customPlatforms?: CustomPlatform[];
  walletBankFileSessionId?: number;
  walletCashFileSessionId?: number;
  jawwalBankFileSessionId?: number;
  jawwalCashFileSessionId?: number;
  jawwalBankOptions?: { useTotalAmount: boolean; readDebitCredit: boolean };
}

async function loadProjectsList(): Promise<SavedProject[]> {
  if (centralStorageEnabled) {
    const storedRemote = await centralStorageGet<SavedProject[]>("recon_projects");
    if (storedRemote !== undefined) {
      if (!Array.isArray(storedRemote)) throw new Error("صيغة قائمة المشاريع المحفوظة في قاعدة البيانات غير صالحة.");
      return storedRemote;
    }
    if (isCentralAdmin()) {
      const stored = await projectsIDBGet<SavedProject[]>("recon_projects");
      if (Array.isArray(stored) && stored.length) {
        await centralStorageSet("recon_projects", stored);
        return stored;
      }
      const legacy = await storageGet<SavedProject[]>("recon_projects", []);
      if (legacy.length) {
        await centralStorageSet("recon_projects", legacy);
        return legacy;
      }
    }
    return [];
  }
  let indexedDbError: unknown;
  try {
    const stored = await projectsIDBGet<SavedProject[]>("recon_projects");
    if (Array.isArray(stored)) return stored;
  } catch (error) {
    indexedDbError = error;
    console.error("تعذر قراءة المشاريع من التخزين الدائم:", error);
  }

  const legacyProjects = await storageGet<SavedProject[]>("recon_projects", []);
  if (!Array.isArray(legacyProjects) || legacyProjects.length === 0) {
    if (indexedDbError) throw new Error("تعذر الوصول إلى التخزين الدائم للمشاريع. أعد تحميل الصفحة وحاول مجدداً.", { cause: indexedDbError });
    return [];
  }
  try {
    await projectsIDBSet("recon_projects", legacyProjects);
  } catch (error) {
    console.error("تعذرت ترقية المشاريع القديمة إلى التخزين الدائم:", error);
  }
  return legacyProjects;
}
async function persistProjectsList(list: SavedProject[]): Promise<void> {
  if (centralStorageEnabled) {
    await storageSet("recon_projects", list);
    const saved = await storageGet<SavedProject[]>("recon_projects", []);
    if (saved.length !== list.length || saved.some((project, index) => project.id !== list[index]?.id)) {
      throw new Error("لم يتم التحقق من حفظ قائمة المشاريع في قاعدة البيانات.");
    }
    return;
  }
  await projectsIDBSet("recon_projects", list);
  const saved = await projectsIDBGet<SavedProject[]>("recon_projects");
  if (!Array.isArray(saved) || saved.length !== list.length || saved.some((project, index) => project.id !== list[index]?.id)) {
    throw new Error("لم يتم التحقق من حفظ قائمة المشاريع.");
  }
}

const PROJECTS_DB_NAME = "recon_store";
const PROJECTS_DB_STORE = "kv";
let projectsDatabase: Promise<IDBDatabase> | null = null;

function openProjectsDatabase(): Promise<IDBDatabase> {
  if (projectsDatabase) return projectsDatabase;
  const opening = new Promise<IDBDatabase>((resolve, reject) => {
    if (!("indexedDB" in window)) {
      reject(new Error("المتصفح لا يدعم التخزين الدائم للمشاريع."));
      return;
    }
    const request = indexedDB.open(PROJECTS_DB_NAME, 1);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(PROJECTS_DB_STORE)) {
        request.result.createObjectStore(PROJECTS_DB_STORE);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("تعذر فتح قاعدة بيانات المشاريع."));
    request.onblocked = () => reject(new Error("قاعدة بيانات المشاريع مشغولة في نافذة أخرى."));
  });
  const cached = opening.catch(error => {
    projectsDatabase = null;
    throw error;
  });
  projectsDatabase = cached;
  return cached;
}

async function projectsIDBGet<T>(key: string): Promise<T | undefined> {
  const database = await openProjectsDatabase();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(PROJECTS_DB_STORE, "readonly");
    const request = transaction.objectStore(PROJECTS_DB_STORE).get(key);
    request.onsuccess = () => resolve(request.result as T | undefined);
    request.onerror = () => reject(request.error ?? new Error("تعذرت قراءة المشاريع المحفوظة."));
  });
}

async function projectsIDBSet(key: string, value: unknown): Promise<void> {
  const database = await openProjectsDatabase();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(PROJECTS_DB_STORE, "readwrite");
    transaction.objectStore(PROJECTS_DB_STORE).put(value, key);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error ?? new Error("تعذر حفظ المشاريع."));
    transaction.onabort = () => reject(transaction.error ?? new Error("أُلغيت عملية حفظ المشاريع."));
  });
}

function ProjectsBar({
  onSave, onLoad, onDelete
}: {
  onSave: (name:string)=>Promise<void>;
  onLoad: (p:SavedProject)=>void;
  onDelete: (id:string)=>Promise<void>;
}) {
  const [open, setOpen] = useState(false);
  const [projects, setProjects] = useState<SavedProject[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const refresh = async () => {
    setLoading(true);
    setError("");
    try {
      const list = await loadProjectsList();
      setProjects(list);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذر تحميل المشاريع المحفوظة.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { refresh(); }, []);

  return (
    <div className="relative">
      <div className="flex items-center gap-2">
        <button
          onClick={() => {
            const name = window.prompt("اسم المشروع (مثال: تسوية 17-07-2026):", "");
            if (name && name.trim()) {
              setError("");
              void onSave(name.trim()).then(refresh).catch(cause => {
                setError(cause instanceof Error ? cause.message : "تعذر حفظ المشروع. تحقق من مساحة التخزين ثم أعد المحاولة.");
                setOpen(true);
              });
            }
          }}
          className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700 transition-colors">
          <FolderPlus className="w-4 h-4"/>حفظ كمشروع
        </button>
        <button
          onClick={() => { refresh(); setOpen(o=>!o); }}
          className="flex items-center gap-2 px-4 py-2 bg-card border border-border rounded-lg text-sm font-medium hover:bg-muted/40 transition-colors">
          <FolderOpen className="w-4 h-4"/>المشاريع المحفوظة ({projects.length})
        </button>
      </div>
      {open && (
        <div className="absolute z-30 mt-2 w-96 bg-card border border-border rounded-xl shadow-lg p-2 max-h-80 overflow-y-auto">
          {loading && <p className="text-xs text-muted-foreground p-3 text-center">جارِ التحميل...</p>}
          {error && <p role="alert" className="m-2 rounded-lg bg-red-50 p-3 text-xs text-red-700">{error}</p>}
          {!loading && projects.length === 0 && (
            <p className="text-xs text-muted-foreground p-3 text-center">لا توجد مشاريع محفوظة بعد</p>
          )}
          {projects.map(p=>(
            <div key={p.id} className="flex items-center justify-between gap-2 p-2.5 rounded-lg hover:bg-muted/30">
              <button
                type="button"
                onClick={()=>{ onLoad(p); setOpen(false); }}
                aria-label={`فتح مشروع ${p.name}`}
                className="min-w-0 flex-1 text-right"
              >
                <span className="block truncate text-sm font-medium">{p.name}</span>
                <span className="block text-[10px] text-muted-foreground">{p.savedAt} · {p.savedMatches.length} مطابقة مؤكدة</span>
              </button>
              <div className="flex items-center gap-1 shrink-0">
                <button type="button" onClick={()=>{ onLoad(p); setOpen(false); }}
                  className="px-2.5 py-1.5 bg-blue-50 text-blue-700 border border-blue-200 rounded-lg text-xs hover:bg-blue-100 transition-colors">
                  فتح
                </button>
                <button type="button" onClick={async ()=>{
                    if (window.confirm(`حذف مشروع "${p.name}" نهائياً؟`)) {
                      try { await onDelete(p.id); await refresh(); }
                      catch (cause) { setError(cause instanceof Error ? cause.message : "تعذر حذف المشروع."); }
                    }
                  }}
                  className="p-1.5 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors">
                  <Trash2 className="w-3.5 h-3.5"/>
                </button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ConfidenceBadge({ score }: { score?: number }) {
  if (score == null) return <span className="text-xs text-muted-foreground">—</span>;
  const tone = score > 90 ? "bg-green-100 text-green-700 border-green-200" : score >= 70 ? "bg-amber-100 text-amber-700 border-amber-200" : "bg-red-100 text-red-700 border-red-200";
  return <span className={`inline-flex items-center rounded-full border px-2 py-0.5 text-[11px] font-semibold ${tone}`}>{score}% ثقة</span>;
}

function FilterBar({
  search, onSearch, type, onType, minAmount, onMinAmount, maxAmount, onMaxAmount
}: {
  search: string; onSearch: (value: string) => void;
  type: "all" | "مدفوع" | "مستلم"; onType: (value: "all" | "مدفوع" | "مستلم") => void;
  minAmount: string; onMinAmount: (value: string) => void;
  maxAmount: string; onMaxAmount: (value: string) => void;
}) {
  return (
    <div className="flex flex-wrap items-center gap-2 border-b bg-slate-50/80 px-4 py-3">
      <div className="relative min-w-52 flex-1">
        <Search className="absolute right-2.5 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
        <input value={search} onChange={e => onSearch(e.target.value)} placeholder="ابحث بالاسم أو البيان..." className="w-full rounded-lg border bg-white py-2 pl-3 pr-9 text-xs outline-none focus:ring-2 focus:ring-blue-200" />
      </div>
      <select value={type} onChange={e => onType(e.target.value as "all" | "مدفوع" | "مستلم")} className="rounded-lg border bg-white px-3 py-2 text-xs outline-none">
        <option value="all">كل الأنواع</option><option value="مدفوع">مدفوع</option><option value="مستلم">مستلم</option>
      </select>
      <details className="relative">
        <summary className="cursor-pointer list-none rounded-lg border bg-white px-3 py-2 text-xs">{minAmount || maxAmount ? "نطاق المبلغ · مفعّل" : "نطاق المبلغ"}</summary>
        <div className="absolute left-0 top-full z-30 mt-1 flex w-64 gap-2 rounded-lg border bg-white p-2 shadow-lg">
          <input value={minAmount} onChange={e => onMinAmount(e.target.value)} type="number" min="0" placeholder="من مبلغ" className="w-1/2 rounded-lg border bg-white px-2.5 py-2 text-xs outline-none" />
          <input value={maxAmount} onChange={e => onMaxAmount(e.target.value)} type="number" min="0" placeholder="إلى مبلغ" className="w-1/2 rounded-lg border bg-white px-2.5 py-2 text-xs outline-none" />
        </div>
      </details>
    </div>
  );
}

// ─── لوحة السحب والاستبدال (تظهر جمب الجدول وتزيحه) ──────────────────────────
function MatchDrawer({
  mode, cashier, oldBank, nameSearch, amountFrom, amountTo, banks, cashiers, draggedItemId,
  onNameSearch, onAmountFrom, onAmountTo, onDragStart, onDragEnd, onReplaceBank, onReplaceCashier, onClose
}: {
  mode: "bank" | "cashier"; cashier: CashierRow; oldBank: BankRow;
  nameSearch: string; amountFrom: string; amountTo: string;
  banks: BankRow[]; cashiers: CashierRow[]; draggedItemId: number | null;
  onNameSearch: (v:string)=>void; onAmountFrom: (v:string)=>void; onAmountTo: (v:string)=>void;
  onDragStart: (id:number)=>void; onDragEnd: ()=>void;
  onReplaceBank: (b:BankRow)=>void; onReplaceCashier: (c:CashierRow)=>void; onClose: ()=>void;
}) {
  const [position, setPosition] = useState({ x: 0, y: 0 });
  const dragState = useRef<{ startX: number; startY: number; originX: number; originY: number } | null>(null);
  const handleDragMove = (event: PointerEvent) => {
    const current = dragState.current;
    if (!current) return;
    setPosition({
      x: current.originX + event.clientX - current.startX,
      y: current.originY + event.clientY - current.startY
    });
  };
  const stopDragging = () => {
    dragState.current = null;
    window.removeEventListener("pointermove", handleDragMove);
    window.removeEventListener("pointerup", stopDragging);
  };
  const startDragging = (event: React.PointerEvent<HTMLDivElement>) => {
    if ((event.target as HTMLElement).closest("button, input, select")) return;
    event.currentTarget.setPointerCapture(event.pointerId);
    dragState.current = { startX: event.clientX, startY: event.clientY, originX: position.x, originY: position.y };
    window.addEventListener("pointermove", handleDragMove);
    window.addEventListener("pointerup", stopDragging);
  };
  return (
    <div
      className="fixed z-50 w-[550px] max-w-[95vw] h-[70vh] max-h-[calc(100vh-2rem)] bg-white border border-border rounded-xl shadow-2xl flex flex-col overflow-hidden isolate"
      style={{ top: "15vh", right: "1rem", transform: `translate(${position.x}px, ${position.y}px)` }}
      dir="rtl"
    >
      <div
        onPointerDown={startDragging}
        className={`flex cursor-move items-center justify-between gap-3 px-4 py-3.5 text-white shrink-0 select-none ${mode === "bank" ? "bg-gradient-to-l from-blue-600 to-indigo-600" : "bg-gradient-to-l from-emerald-600 to-teal-600"}`}
      >
        <div className="min-w-0 flex items-center gap-2.5">
          <span className="flex items-center justify-center w-9 h-9 rounded-xl bg-white/15 shrink-0">
            {mode === "bank" ? <Landmark className="w-5 h-5"/> : <Users className="w-5 h-5"/>}
          </span>
          <div className="min-w-0">
            <div className="text-[11px] text-white/80">{mode === "bank" ? "قائمة سحب الحوالات" : "قائمة سحب فواتير الكاشير"} · اسحب من هنا لتحريك اللوحة</div>
            <div className="max-w-[min(480px,42vw)] whitespace-normal break-words text-sm font-bold leading-6">{mode === "bank" ? cashier.name : oldBank.description}</div>
          </div>
        </div>
        <button onClick={onClose} className="p-1.5 hover:bg-white/15 rounded-lg transition-colors shrink-0"><X className="w-4 h-4"/></button>
      </div>

      <div className="px-4 py-3 border-b border-border bg-slate-50 space-y-1.5 shrink-0">
        <div className="flex items-center gap-2 flex-wrap text-xs">
          <span className={`font-semibold px-2 py-0.5 rounded-full ${(mode === "bank" ? cashier.type : oldBank.type) === "مدفوع" ? "bg-red-100 text-red-700" : "bg-green-100 text-green-700"}`}>
            {mode === "bank" ? cashier.type : oldBank.type}
          </span>
          <span className="font-mono font-bold text-indigo-700">{fmtNum(mode === "bank" ? cashier.matchAmount : oldBank.rawAmount)}</span>
        </div>
        <div className="text-[11px] text-muted-foreground">
          {mode === "bank"
            ? <>سيتم استبدال الحوالة الحالية: <span className="text-foreground font-medium">{oldBank.description}</span> ({fmtNum(oldBank.rawAmount)})</>
            : <>سيتم استبدال الفاتورة الحالية: <span className="text-foreground font-medium">{cashier.name}</span> ({fmtNum(cashier.matchAmount)})</>}
        </div>
      </div>

      <div className="px-4 py-3 border-b border-border space-y-2 bg-white shrink-0">
        <div className="relative">
          <Search className="absolute right-2.5 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground"/>
          <input value={nameSearch} onChange={e => onNameSearch(e.target.value)}
            placeholder={mode === "bank" ? "بحث ببيان الحوالة..." : "بحث باسم الفاتورة..."}
            className="w-full pl-3 pr-9 py-2 text-sm border border-border rounded-lg bg-input-background focus:outline-none focus:ring-1 focus:ring-ring"/>
        </div>
        <details className="relative">
          <summary className="cursor-pointer list-none rounded-lg border border-border bg-input-background px-3 py-2 text-xs">{amountFrom || amountTo ? "نطاق المبلغ · مفعّل" : "تصفية حسب المبلغ"}</summary>
          <div className="mt-2 flex items-center gap-2">
            <input type="number" value={amountFrom} onChange={e => onAmountFrom(e.target.value)} placeholder="من مبلغ" className="flex-1 rounded-lg border border-border bg-input-background px-2.5 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-ring"/>
            <span className="text-muted-foreground text-xs">إلى</span>
            <input type="number" value={amountTo} onChange={e => onAmountTo(e.target.value)} placeholder="إلى مبلغ" className="flex-1 rounded-lg border border-border bg-input-background px-2.5 py-2 text-sm focus:outline-none focus:ring-1 focus:ring-ring"/>
          </div>
        </details>
        <div className="text-[10px] text-muted-foreground leading-relaxed">
          اسحب {mode === "bank" ? "الحوالة" : "الفاتورة"} من القائمة وأفلتها فوق خانة {mode === "bank" ? "«بيان البنك»" : "«البيان (الكاشير)»"} بالجدول جنب هاي اللوحة، أو اضغط زر "استبدال".
        </div>
      </div>

      <div className="flex-1 overflow-y-auto bg-white">
        {mode === "bank" ? (
          banks.length === 0 ? (
            <div className="flex flex-col items-center justify-center gap-2 p-10 text-center text-xs text-muted-foreground"><Search className="w-8 h-8 opacity-30"/>لا توجد حوالات تطابق البحث.</div>
          ) : banks.map(b => {
            const diff = Math.abs(b.rawAmount - cashier.matchAmount);
            const exact = diff <= 0.01;
            const sim = nameSim(cashier.name, b.description);
            const isDragged = draggedItemId === b.id;
            return (
              <div key={b.id} draggable onDragStart={() => onDragStart(b.id)} onDragEnd={onDragEnd}
                className={`group flex items-center gap-3 px-4 py-3 border-b border-border last:border-b-0 cursor-grab active:cursor-grabbing hover:bg-blue-50 transition-colors ${isDragged ? "opacity-40" : ""}`}>
                <GripVertical className="w-4 h-4 text-muted-foreground/40 shrink-0"/>
                <div className="flex-1 min-w-0">
                  <div className="whitespace-normal break-words text-sm font-medium leading-6">{b.description}</div>
                  <div className="flex items-center gap-2 mt-1 flex-wrap">
                    <span className="font-mono font-bold text-blue-700 text-sm">{fmtNum(b.rawAmount)}</span>
                    {exact ? <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-green-100 text-green-700 border border-green-300">نفس المبلغ</span>
                      : <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-amber-100 text-amber-700 border border-amber-300">فرق: {fmtNum(diff)}</span>}
                    <span className="text-[10px] text-muted-foreground">تشابه: {(sim*100).toFixed(0)}%</span>
                    {b.accountType && <span className="text-[10px] text-muted-foreground">| {b.accountType}</span>}
                  </div>
                </div>
                <button onClick={() => onReplaceBank(b)} className="px-2.5 py-1.5 bg-blue-600 text-white rounded-lg text-xs font-medium hover:bg-blue-700 transition-colors shrink-0">استبدال</button>
              </div>
            );
          })
        ) : (
          cashiers.length === 0 ? (
            <div className="flex flex-col items-center justify-center gap-2 p-10 text-center text-xs text-muted-foreground"><Search className="w-8 h-8 opacity-30"/>لا توجد فواتير كاشير تطابق البحث.</div>
          ) : cashiers.map(c => {
            const diff = Math.abs(c.matchAmount - oldBank.rawAmount);
            const exact = diff <= 0.01;
            const sim = nameSim(oldBank.description, c.name);
            const isDragged = draggedItemId === c.id;
            return (
              <div key={c.id} draggable onDragStart={() => onDragStart(c.id)} onDragEnd={onDragEnd}
                className={`group flex items-center gap-3 px-4 py-3 border-b border-border last:border-b-0 cursor-grab active:cursor-grabbing hover:bg-emerald-50 transition-colors ${isDragged ? "opacity-40" : ""}`}>
                <GripVertical className="w-4 h-4 text-muted-foreground/40 shrink-0"/>
                <div className="flex-1 min-w-0">
                  <div className="whitespace-normal break-words text-sm font-medium leading-6">{c.name}</div>
                  <div className="flex items-center gap-2 mt-1 flex-wrap">
                    <span className="font-mono font-bold text-emerald-700 text-sm">{fmtNum(c.matchAmount)}</span>
                    {exact ? <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-green-100 text-green-700 border border-green-300">نفس المبلغ</span>
                      : <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-amber-100 text-amber-700 border border-amber-300">فرق: {fmtNum(diff)}</span>}
                    <span className="text-[10px] text-muted-foreground">تشابه: {(sim*100).toFixed(0)}%</span>
                    {c.accountType && <span className="text-[10px] text-muted-foreground">| {c.accountType}</span>}
                  </div>
                </div>
                <button onClick={() => onReplaceCashier(c)} className="px-2.5 py-1.5 bg-emerald-600 text-white rounded-lg text-xs font-medium hover:bg-emerald-700 transition-colors shrink-0">استبدال</button>
              </div>
            );
          })
        )}
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
function ReportPage({
  bank,
  cashier,
  results,
  stageAInvoices = [],
  stageADatabaseSummary = { databaseInvoiceCount: 0, databaseInvoiceTotal: 0, databaseCashTotal: 0, databaseDebtTotal: 0, databaseCardTotal: 0 },
  bankSplits,
  cutoffBatches,
  clearingGroups,
  sourceCashierRows,
  onBack,
}: {
  bank: BankRow[];
  cashier: CashierRow[];
  results: MatchResult[] | null;
  stageAInvoices?: StageAResult[];
  stageADatabaseSummary?: StageASummary;
  bankSplits: BankSplit[];
  cutoffBatches: CutoffBatch[];
  clearingGroups: ClearingGroup[];
  sourceCashierRows: CashierRow[];
  onBack: () => void;
}) {
  const [matchedPage, setMatchedPage] = useState(1);
  const [matchedPageSize, setMatchedPageSize] = useState(10);
  const [stageAInvoicePage, setStageAInvoicePage] = useState(1);
  const [stageAInvoicePageSize, setStageAInvoicePageSize] = useState(10);
  const netTotal = (rows: Array<BankRow | CashierRow>) => rows.reduce((sum, row) => sum + row.debit - row.credit, 0);
  const bankTotal = netTotal(bank);
  const cashierTotal = netTotal(cashier);
  const difference = bankTotal - cashierTotal;
  const reportResults = results ?? [];
  const stageAInvoicePageCount = Math.max(1, Math.ceil(stageAInvoices.length / stageAInvoicePageSize));
  const currentStageAInvoicePage = Math.min(stageAInvoicePage, stageAInvoicePageCount);
  const visibleStageAInvoices = stageAInvoices.slice(
    (currentStageAInvoicePage - 1) * stageAInvoicePageSize,
    currentStageAInvoicePage * stageAInvoicePageSize,
  );
  const matchedRows = reportResults.filter(row => row.type === "saved");
  const reviewRows = reportResults.filter(row => row.type === "pending");
  const unmatchedCashierRows = reportResults.filter(row => row.type === "unmatchedCashier");
  const unmatchedBankRows = reportResults.filter(row => row.type === "unmatchedBank");
  const matched = matchedRows.length;
  const debitTotal = (rows: Array<BankRow | CashierRow>) => rows.reduce((sum, row) => sum + row.debit, 0);
  const creditTotal = (rows: Array<BankRow | CashierRow>) => rows.reduce((sum, row) => sum + row.credit, 0);
  const matchedDetails = matchedChannelRows(reportResults);
  const matchedPageCount = Math.max(1, Math.ceil(matchedDetails.length / matchedPageSize));
  const currentMatchedPage = Math.min(matchedPage, matchedPageCount);
  const visibleMatchedDetails = matchedDetails.slice(
    (currentMatchedPage - 1) * matchedPageSize,
    currentMatchedPage * matchedPageSize,
  );
  const financialIssues = financialIssueRows(reportResults);
  const amountFinancialIssues = financialIssues.filter(row => row.hasAmountDifference);
  const financialExpected = amountFinancialIssues.reduce((sum, row) => sum + row.cashier.matchAmount, 0);
  const financialActual = amountFinancialIssues.reduce((sum, row) => sum + row.bank.rawAmount, 0);
  const matchedBankNet = matchedDetails.reduce((sum, row) => sum + row.bankNet, 0);
  const matchedCashierNet = matchedDetails.reduce((sum, row) => sum + row.cashierNet, 0);
  const beforeMatchDifference = difference;
  const adjustedCashierTotal = cashierTotal - matchedCashierNet + matchedBankNet;
  const afterMatchDifference = bankTotal - adjustedCashierTotal;
  const unmatchedCashierDebit = unmatchedCashierRows.reduce((sum, row) => sum + row.cashier.debit, 0);
  const unmatchedCashierCredit = unmatchedCashierRows.reduce((sum, row) => sum + row.cashier.credit, 0);
  const unmatchedBankDebit = unmatchedBankRows.reduce((sum, row) => sum + row.bank.debit, 0);
  const unmatchedBankCredit = unmatchedBankRows.reduce((sum, row) => sum + row.bank.credit, 0);
  const savedBankIds = new Set(matchedRows.map(row => row.bank.id));
  const savedCashierIds = new Set(matchedRows.map(row => row.cashier.id));
  const reviewBankIds = new Set(reviewRows.map(row => row.bank.id));
  const reviewCashierIds = new Set(reviewRows.map(row => row.cashier.id));
  const visaCashierIds = new Set(reportResults.filter(row => row.type === "visa").map(row => row.cashier.id));
  const jawwalCashierIds = new Set(reportResults.filter(row => row.type === "jawwalPay").map(row => row.cashier.id));
  const channelDefinitions = [
    { label: "بنك فلسطين", matcher: (row: BankRow | CashierRow) => row.accountType === "بنك فلسطين" },
    { label: "بال بي", matcher: (row: BankRow | CashierRow) => row.accountType === "بال بي" },
    { label: "جوال بي", matcher: (row: BankRow | CashierRow) => row.accountType === "جوال بي" },
    {
      label: "فيزا",
      matcher: (row: BankRow | CashierRow) =>
        ("id" in row && visaCashierIds.has(row.id)) ||
        ("accountType" in row && /فيزا|visa/i.test(row.accountType || "")),
    },
  ];
  const channelSummaries = channelDefinitions.map(channel => {
    const channelBank = bank.filter(channel.matcher);
    const channelCashier = cashier.filter(channel.matcher);
    const channelCashierIds = new Set(channelCashier.map(row => row.id));
    const isVisaChannel = channel.label === "فيزا";
    const confirmed = matchedRows.filter(row =>
      isVisaChannel
        ? visaCashierIds.has(row.cashier.id)
        : channelCashierIds.has(row.cashier.id) || channelBank.some(bankRow => bankRow.id === row.bank.id)
    ).length;
    const needsReview = reviewRows.filter(row =>
      isVisaChannel
        ? visaCashierIds.has(row.cashier.id)
        : channelCashierIds.has(row.cashier.id) || channelBank.some(bankRow => bankRow.id === row.bank.id)
    ).length;
    const noTransfer = unmatchedCashierRows.filter(row => channelCashierIds.has(row.cashier.id)).length;
    const noInvoice = unmatchedBankRows.filter(row => channelBank.some(bankRow => bankRow.id === row.bank.id)).length;
    return {
      ...channel,
      bank: channelBank,
      cashier: channelCashier,
      bankTotal: netTotal(channelBank),
      cashierTotal: netTotal(channelCashier),
      difference: netTotal(channelBank) - netTotal(channelCashier),
      debitBank: debitTotal(channelBank),
      creditBank: creditTotal(channelBank),
      debitCashier: debitTotal(channelCashier),
      creditCashier: creditTotal(channelCashier),
      confirmed,
      needsReview,
      noTransfer,
      noInvoice,
    };
  });
  const totals = {
    debitBank: debitTotal(bank),
    creditBank: creditTotal(bank),
    debitCashier: debitTotal(cashier),
    creditCashier: creditTotal(cashier),
  };
  const categories = [
    { label: "المطابقات المؤكدة", value: matched, amount: matchedRows.reduce((sum, row) => sum + row.bank.rawAmount, 0), color: "#fbb313" },
    { label: "مطابقات تحتاج مراجعة", value: reviewRows.length, amount: reviewRows.reduce((sum, row) => sum + row.bank.rawAmount, 0), color: "#1554ad" },
    { label: "فواتير بلا حوالات · مدين", value: unmatchedCashierRows.filter(row => row.cashier.debit > 0).length, amount: unmatchedCashierDebit, color: "#df2828" },
    { label: "فواتير بلا حوالات · دائن", value: unmatchedCashierRows.filter(row => row.cashier.credit > 0).length, amount: unmatchedCashierCredit, color: "#ef6c52" },
    { label: "حوالات بلا فواتير · مدين", value: unmatchedBankRows.filter(row => row.bank.debit > 0).length, amount: unmatchedBankDebit, color: "#9250c8" },
    { label: "حوالات بلا فواتير · دائن", value: unmatchedBankRows.filter(row => row.bank.credit > 0).length, amount: unmatchedBankCredit, color: "#b477d1" },
    { label: "فيزا", value: reportResults.filter(row => row.type === "visa").length, amount: reportResults.filter(row => row.type === "visa").reduce((sum, row) => sum + row.cashier.matchAmount, 0), color: "#19a8bf" },
    { label: "محفظة شخصية", value: reportResults.filter(row => row.type === "mahmoudWallet").length, amount: reportResults.filter(row => row.type === "mahmoudWallet").reduce((sum, row) => sum + row.cashier.matchAmount, 0), color: "#f58a12" },
    { label: "جوال بي", value: reportResults.filter(row => row.type === "jawwalPay").length, amount: reportResults.filter(row => row.type === "jawwalPay").reduce((sum, row) => sum + row.cashier.matchAmount, 0), color: "#19a8bf" },
    { label: "فواتير A مطابقة مع قاعدة البيانات", value: stageAInvoices.filter(invoice => invoice.matched).length, amount: stageAInvoices.filter(invoice => invoice.matched).reduce((sum, invoice) => sum + invoice.amount, 0), color: "#16a34a" },
    { label: "فواتير A غير مطابقة مع قاعدة البيانات", value: stageAInvoices.filter(invoice => !invoice.matched).length, amount: stageAInvoices.filter(invoice => !invoice.matched).reduce((sum, invoice) => sum + invoice.amount, 0), color: "#dc2626" },
  ];
  const maxAmount = Math.max(...categories.map(category => category.amount), 1);
  return (
    <div dir="rtl" className="report-page min-h-[100dvh] bg-background p-4 sm:p-8">
      <div className="mx-auto max-w-6xl space-y-6">
        <header className="flex flex-wrap items-end justify-between gap-4">
          <div>
            <div className="mb-2 flex items-center gap-2 text-xs font-semibold text-primary"><BarChart3 className="h-4 w-4" /> مركز التقارير</div>
            <h1 className="text-3xl font-bold tracking-tight">تقرير التسوية</h1>
            <p className="mt-2 text-sm text-muted-foreground">صورة واحدة لما سُجّل، وما بقي يحتاج قراراً.</p>
          </div>
          <div className="flex flex-wrap items-center gap-2 report-actions">
            <button onClick={() => window.print()} className="flex items-center gap-2 rounded-xl border border-primary/30 bg-primary/5 px-4 py-2.5 text-sm font-semibold text-primary hover:bg-primary/10">
              <Download className="h-4 w-4" /> تنزيل PDF
            </button>
            <button onClick={onBack} className="flex items-center gap-2 rounded-xl border bg-card px-4 py-2.5 text-sm font-semibold hover:bg-muted"><ArrowUpRight className="h-4 w-4 rotate-180" /> العودة إلى مساحة العمل</button>
          </div>
        </header>
        {!bank.length && !cashier.length && !clearingGroups.length && !stageAInvoices.length && !stageADatabaseSummary.databaseInvoiceCount ? (
          <div className="rounded-2xl border border-dashed bg-card p-12 text-center">
            <BarChart3 className="mx-auto mb-4 h-10 w-10 text-muted-foreground/50" />
            <h2 className="font-bold">لا يوجد كشف محفوظ بعد</h2>
            <p className="mt-2 text-sm text-muted-foreground">ارفع كشفي الكاشير والبنك ثم شغّل المطابقة لعرض التقرير.</p>
            <button onClick={onBack} className="mt-6 rounded-xl bg-primary px-5 py-2.5 text-sm font-bold text-primary-foreground">بدء جلسة جديدة</button>
          </div>
        ) : (
          <>
            <div className="grid gap-4 md:grid-cols-3">
              <div className="rounded-2xl border bg-card p-5"><div className="text-xs font-semibold text-muted-foreground">إجمالي الكاشير</div><div className="mt-3 font-mono text-2xl font-bold">{fmtNum(cashierTotal)} <span className="text-xs font-normal text-muted-foreground">₪</span></div><div className="mt-3 h-1 rounded-full bg-primary/15"><div className="h-full w-full rounded-full bg-primary" /></div></div>
              <div className="rounded-2xl border bg-card p-5"><div className="text-xs font-semibold text-muted-foreground">إجمالي الحوالات البنكية</div><div className="mt-3 font-mono text-2xl font-bold">{fmtNum(bankTotal)} <span className="text-xs font-normal text-muted-foreground">₪</span></div><div className="mt-3 h-1 rounded-full bg-[hsl(43_96%_52%/.18)]"><div className="h-full w-[88%] rounded-full bg-[hsl(43_96%_52%)]" /></div></div>
              <div className={`rounded-2xl border p-5 ${Math.abs(difference) < .01 ? "border-emerald-200 bg-emerald-50/60" : "border-amber-200 bg-amber-50/70"}`}><div className="text-xs font-semibold text-muted-foreground">الفرق بين السجلين</div><div className="mt-3 font-mono text-2xl font-bold">{difference > 0 ? "+" : ""}{fmtNum(difference)} <span className="text-xs font-normal text-muted-foreground">₪</span></div><div className="mt-3 text-xs font-semibold">{Math.abs(difference) < .01 ? "السجلان متوازنان" : "يحتاج إلى مراجعة الفروقات"}</div></div>
            </div>
            {!!stageADatabaseSummary.databaseInvoiceCount && <section className="space-y-4 rounded-2xl border border-blue-200 bg-blue-50/40 p-5 sm:p-6">
              <div><h2 className="font-bold">ملخص فواتير قاعدة البيانات — المرحلة A</h2><p className="mt-1 text-xs text-muted-foreground">إجمالي فواتير الكاشير والمدفوعات وقيمة الخصومات، مع حالات الفواتير المرحّلة إلى المطابقة.</p></div>
              <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-6">
                {[
                  ["عدد الفواتير", stageADatabaseSummary.databaseInvoiceCount],
                  ["إجمالي الفواتير", stageADatabaseSummary.databaseInvoiceTotal],
                  ["النقدي", stageADatabaseSummary.databaseCashTotal],
                  ["مدفوع بالكارد / التطبيقات", stageADatabaseSummary.databaseCardTotal],
                  ["المتبقي على الزبائن", stageADatabaseSummary.databaseDebtTotal],
                  ["الخصم", stageADatabaseSummary.databaseDiscountTotal ?? 0],
                ].map(([label, value]) => <div key={String(label)} className="rounded-xl border bg-white p-3"><div className="text-xs text-muted-foreground">{label}</div><div className="mt-1 font-mono text-sm font-bold">{label === "عدد الفواتير" ? value : `${fmtNum(Number(value))} ₪`}</div></div>)}
              </div>
              {!!stageAInvoices.length && <div className="space-y-3">
                <div className="flex flex-wrap items-center justify-between gap-2 text-xs">
                  <span className="text-muted-foreground">عرض فواتير قاعدة البيانات كسجلات واضحة، {stageAInvoices.length} فاتورة إجمالاً.</span>
                  <label className="flex items-center gap-2 text-muted-foreground">عدد السطور
                    <select value={stageAInvoicePageSize} onChange={event => { setStageAInvoicePageSize(Number(event.target.value)); setStageAInvoicePage(1); }} className="rounded-lg border bg-white px-2 py-1.5 text-foreground">
                      <option value={10}>10</option><option value={25}>25</option><option value={50}>50</option>
                    </select>
                  </label>
                </div>
                <div className="space-y-2">
                  {visibleStageAInvoices.map(invoice => <article key={`${invoice.source}:${invoice.id}`} className="rounded-xl border border-slate-200 bg-white p-3 sm:p-4">
                    <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 pb-2">
                      <div className="flex min-w-0 flex-wrap items-center gap-x-3 gap-y-1">
                        <strong className="text-sm">{invoice.databaseName || invoice.name || invoice.originalName}</strong>
                        <span className="text-xs text-muted-foreground">فاتورة #{invoice.invoiceNumber || "—"}</span>
                        <span className="text-xs text-muted-foreground">{invoice.category || invoice.source}</span>
                      </div>
                      <span className={`rounded-full px-2.5 py-1 text-xs font-semibold ${invoice.matched ? "bg-emerald-50 text-emerald-700" : "bg-red-50 text-red-700"}`}>{invoice.matched ? "مطابق" : "غير مطابق"}</span>
                    </div>
                    <div className="grid grid-cols-2 gap-x-4 gap-y-2 py-3 text-xs sm:grid-cols-4">
                      {[["رقم الزبون", invoice.customerNumber], ["رقم المستخدم", invoice.userId], ["التاريخ", invoice.invoiceDate], ["الوقت", invoice.registrationTime]].map(([label, value]) => <div key={label}><span className="text-muted-foreground">{label}: </span><span className="font-medium">{value || "—"}</span></div>)}
                    </div>
                    <div className="grid grid-cols-2 gap-2 sm:grid-cols-3 lg:grid-cols-6">
                      {[["الإجمالي", invoice.invoiceTotal, "text-slate-900"], ["نقدي", invoice.cashAmount, "text-slate-900"], ["كارد / تطبيق", invoice.databaseAmount, "text-blue-700"], ["المتبقي", invoice.debtAmount, "text-slate-900"], ["الخصم", invoice.discountAmount, invoice.discountAmount > 0 ? "text-amber-700" : "text-slate-900"]].map(([label, value, tone]) => <div key={String(label)} className="rounded-lg bg-slate-50 px-2.5 py-2"><div className="text-[10px] text-muted-foreground">{label}</div><div className={`mt-1 font-mono text-xs font-semibold ${tone}`}>{fmtNum(Number(value))} ₪</div></div>)}
                    </div>
                  </article>)}
                </div>
                {stageAInvoicePageCount > 1 && <div className="flex flex-wrap items-center justify-center gap-3 rounded-xl border bg-white p-3 text-xs">
                  <button type="button" onClick={() => setStageAInvoicePage(page => Math.max(1, page - 1))} disabled={currentStageAInvoicePage <= 1} className="rounded-lg border px-3 py-1.5 disabled:opacity-40">السابق</button>
                  <span className="text-muted-foreground">صفحة {currentStageAInvoicePage} من {stageAInvoicePageCount} · عرض {visibleStageAInvoices.length} من {stageAInvoices.length}</span>
                  <button type="button" onClick={() => setStageAInvoicePage(page => Math.min(stageAInvoicePageCount, page + 1))} disabled={currentStageAInvoicePage >= stageAInvoicePageCount} className="rounded-lg border px-3 py-1.5 disabled:opacity-40">التالي</button>
                </div>}
              </div>}
            </section>}
            <section className="rounded-2xl border bg-card p-5 sm:p-6">
              <div className="mb-4">
                <h2 className="font-bold">تفصيل الدائن والمدين</h2>
                <p className="mt-1 text-xs text-muted-foreground">المدين يُجمع، والدائن يُطرح منه لحساب الصافي. وتبقى قيم المدين والدائن معروضة بشكل مستقل.</p>
              </div>
              <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                {[
                  { label: "الكاشير · مدين", value: totals.debitCashier, tone: "text-emerald-700" },
                  { label: "الكاشير · دائن", value: totals.creditCashier, tone: "text-red-600" },
                  { label: "الحوالات · مدين", value: totals.debitBank, tone: "text-blue-700" },
                  { label: "الحوالات · دائن", value: totals.creditBank, tone: "text-orange-600" },
                ].map(item => (
                  <div key={item.label} className="rounded-xl border bg-muted/20 p-4">
                    <div className="text-xs font-semibold text-muted-foreground">{item.label}</div>
                    <div className={`mt-2 font-mono text-xl font-bold ${item.tone}`}>{fmtNum(item.value)} <span className="text-xs font-normal">₪</span></div>
                  </div>
                ))}
              </div>
            </section>
            {(bankSplits.length > 0 || cutoffBatches.length > 0) && (
              <section className="space-y-5 rounded-2xl border border-violet-200 bg-violet-50/40 p-5 sm:p-6">
                <div>
                  <h2 className="font-bold">تفصيل أدوات التجزئة وتجميع فترة التعطل</h2>
                  <p className="mt-1 text-xs text-muted-foreground">سجل تدقيقي مساعد يوضح التوزيعات والمجاميع دون إنشاء قيود في النظام المحاسبي.</p>
                </div>
                {bankSplits.length > 0 && <div className="space-y-2">
                  <h3 className="text-sm font-semibold">تجزئة الحوالات</h3>
                  <div className="overflow-x-auto rounded-xl border bg-card">
                    <table className="w-full min-w-[780px] text-xs">
                      <thead className="bg-muted"><tr>
                        <th className="px-3 py-2 text-right">الحوالة</th><th className="px-3 py-2 text-right">قيمة الأصل</th><th className="px-3 py-2 text-right">التخصيص</th><th className="px-3 py-2 text-right">المبلغ</th><th className="px-3 py-2 text-right">التوضيح</th>
                      </tr></thead>
                      <tbody>{bankSplits.flatMap(split => split.allocations.map((allocation, index) => {
                        const cashierRow = allocation.kind === "company"
                          ? sourceCashierRows.find(row => row.id === allocation.cashierId && (row._fileSessionId ?? 0) === (allocation.cashierSessionId ?? 0))
                          : null;
                        return <tr key={`${split.id}:${allocation.id}`} className="border-t">
                          <td className="px-3 py-2">{split.date} · {split.description} · حصة {index + 1}</td>
                          <td className="px-3 py-2 font-mono">{fmtNum(split.originalAmount)}</td>
                          <td className="px-3 py-2">{allocation.kind === "company" ? `شركتنا${cashierRow ? ` · ${cashierRow.name}` : ""}` : "جهة خارجية"}</td>
                          <td className="px-3 py-2 font-mono">{fmtNum(allocation.amount)}</td>
                          <td className="px-3 py-2">{allocation.note || "—"}</td>
                        </tr>;
                      }) )}</tbody>
                    </table>
                  </div>
                </div>}
                {cutoffBatches.length > 0 && <div className="space-y-2">
                  <h3 className="text-sm font-semibold">تجميع فترة التعطل</h3>
                  <div className="overflow-x-auto rounded-xl border bg-card">
                    <table className="w-full min-w-[850px] text-xs">
                      <thead className="bg-muted"><tr>
                        <th className="px-3 py-2 text-right">الفترة</th><th className="px-3 py-2 text-right">سند الإجمالي</th><th className="px-3 py-2 text-right">الفواتير الحقيقية</th><th className="px-3 py-2 text-right">عددها</th><th className="px-3 py-2 text-right">نافذة التداخل</th><th className="px-3 py-2 text-right">ملاحظة</th>
                      </tr></thead>
                      <tbody>{cutoffBatches.map(batch => {
                        const aggregate = sourceCashierRows.find(row => row.id === batch.aggregateId && (row._fileSessionId ?? 0) === batch.aggregateSessionId);
                        const sources = batch.sourceRows.map(reference => sourceCashierRows.find(row => row.id === reference.id && (row._fileSessionId ?? 0) === reference.sessionId)).filter((row): row is CashierRow => !!row);
                        return <tr key={batch.id} className="border-t align-top">
                          <td className="px-3 py-2">{batch.dateFrom} إلى {batch.dateTo}</td>
                          <td className="px-3 py-2">{aggregate?.name || "سند إجمالي"} · {fmtNum(batch.total)}</td>
                          <td className="px-3 py-2">{sources.map(row => `${row.date} · ${row.name} (${fmtNum(row.matchAmount)})`).join("، ")}</td>
                          <td className="px-3 py-2">{sources.length}</td>
                          <td className="px-3 py-2">{batch.graceDays ? `±${batch.graceDays} يوم` : "الفترة المحددة"}</td>
                          <td className="px-3 py-2">{batch.note || "—"}</td>
                        </tr>;
                      })}</tbody>
                    </table>
                  </div>
                </div>}
              </section>
            )}
            <section className="report-channel-section report-page-break-after rounded-2xl border bg-card p-5 sm:p-6">
              <div className="mb-4">
                <h2 className="font-bold">تفصيل القنوات المالية</h2>
                <p className="mt-1 text-xs text-muted-foreground">لكل قناة: إجمالي الكاشير والبنك، الفرق، الدائن والمدين، وحالة المطابقات.</p>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full min-w-[980px] text-xs">
                  <thead className="bg-muted/60 text-muted-foreground">
                    <tr>
                      <th className="px-3 py-3 text-right font-semibold">القناة</th>
                      <th className="px-3 py-3 text-right font-semibold">الكاشير</th>
                      <th className="px-3 py-3 text-right font-semibold">البنك</th>
                      <th className="px-3 py-3 text-right font-semibold">الفرق</th>
                      <th className="px-3 py-3 text-right font-semibold">كاشير: مدين / دائن</th>
                      <th className="px-3 py-3 text-right font-semibold">بنك: مدين / دائن</th>
                      <th className="px-3 py-3 text-right font-semibold">مؤكدة</th>
                      <th className="px-3 py-3 text-right font-semibold">تحتاج مراجعة</th>
                      <th className="px-3 py-3 text-right font-semibold">بلا حوالات</th>
                      <th className="px-3 py-3 text-right font-semibold">بلا فواتير</th>
                    </tr>
                  </thead>
                  <tbody>
                    {channelSummaries.map(channel => (
                      <tr key={channel.label} className="border-t border-border">
                        <td className="px-3 py-3 font-bold">{channel.label}</td>
                        <td className="px-3 py-3 font-mono text-emerald-700">{fmtNum(channel.cashierTotal)}</td>
                        <td className="px-3 py-3 font-mono text-blue-700">{fmtNum(channel.bankTotal)}</td>
                        <td className={`px-3 py-3 font-mono font-bold ${Math.abs(channel.difference) < .01 ? "text-emerald-700" : "text-amber-700"}`}>{fmtNum(channel.difference)}</td>
                        <td className="px-3 py-3 font-mono">{fmtNum(channel.debitCashier)} / {fmtNum(channel.creditCashier)}</td>
                        <td className="px-3 py-3 font-mono">{fmtNum(channel.debitBank)} / {fmtNum(channel.creditBank)}</td>
                        <td className="px-3 py-3 font-semibold text-emerald-700">{channel.confirmed}</td>
                        <td className="px-3 py-3 font-semibold text-amber-700">{channel.needsReview}</td>
                        <td className="px-3 py-3 font-semibold text-red-600">{channel.noTransfer}</td>
                        <td className="px-3 py-3 font-semibold text-violet-700">{channel.noInvoice}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
            <section className="rounded-2xl border border-blue-100 bg-blue-50/60 p-5 text-sm text-blue-950">
              <h2 className="font-bold">طريقة قراءة التقرير</h2>
              <div className="mt-3 grid gap-2 text-xs leading-6 text-blue-900 sm:grid-cols-2">
                <p><strong>الدائن والمدين:</strong> القيم مأخوذة من الأعمدة الأصلية بعد تطبيق خيار عكس المدين والدائن إن كان مفعّلاً؛ لا نحتاج إلى كتابة إشارة سالبة للدائن.</p>
                <p><strong>الصافي:</strong> المدين ناقص الدائن. مثال: 50 مدين و30 دائن = 20 صافي، وليس 80.</p>
                <p><strong>الفرق:</strong> صافي البنك ناقص صافي الكاشير؛ الرقم الموجب يعني أن صافي الحوالات أعلى.</p>
                <p><strong>تحتاج مراجعة:</strong> مطابقات مقترحة لم تُعتمد بعد ويمكن مراجعتها من مساعد المطابقة.</p>
                <p><strong>بلا حوالات / بلا فواتير:</strong> فواتير لا يقابلها تحويل، أو تحويلات لا تقابلها فاتورة في الكشف الآخر.</p>
              </div>
            </section>
            <div className="report-page-break-after grid gap-6 lg:grid-cols-[1.35fr_.65fr]">
              <section className="rounded-2xl border bg-card p-5 sm:p-6">
                <div className="mb-6 flex items-center justify-between"><div><h2 className="font-bold">توزيع الفروقات والنتائج</h2><p className="mt-1 text-xs text-muted-foreground">القيمة المالية لكل فئة في الجلسة الحالية</p></div><span className="rounded-full bg-muted px-3 py-1 text-xs">{categories.reduce((sum, item) => sum + item.value, 0)} بند</span></div>
                <div className="space-y-5">
                  {categories.map(category => <div key={category.label}><div className="mb-2 flex items-center justify-between gap-3 text-sm"><span className="font-semibold">{category.label}</span><span className="font-mono text-xs text-muted-foreground">{fmtNum(category.amount)} ₪ · {category.value}</span></div><div className="h-3 overflow-hidden rounded-full bg-muted"><div className="h-full rounded-full" style={{ width: `${Math.max(category.value ? 4 : 0, category.amount / maxAmount * 100)}%`, backgroundColor: category.color }} /></div></div>)}
                </div>
              </section>
              <section className="rounded-2xl border bg-[hsl(216_48%_15%)] p-6 text-white">
                <div className="text-xs font-semibold text-white/55">ملخص القرار</div><h2 className="mt-3 text-2xl font-bold leading-relaxed">كل رقم له مكان واضح.</h2>
                <p className="mt-3 text-sm leading-7 text-white/65">راجع العناصر المعلّقة أولاً، ثم صدّر ملف التسوية ليبقى أثر القرار مع الفريق.</p>
                <div className="mt-8 border-t border-white/10 pt-5"><div className="flex justify-between text-xs text-white/55"><span>نسبة المطابقة المؤكدة</span><span className="font-mono text-[hsl(43_96%_52%)]">{cashier.length ? Math.round(matched / cashier.length * 100) : 0}%</span></div><div className="mt-3 h-2 rounded-full bg-white/10"><div className="h-full rounded-full bg-[hsl(43_96%_52%)]" style={{ width: `${cashier.length ? matched / cashier.length * 100 : 0}%` }} /></div></div>
              </section>
            </div>
            <section className="report-page-three rounded-2xl border bg-card p-5 sm:p-6">
              <div className="mb-8 border-b border-border pb-6">
                <div className="mb-4">
                  <h2 className="font-bold">الفواتير التي فيها مشاكل مالية</h2>
                  <p className="mt-1 text-xs text-muted-foreground">تظهر هنا المطابقات التي تحتاج مراجعة بسبب اختلاف المبلغ أو نوع الحساب أو التاريخ أو وجود تعديل يدوي؛ اختلاف رقم الجوال يُعرض في ملاحظات المطابقة وليس كفرق مالي.</p>
                </div>
                <div className="mb-4 grid gap-3 sm:grid-cols-4">
                  <div className="rounded-xl border bg-muted/20 p-4"><div className="text-xs font-semibold text-muted-foreground">فواتير تحتاج تعديل</div><div className="mt-2 font-mono text-xl font-bold text-amber-700">{financialIssues.length}</div></div>
                  <div className="rounded-xl border bg-muted/20 p-4"><div className="text-xs font-semibold text-muted-foreground">إجمالي الفواتير للمقارنة</div><div className="mt-2 font-mono text-xl font-bold text-blue-700">{fmtNum(financialExpected)} ₪</div></div>
                  <div className="rounded-xl border bg-muted/20 p-4"><div className="text-xs font-semibold text-muted-foreground">إجمالي الواصل الفعلي بالبنك</div><div className="mt-2 font-mono text-xl font-bold text-emerald-700">{fmtNum(financialActual)} ₪</div></div>
                  <div className="rounded-xl border border-red-200 bg-red-50/60 p-4"><div className="text-xs font-semibold text-muted-foreground">فرق الفاتورة عن البنك</div><div className="mt-2 font-mono text-xl font-bold text-red-600">{fmtNum(financialExpected - financialActual)} ₪</div></div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full min-w-[1250px] text-xs">
                    <thead className="bg-red-50 text-muted-foreground"><tr>
                      <th className="px-3 py-3 text-right">القناة</th><th className="px-3 py-3 text-right">الفاتورة</th><th className="px-3 py-3 text-right">الحوالة</th><th className="px-3 py-3 text-right">مبلغ الفاتورة</th><th className="px-3 py-3 text-right">الواصل بالبنك</th><th className="px-3 py-3 text-right">فرق المبلغ</th><th className="px-3 py-3 text-right">نوع المشكلة / الحالة</th><th className="px-3 py-3 text-right">الفاتورة: مدين / دائن</th><th className="px-3 py-3 text-right">البنك: مدين / دائن</th><th className="px-3 py-3 text-right">تاريخ الفاتورة</th><th className="px-3 py-3 text-right">تاريخ الحوالة</th>
                    </tr></thead>
                    <tbody>{financialIssues.map(row => <tr key={`financial-top-${row.savedMatch.id}`} className="border-t border-border">
                      <td className="px-3 py-3 font-bold">{row.channel}</td>
                      <td className="px-3 py-3">{row.cashier.name}</td>
                      <td className="px-3 py-3">{row.bank.description}</td>
                      <td className="px-3 py-3 font-mono text-blue-700">{fmtNum(row.cashier.matchAmount)}</td>
                      <td className="px-3 py-3 font-mono text-emerald-700">{fmtNum(row.bank.rawAmount)}</td>
                      <td className={`px-3 py-3 font-mono font-bold ${row.hasAmountDifference ? "text-red-600" : "text-muted-foreground"}`}>{fmtNum(row.cashier.matchAmount - row.bank.rawAmount)}</td>
                      <td className="px-3 py-3 font-semibold text-amber-800">
                        {row.manualEditIssue || [
                          row.hasAmountDifference ? "اختلاف مبلغ — البنك هو الصحيح" : "",
                          row.hasAccountTypeDifference ? `اختلاف نوع الحساب: ${row.cashier.accountType || "—"} ↔ ${row.bank.accountType || "—"}` : "",
                          row.hasDateDifference ? `اختلاف التاريخ: الفاتورة ${row.cashier.date} ↔ الحوالة ${row.bank.date}` : "",
                        ].filter(Boolean).join(" · ")}
                      </td>
                      <td className="px-3 py-3 font-mono">{fmtNum(row.cashier.debit)} / {fmtNum(row.cashier.credit)}</td>
                      <td className="px-3 py-3 font-mono">{fmtNum(row.bank.debit)} / {fmtNum(row.bank.credit)}</td>
                      <td className="px-3 py-3">{row.cashier.date || "—"}</td>
                      <td className="px-3 py-3">{row.bank.date || "—"}</td>
                    </tr>)}</tbody>
                  </table>
                  {!financialIssues.length && <p className="py-5 text-center text-xs text-emerald-700">لا توجد مطابقات مؤكدة تحتاج تعديلًا.</p>}
                </div>
              </div>
              <div className="mb-4">
                <h2 className="font-bold">الفرق قبل المطابقة وبعدها</h2>
                <p className="mt-1 text-xs text-muted-foreground">الفرق محسوب من صافي المدين ناقص الدائن، مع عزل المطابقات المؤكدة فقط.</p>
              </div>
              <div className="mb-6 grid gap-3 sm:grid-cols-2">
                <div className="rounded-xl border border-amber-200 bg-amber-50/60 p-4">
                  <div className="text-xs font-semibold text-muted-foreground">الفرق قبل المطابقة</div>
                  <div className="mt-2 font-mono text-2xl font-bold text-amber-700">{fmtNum(beforeMatchDifference)} ₪</div>
                </div>
                <div className="rounded-xl border border-emerald-200 bg-emerald-50/60 p-4">
                  <div className="text-xs font-semibold text-muted-foreground">الفرق بعد المطابقة المؤكدة</div>
                  <div className="mt-2 font-mono text-2xl font-bold text-emerald-700">{fmtNum(afterMatchDifference)} ₪</div>
                </div>
              </div>
              <h3 className="mb-3 font-bold">إجماليات المطابقات حسب القناة</h3>
              <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-3">
                {Array.from(new Set(matchedDetails.map(row => row.channel))).map(channel => {
                  const rows = matchedDetails.filter(row => row.channel === channel);
                  const sum = (pick: (row: typeof rows[number]) => number) => rows.reduce((total, row) => total + pick(row), 0);
                  return <article key={channel} className="rounded-xl border border-slate-200 bg-white p-4">
                    <div className="flex items-center justify-between gap-2 border-b border-slate-100 pb-2">
                      <h4 className="font-bold">{channel}</h4><span className="rounded-full bg-blue-50 px-2.5 py-1 text-xs font-semibold text-blue-700">{rows.length} مطابقة</span>
                    </div>
                    <div className="mt-3 grid grid-cols-2 gap-2 text-xs">
                      {[["البنك · مدين", sum(row => row.bank.debit)], ["البنك · دائن", sum(row => row.bank.credit)], ["الفواتير · مدين", sum(row => row.cashier.debit)], ["الفواتير · دائن", sum(row => row.cashier.credit)], ["فرق القناة", sum(row => row.bankNet - row.cashierNet)]].map(([label, value]) => <div key={String(label)} className="rounded-lg bg-slate-50 p-2">
                        <div className="text-[10px] text-muted-foreground">{label}</div><div className="mt-1 font-mono font-semibold">{fmtNum(Number(value))} ₪</div>
                      </div>)}
                    </div>
                  </article>;
                })}
                {!matchedDetails.length && <p className="rounded-lg border bg-white p-4 text-center text-xs text-muted-foreground">لا توجد مطابقات مؤكدة لعرضها.</p>}
              </div>
              <div className="mb-3 mt-6 flex flex-wrap items-center justify-between gap-3">
                <div><h3 className="font-bold">تفصيل كل مطابقة مؤكدة</h3><p className="mt-1 text-xs text-muted-foreground">{matchedDetails.length} مطابقة محفوظة، معروضة كسجلات مستقلة لسهولة القراءة.</p></div>
                {!!matchedDetails.length && <label className="flex items-center gap-2 text-xs text-muted-foreground">عدد السطور
                  <select value={matchedPageSize} onChange={event => { setMatchedPageSize(Number(event.target.value)); setMatchedPage(1); }} className="rounded-lg border bg-white px-2 py-1.5 text-foreground">
                    <option value={5}>5</option><option value={10}>10</option><option value={25}>25</option>
                  </select>
                </label>}
              </div>
              <div className="space-y-3">
                {visibleMatchedDetails.map(row => <article key={row.savedMatch.id} className="rounded-xl border border-slate-200 bg-white p-4">
                  <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 pb-3">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="rounded-full bg-blue-50 px-2.5 py-1 text-xs font-bold text-blue-700">{row.channel}</span>
                      <span className="rounded-full bg-emerald-50 px-2.5 py-1 text-xs font-semibold text-emerald-700">مؤكدة</span>
                      <span className="text-[11px] text-muted-foreground">{row.savedMatch.isManual ? "مطابقة يدوية" : "مطابقة تلقائية"}</span>
                    </div>
                    <div className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${Math.abs(row.bankNet - row.cashierNet) < 0.01 ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-800"}`}>
                      الفرق: {fmtNum(row.bankNet - row.cashierNet)} ₪
                    </div>
                  </div>
                  <div className="mt-3 grid gap-3 lg:grid-cols-2">
                    <div className="min-w-0 rounded-lg border border-slate-100 p-3">
                      <div className="mb-1 text-[10px] font-semibold text-muted-foreground">فاتورة الكاشير</div>
                      <div className="break-words text-sm font-semibold">{row.cashier.name || "—"}</div>
                      <div className="mt-1 text-[11px] text-muted-foreground">رقم المستخدم: {row.cashier.userId || "—"} · التاريخ: {row.cashier.date || "—"}{row.cashier.registrationTime ? ` · الوقت: ${row.cashier.registrationTime}` : ""}</div>
                    </div>
                    <div className="min-w-0 rounded-lg border border-slate-100 p-3">
                      <div className="mb-1 text-[10px] font-semibold text-muted-foreground">حوالة البنك</div>
                      <div className="break-words text-sm font-semibold">{row.bank.description || "—"}</div>
                      <div className="mt-1 text-[11px] text-muted-foreground">التاريخ: {row.bank.date || "—"} · نوع الحساب: {row.bank.accountType || row.channel || "—"}</div>
                    </div>
                  </div>
                  <div className="mt-3 grid grid-cols-2 gap-2 sm:grid-cols-4">
                    {[["الكاشير · مدين", row.cashier.debit], ["الكاشير · دائن", row.cashier.credit], ["البنك · مدين", row.bank.debit], ["البنك · دائن", row.bank.credit]].map(([label, value]) => <div key={String(label)} className="rounded-lg bg-slate-50 px-3 py-2">
                      <div className="text-[10px] text-muted-foreground">{label}</div><div className="mt-1 font-mono text-xs font-semibold">{fmtNum(Number(value))} ₪</div>
                    </div>)}
                  </div>
                </article>)}
                {!matchedDetails.length && <p className="rounded-lg border bg-white p-4 text-center text-xs text-muted-foreground">لا توجد مطابقات مؤكدة لعرضها.</p>}
              </div>
              {matchedDetails.length > matchedPageSize && <div className="mt-3 flex flex-wrap items-center justify-center gap-3 rounded-xl border bg-white p-3 text-xs">
                <button type="button" onClick={() => setMatchedPage(page => Math.max(1, page - 1))} disabled={currentMatchedPage <= 1} className="rounded-lg border px-3 py-1.5 disabled:opacity-40">السابق</button>
                <span className="text-muted-foreground">صفحة {currentMatchedPage} من {matchedPageCount} · عرض {visibleMatchedDetails.length} من {matchedDetails.length}</span>
                <button type="button" onClick={() => setMatchedPage(page => Math.min(matchedPageCount, page + 1))} disabled={currentMatchedPage >= matchedPageCount} className="rounded-lg border px-3 py-1.5 disabled:opacity-40">التالي</button>
              </div>}
            </section>
            {clearingGroups.length > 0 && <section className="rounded-2xl border bg-card p-5 sm:p-6">
              <h2 className="font-bold">الحركات المسكّرة يدوياً ({clearingGroups.length} مجموعة)</h2>
              <p className="mt-1 text-xs text-muted-foreground">هذه الحركات عُزلت من العناصر المفتوحة بعد التحقق من تساوي المدين والدائن.</p>
              <div className="mt-4 space-y-3">
                {clearingGroups.map(group => {
                  const getRows = (references: CutoffRowRef[]) => references.map(reference =>
                    sourceCashierRows.find(row => row.id === reference.id && (row._fileSessionId ?? 0) === reference.sessionId)
                  );
                  const debits = getRows(group.debitRows);
                  const credits = getRows(group.creditRows);
                  return <article key={group.id} className="rounded-xl border p-3">
                    <div className="flex flex-wrap justify-between gap-2 text-sm font-semibold"><span>مجموعة {group.id}</span><span className="font-mono">{fmtNum(group.total)} ₪</span></div>
                    <div className="mt-2 grid gap-3 md:grid-cols-2">
                      {[{ label: "مدين", rows: debits }, { label: "دائن", rows: credits }].map(side => <div key={side.label} className="space-y-1">
                        <p className="text-xs font-semibold">{side.label}</p>
                        {side.rows.map((row, index) => <p key={`${group.id}-${side.label}-${index}`} className="flex justify-between gap-2 rounded bg-muted/40 px-2 py-1.5 text-xs"><span>{row ? `${row.date} · ${row.name}` : "حركة غير متاحة في المصدر الحالي"}</span><b className="font-mono">{row ? fmtNum(side.label === "مدين" ? row.debit : row.credit) : "—"}</b></p>)}
                      </div>)}
                    </div>
                    {group.note && <p className="mt-2 text-xs text-muted-foreground">{group.note}</p>}
                  </article>;
                })}
              </div>
            </section>}
          </>
        )}
      </div>
    </div>
  );
}

function PlatformLoading({ message = "جارٍ تحميل منصة الملفات..." }: { message?: string }) {
  return <div dir="rtl" className="flex min-h-[60vh] flex-col items-center justify-center gap-4 bg-slate-50 px-5 text-center">
    <div className="flex h-20 w-20 items-center justify-center rounded-2xl border border-amber-200 bg-white p-2 shadow-sm">
      <img src="/taswiyah-logo.png" alt="شعار تسوية هب" className="h-full w-full object-contain" />
    </div>
    <div>
      <p className="text-base font-bold text-slate-900">{message}</p>
      <p className="mt-1 text-sm font-medium text-slate-700">يرجى الانتظار حتى تكتمل العملية.</p>
    </div>
    <div className="h-1.5 w-36 overflow-hidden rounded-full bg-slate-200">
      <div className="h-full w-1/2 animate-pulse rounded-full bg-amber-500" />
    </div>
  </div>;
}

export default function App({ initialPage = "recon2", onSafeActivities }: { initialPage?: PageId; onSafeActivities?: (activities: string[]) => void } = {}) {
  const [page, setPage] = useState<PageId>(initialPage);
  const [sessionLoaded, setSessionLoaded] = useState(false);
  const [showDeleteMenu, setShowDeleteMenu] = useState(false);
  const [sessionRestoredNotice, setSessionRestoredNotice] = useState(false);
  const [activeProjectSourceFiles, setActiveProjectSourceFiles] = useState<NonNullable<SavedProject["stageASourceFiles"]>>([]);

  const [bankFile, setBankFile]   = useState<File|null>(null);
  const [bankFileCount, setBankFileCount] = useState(0);
  const [bankHeaders, setBankH]   = useState<string[]>([]);
  const [bankRowsRaw, setBankRows] = useState<Record<string,unknown>[]>([]);
  const [bankMap, setBankMap]     = useState({ date:"", desc:"", debit:"", credit:"", accountType:"", ref:"" });
  const [bankSwap, setBankSwap]   = useState(false);
  const [bankFileSessionId, setBankFileSessionId] = useState(0);

  const [cashFile, setCashFile]   = useState<File|null>(null);
  const [cashFileCount, setCashFileCount] = useState(0);
  const [cashHeaders, setCashH]   = useState<string[]>([]);
  const [cashRowsRaw, setCashRows] = useState<Record<string,unknown>[]>([]);
  const [cashMap, setCashMap]     = useState({ date:"", name:"", debit:"", credit:"", accountType:"", ref:"", userId:"" });
  const [cashSwap, setCashSwap]   = useState(false);
  const [cashFileSessionId, setCashFileSessionId] = useState(0);
  const [stageAInvoices, setStageAInvoices] = useState<StageAResult[]>([]);
  const [stageADatabaseSummary, setStageADatabaseSummary] = useState<StageASummary>({ databaseInvoiceCount: 0, databaseInvoiceTotal: 0, databaseCashTotal: 0, databaseDebtTotal: 0, databaseCardTotal: 0 });
  const [walletBankFile, setWalletBankFile] = useState<File|null>(null);
  const [walletBankFileCount, setWalletBankFileCount] = useState(0);
  const [walletBankHeaders, setWalletBankHeaders] = useState<string[]>([]);
  const [walletBankRows, setWalletBankRows] = useState<Record<string,unknown>[]>([]);
  const [walletBankMap, setWalletBankMap] = useState({ date:"", desc:"", debit:"", credit:"", ref:"" });
  const [walletBankSwap, setWalletBankSwap] = useState(false);
  const [walletBankFileSessionId, setWalletBankFileSessionId] = useState(0);
  const [walletCashFile, setWalletCashFile] = useState<File|null>(null);
  const [walletCashFileCount, setWalletCashFileCount] = useState(0);
  const [walletCashHeaders, setWalletCashHeaders] = useState<string[]>([]);
  const [walletCashRows, setWalletCashRows] = useState<Record<string,unknown>[]>([]);
  const [walletCashMap, setWalletCashMap] = useState({ date:"", name:"", debit:"", credit:"", ref:"", userId:"" });
  const [walletCashSwap, setWalletCashSwap] = useState(false);
  const [walletCashFileSessionId, setWalletCashFileSessionId] = useState(0);
  const [jawwalBankFile, setJawwalBankFile] = useState<File|null>(null);
  const [jawwalBankFileCount, setJawwalBankFileCount] = useState(0);
  const [jawwalBankHeaders, setJawwalBankHeaders] = useState<string[]>([]);
  const [jawwalBankRows, setJawwalBankRows] = useState<Record<string,unknown>[]>([]);
  const [jawwalBankMap, setJawwalBankMap] = useState({ date:"", desc:"", debit:"", credit:"", ref:"", totalAmount:"", movementAmount:"", direction:"", movementType:"" });
  const [jawwalBankSwap, setJawwalBankSwap] = useState(true);
  const [jawwalBankOptions, setJawwalBankOptions] = useState({ useTotalAmount: true, readDebitCredit: true });
  const [jawwalBankFileSessionId, setJawwalBankFileSessionId] = useState(0);
  const [jawwalCashFile, setJawwalCashFile] = useState<File|null>(null);
  const [jawwalCashFileCount, setJawwalCashFileCount] = useState(0);
  const [jawwalCashHeaders, setJawwalCashHeaders] = useState<string[]>([]);
  const [jawwalCashRows, setJawwalCashRows] = useState<Record<string,unknown>[]>([]);
  const [jawwalCashMap, setJawwalCashMap] = useState({ date:"", name:"", debit:"", credit:"", ref:"", userId:"" });
  const [jawwalCashSwap, setJawwalCashSwap] = useState(false);
  const [jawwalCashFileSessionId, setJawwalCashFileSessionId] = useState(0);
  const nextCustomSourceSessionId = useRef(1);
  const [platformNames, setPlatformNames] = useState({ bank: "بنك فلسطين", wallet: "بال بي", jawwal: "جوال بي" });
  const [customPlatforms, setCustomPlatforms] = useState<CustomPlatform[]>([]);
  const nextCustomPlatformId = useRef(0);

  // ─── File snapshots for restore functionality ──────────────────────────────
  const [bankFileSnapshots, setBankFileSnapshots] = useState<FileSnapshot[]>([]);
  const [cashFileSnapshots, setCashFileSnapshots] = useState<FileSnapshot[]>([]);
  const [walletBankFileSnapshots, setWalletBankFileSnapshots] = useState<FileSnapshot[]>([]);
  const [walletCashFileSnapshots, setWalletCashFileSnapshots] = useState<FileSnapshot[]>([]);
  const [jawwalBankFileSnapshots, setJawwalBankFileSnapshots] = useState<FileSnapshot[]>([]);
  const [jawwalCashFileSnapshots, setJawwalCashFileSnapshots] = useState<FileSnapshot[]>([]);
  const [showBankSnapshots, setShowBankSnapshots] = useState(false);
  const [showCashSnapshots, setShowCashSnapshots] = useState(false);
  const [showWalletBankSnapshots, setShowWalletBankSnapshots] = useState(false);
  const [showWalletCashSnapshots, setShowWalletCashSnapshots] = useState(false);
  const [showJawwalBankSnapshots, setShowJawwalBankSnapshots] = useState(false);
  const [showJawwalCashSnapshots, setShowJawwalCashSnapshots] = useState(false);

  const [loading, setLoading] = useState(false);
  const [activeFileImports, setActiveFileImports] = useState(0);
  const [error, setError]     = useState<string|null>(null);
  const [results, setResults] = useState<MatchResult[]|null>(null);
  const [tab, setTab]         = useState<TabId>("saved");

  const [manualGroups, setManualGroups]     = useState<ManualMatchGroup[]>([]);
  const [savedMatches, setSavedMatches]     = useState<SavedMatch[]>([]);
  const [reviewItems, setReviewItems] = useState<MatchReviewItem[]>([]);
  const [bankSplits, setBankSplits] = useState<BankSplit[]>([]);
  const [cutoffBatches, setCutoffBatches] = useState<CutoffBatch[]>([]);
  const [clearingGroups, setClearingGroups] = useState<ClearingGroup[]>([]);
  const [nameAliases, setNameAliases]       = useState<Record<string, string>>({});
  const [rejectedPairs, setRejected]        = useState<Set<string>>(new Set());
  const [expandedMatchKey, setExpandedMatchKey] = useState<string|null>(null);
  const [expandedUnmatchedCashier, setExpandedUnmatchedCashier] = useState<number|null>(null);
  const [selectedPendingKeys, setSelectedPendingKeys] = useState<Set<string>>(new Set());
  const [selectedSavedIds, setSelectedSavedIds] = useState<Set<string>>(new Set());
  const [nonReconBankTypeFilter, setNonReconBankTypeFilter] = useState("all");
  const [selectedNonReconBankKeys, setSelectedNonReconBankKeys] = useState<Set<string>>(new Set());
  const [stageDBankInputs, setStageDBankInputs] = useState<StageDBankInput[]>([]);
  const [canUndoMatch, setCanUndoMatch] = useState(false);
  const [undoMatchCount, setUndoMatchCount] = useState(0);
  const matchUndoRef = useRef<UndoAction[]>([]);

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [drawerMode, setDrawerMode] = useState<"bank" | "cashier">("bank");
  const [drawerCashier, setDrawerCashier] = useState<CashierRow | null>(null);
  const [drawerOldBank, setDrawerOldBank] = useState<BankRow | null>(null);
  const [drawerNameSearch, setDrawerNameSearch] = useState("");
  const [drawerAmountFrom, setDrawerAmountFrom] = useState("");
  const [drawerAmountTo, setDrawerAmountTo] = useState("");
  const [draggedItemId, setDraggedItemId] = useState<number | null>(null);
  const [dropTargetKey, setDropTargetKey] = useState<string | null>(null);
  const [pendingPage, setPendingPage] = useState(1);
  const [keyboardPendingIndex, setKeyboardPendingIndex] = useState(0);
  const PENDING_PAGE_SIZE = 20;
  const pendingRowRefs = useRef<Record<string, HTMLTableRowElement | null>>({});
  const pendingRowCheckboxRefs = useRef<Record<string, HTMLDivElement | null>>({});
  const assistantSkipRef = useRef<((row: CashierRow) => boolean) | null>(null);
  const [amountTolerancePercent, setAmountTolerancePercent] = useState(0.5);
  const [showSettings, setShowSettings] = useState(false);
  const [resultSearch, setResultSearch] = useState("");
  const [resultType, setResultType] = useState<"all" | "مدفوع" | "مستلم">("all");
  const [minAmount, setMinAmount] = useState("");
  const [maxAmount, setMaxAmount] = useState("");
  const [toast, setToast] = useState<string | null>(null);
  const [referenceLocator, setReferenceLocator] = useState<{ row: BankRow | CashierRow; locations: ReferenceLocation[] } | null>(null);
  const [referenceFocus, setReferenceFocus] = useState<string | null>(null);

  const [visaItems, setVisaItems] = useState<VisaItem[]>([]);

  const [heldItems, setHeldItems] = useState<HeldItem[]>([]);
  const [returnedHeldBank, setReturnedHeldBank] = useState<BankRow[]>([]);
  const [returnedHeldCashier, setReturnedHeldCashier] = useState<CashierRow[]>([]);
  const [rejectedSpecialCashierIds, setRejectedSpecialCashierIds] = useState<Set<string>>(new Set());
  const [mahmoudWalletSelectedIds, setMahmoudWalletSelectedIds] = useState<Set<string>>(new Set());
  const [rowNotes, setRowNotes] = useState<Record<string, string>>({});
  const skipNextAutoReconcile = useRef(false);
  const pushUndoAction = (action: UndoAction) => {
    matchUndoRef.current.push(action);
    if (matchUndoRef.current.length > 50) matchUndoRef.current.shift();
    setCanUndoMatch(true);
    setUndoMatchCount(matchUndoRef.current.length);
  };

  useEffect(() => {
    (async () => {
      const session = await storageGet<any>("current_session", null);
      if (session) {
        setBankH(session.bankHeaders || []);
        setBankRows(session.bankRowsRaw || []);
        setBankMap({ date:"", desc:"", debit:"", credit:"", accountType:"", ...(session.bankMap||{}) });
        setBankSwap(!!session.bankSwap);
        setBankFileSessionId(session.bankFileSessionId ?? 0);
        setCashH(session.cashHeaders || []);
        setCashRows(session.cashRowsRaw || []);
        setCashMap({ date:"", name:"", debit:"", credit:"", accountType:"", ref:"", userId:"", ...(session.cashMap||{}) });
        setCashSwap(!!session.cashSwap);
        setCashFileSessionId(session.cashFileSessionId ?? 0);
        setStageAInvoices(session.stageAInvoices || []);
        setStageADatabaseSummary(session.stageADatabaseSummary ?? { databaseInvoiceCount: 0, databaseInvoiceTotal: 0, databaseCashTotal: 0, databaseDebtTotal: 0, databaseCardTotal: 0 });
        setWalletBankHeaders(session.walletBankHeaders || []);
        setWalletBankRows(session.walletBankRows || []);
        setWalletBankMap({ date:"", desc:"", debit:"", credit:"", ref:"", ...(session.walletBankMap || {}) });
        setWalletBankSwap(!!session.walletBankSwap);
        setWalletBankFileSessionId(session.walletBankFileSessionId ?? 0);
        setWalletCashHeaders(session.walletCashHeaders || []);
        setWalletCashRows(session.walletCashRows || []);
        setWalletCashMap({ date:"", name:"", debit:"", credit:"", ref:"", userId:"", ...(session.walletCashMap || {}) });
        setWalletCashSwap(!!session.walletCashSwap);
        setWalletCashFileSessionId(session.walletCashFileSessionId ?? 0);
        setJawwalBankHeaders(session.jawwalBankHeaders || []);
        setJawwalBankRows(session.jawwalBankRows || []);
        setJawwalBankMap({ date:"", desc:"", debit:"", credit:"", ref:"", totalAmount:"", movementAmount:"", direction:"", movementType:"", ...(session.jawwalBankMap || {}) });
        setJawwalBankSwap(session.jawwalBankSwap ?? true);
        setJawwalBankOptions({ useTotalAmount: true, readDebitCredit: true, ...(session.jawwalBankOptions || {}) });
        setJawwalBankFileSessionId(session.jawwalBankFileSessionId ?? 0);
        setJawwalCashHeaders(session.jawwalCashHeaders || []);
        setJawwalCashRows(session.jawwalCashRows || []);
        setJawwalCashMap({ date:"", name:"", debit:"", credit:"", ref:"", userId:"", ...(session.jawwalCashMap || {}) });
        setJawwalCashSwap(session.jawwalCashSwap ?? false);
        setJawwalCashFileSessionId(session.jawwalCashFileSessionId ?? 0);
        setPlatformNames({ bank: "بنك فلسطين", wallet: "بال بي", jawwal: "جوال بي", ...(session.platformNames || {}) });
        setCustomPlatforms(session.customPlatforms || []);
        nextCustomSourceSessionId.current = Math.max(1, ...(session.customPlatforms || []).flatMap((platform: CustomPlatform) => [platform.transfer?.fileSessionId ?? 0, platform.invoice?.fileSessionId ?? 0]).map((id: number) => id + 1));
        nextCustomPlatformId.current = Math.max(0, ...(session.customPlatforms || []).map((platform: CustomPlatform) => platform.id + 1));
        setBankFileSnapshots(session.bankFileSnapshots || []);
        setCashFileSnapshots(session.cashFileSnapshots || []);
        setWalletBankFileSnapshots(session.walletBankFileSnapshots || []);
        setWalletCashFileSnapshots(session.walletCashFileSnapshots || []);
        setJawwalBankFileSnapshots(session.jawwalBankFileSnapshots || []);
        setJawwalCashFileSnapshots(session.jawwalCashFileSnapshots || []);
        setBankFileCount(session.bankFileSnapshots?.[0]?.fileCount ?? 0);
        setCashFileCount(session.cashFileSnapshots?.[0]?.fileCount ?? 0);
        setWalletBankFileCount(session.walletBankFileSnapshots?.[0]?.fileCount ?? 0);
        setWalletCashFileCount(session.walletCashFileSnapshots?.[0]?.fileCount ?? 0);
        setJawwalBankFileCount(session.jawwalBankFileSnapshots?.[0]?.fileCount ?? 0);
        setJawwalCashFileCount(session.jawwalCashFileSnapshots?.[0]?.fileCount ?? 0);
        setManualGroups(session.manualGroups || []);
        setSavedMatches(dedupeSavedMatches(session.savedMatches || []));
        setReviewItems(session.reviewItems || []);
        setBankSplits(session.bankSplits || []);
        setCutoffBatches(session.cutoffBatches || []);
        setClearingGroups(session.clearingGroups || []);
        setNameAliases(session.nameAliases || {});
        setRejected(new Set(session.rejectedPairs || []));
        setVisaItems(session.visaItems || []);
        setHeldItems(dedupeHeldItems(session.heldItems || []));
        setReturnedHeldBank(session.returnedHeldBank || []);
        setReturnedHeldCashier(session.returnedHeldCashier || []);
        setRejectedSpecialCashierIds(normalizeSpecialKeySet(session.rejectedSpecialCashierIds));
        setMahmoudWalletSelectedIds(new Set(session.mahmoudWalletCashierIds ?? []));
        setRowNotes(session.rowNotes ?? {});
        if ((session.bankRowsRaw?.length || session.savedMatches?.length)) {
          setSessionRestoredNotice(true);
        }
      }
      setSessionLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!sessionLoaded) return;
    const t = setTimeout(() => {
      storageSet("current_session", {
        bankHeaders, bankRowsRaw, bankMap, bankSwap, bankFileSessionId,
        cashHeaders, cashRowsRaw, cashMap, cashSwap, cashFileSessionId, stageAInvoices, stageADatabaseSummary,
        walletBankHeaders, walletBankRows, walletBankMap, walletBankSwap, walletBankFileSessionId,
        walletCashHeaders, walletCashRows, walletCashMap, walletCashSwap, walletCashFileSessionId,
        jawwalBankHeaders, jawwalBankRows, jawwalBankMap, jawwalBankSwap, jawwalBankOptions, jawwalBankFileSessionId,
        jawwalCashHeaders, jawwalCashRows, jawwalCashMap, jawwalCashSwap, jawwalCashFileSessionId,
        platformNames, customPlatforms,
        bankFileSnapshots, cashFileSnapshots, walletBankFileSnapshots, walletCashFileSnapshots,
        jawwalBankFileSnapshots, jawwalCashFileSnapshots,
        manualGroups, savedMatches, reviewItems, bankSplits, cutoffBatches, clearingGroups, nameAliases,
        rejectedPairs: Array.from(rejectedPairs),
        visaItems, heldItems, returnedHeldBank, returnedHeldCashier,
        rejectedSpecialCashierIds: Array.from(rejectedSpecialCashierIds),
        mahmoudWalletCashierIds: Array.from(mahmoudWalletSelectedIds),
        rowNotes
      });
    }, 500);
    return () => clearTimeout(t);
  }, [
    sessionLoaded, bankHeaders, bankRowsRaw, bankMap, bankSwap, bankFileSessionId,
    cashHeaders, cashRowsRaw, cashMap, cashSwap, cashFileSessionId, stageAInvoices, stageADatabaseSummary,
    walletBankHeaders, walletBankRows, walletBankMap, walletBankSwap, walletBankFileSessionId,
    walletCashHeaders, walletCashRows, walletCashMap, walletCashSwap, walletCashFileSessionId,
    jawwalBankHeaders, jawwalBankRows, jawwalBankMap, jawwalBankSwap, jawwalBankOptions, jawwalBankFileSessionId,
    jawwalCashHeaders, jawwalCashRows, jawwalCashMap, jawwalCashSwap, jawwalCashFileSessionId,
    platformNames, customPlatforms,
    bankFileSnapshots, cashFileSnapshots, walletBankFileSnapshots, walletCashFileSnapshots,
    jawwalBankFileSnapshots, jawwalCashFileSnapshots,
    manualGroups, savedMatches, reviewItems, bankSplits, cutoffBatches, clearingGroups, nameAliases, rejectedPairs,
    visaItems, heldItems, returnedHeldBank, returnedHeldCashier, rejectedSpecialCashierIds, mahmoudWalletSelectedIds, rowNotes
  ]);

  const trackFileImport = async (task: () => Promise<void>) => {
    setActiveFileImports(count => count + 1);
    try {
      await task();
    } finally {
      setActiveFileImports(count => Math.max(0, count - 1));
    }
  };

  const loadBank=async(files:File[],append=false)=>{
    setError(null);setResults(null);
    try {
      const previous=append&&bankRowsRaw.length?{fileName:bankFileSnapshots[0]?.label??bankFile?.name??"الكشف السابق",headers:bankHeaders,rows:bankRowsRaw}:undefined;
      const {headers,rows,fileName}=await parseAndMergeFiles(files,"bank",previous);
      const mergedFileCount=(append&&bankRowsRaw.length?Math.max(1,bankFileCount):0)+files.length;
      setBankFile(files[0]);setBankFileCount(mergedFileCount);
      setBankSwap(false);
      setBankH(headers);setBankRows(rows);
      setBankMap({date:autoDetectBankDate(headers),desc:autoDetect(headers,HINTS.desc),debit:autoDetect(headers,HINTS.debit),credit:autoDetect(headers,HINTS.credit),accountType:autoDetect(headers,HINTS.accountType),ref:autoDetect(headers,HINTS.ref)});
      setBankFileSessionId(id => id + 1);
      setBankFileSnapshots(prev => [{ label: fileName, headers, rows, savedAt: new Date().toLocaleString("ar-SA"), fileCount: mergedFileCount }, ...prev].slice(0, 5));
    } catch(e){setError((e as Error).message);}
  };
  const loadCash=async(files:File[],append=false)=>{
    setError(null);setResults(null);
    setRejectedSpecialCashierIds(new Set());
    try {
      const previous=append&&cashRowsRaw.length?{fileName:cashFileSnapshots[0]?.label??cashFile?.name??"الكشف السابق",headers:cashHeaders,rows:cashRowsRaw}:undefined;
      const {headers,rows,fileName}=await parseAndMergeFiles(files,"cash",previous);
      const mergedFileCount=(append&&cashRowsRaw.length?Math.max(1,cashFileCount):0)+files.length;
      setCashFile(files[0]);setCashFileCount(mergedFileCount);
      setCashSwap(false);
      setCashH(headers);setCashRows(rows);
      setCashMap({date:autoDetect(headers,HINTS.date),name:autoDetectInvoiceName(headers),debit:autoDetect(headers,HINTS.debit),credit:autoDetect(headers,HINTS.credit),accountType:autoDetect(headers,HINTS.accountType),ref:autoDetect(headers,HINTS.ref),userId:autoDetect(headers,HINTS.userId)});
      setCashFileSessionId(id => id + 1);
      setCashFileSnapshots(prev => [{ label: fileName, headers, rows, savedAt: new Date().toLocaleString("ar-SA"), fileCount: mergedFileCount }, ...prev].slice(0, 5));
    } catch(e){setError((e as Error).message);}
  };
  const loadWalletBank = async (files: File[], append=false) => {
    setError(null); setResults(null);
    try {
      const previous=append&&walletBankRows.length?{fileName:walletBankFileSnapshots[0]?.label??walletBankFile?.name??"الكشف السابق",headers:walletBankHeaders,rows:walletBankRows}:undefined;
      const { headers, rows, fileName } = await parseAndMergeFiles(files,"bank",previous);
      const mergedFileCount=(append&&walletBankRows.length?Math.max(1,walletBankFileCount):0)+files.length;
      setWalletBankFile(files[0]);setWalletBankFileCount(mergedFileCount);
      setWalletBankHeaders(headers); setWalletBankRows(rows);
      setWalletBankMap({ date:autoDetectBankDate(headers), desc:autoDetect(headers,HINTS.desc), debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit), ref:autoDetect(headers,HINTS.ref) });
      setWalletBankFileSessionId(id => id + 1);
      setWalletBankFileSnapshots(prev => [{ label: fileName, headers, rows, savedAt: new Date().toLocaleString("ar-SA"), fileCount: mergedFileCount }, ...prev].slice(0, 5));
    } catch (e) { setError((e as Error).message); }
  };
  const loadWalletCash = async (files: File[], append=false) => {
    setError(null); setResults(null);
    setRejectedSpecialCashierIds(new Set());
    try {
      const previous=append&&walletCashRows.length?{fileName:walletCashFileSnapshots[0]?.label??walletCashFile?.name??"الكشف السابق",headers:walletCashHeaders,rows:walletCashRows}:undefined;
      const { headers, rows, fileName } = await parseAndMergeFiles(files,"cash",previous);
      const mergedFileCount=(append&&walletCashRows.length?Math.max(1,walletCashFileCount):0)+files.length;
      setWalletCashFile(files[0]);setWalletCashFileCount(mergedFileCount);
      setWalletCashHeaders(headers); setWalletCashRows(rows);
      setWalletCashMap({ date:autoDetect(headers,HINTS.date), name:autoDetectInvoiceName(headers), debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit), ref:autoDetect(headers,HINTS.ref), userId:autoDetect(headers,HINTS.userId) });
      setWalletCashFileSessionId(id => id + 1);
      setWalletCashFileSnapshots(prev => [{ label: fileName, headers, rows, savedAt: new Date().toLocaleString("ar-SA"), fileCount: mergedFileCount }, ...prev].slice(0, 5));
    } catch (e) { setError((e as Error).message); }
  };
  const loadJawwalBank = async (files: File[], append=false) => {
    setError(null); setResults(null);
    try {
      const previous=append&&jawwalBankRows.length?{fileName:jawwalBankFileSnapshots[0]?.label??jawwalBankFile?.name??"الكشف السابق",headers:jawwalBankHeaders,rows:jawwalBankRows}:undefined;
      const merged = await parseAndMergeFiles(files,"jawwalBank",previous);
      const { headers, rows } = cleanJawwalPayLedger(merged.headers, merged.rows);
      const { fileName } = merged;
      const mergedFileCount=(append&&jawwalBankRows.length?Math.max(1,jawwalBankFileCount):0)+files.length;
      setJawwalBankFile(files[0]);setJawwalBankFileCount(mergedFileCount);
      setJawwalBankSwap(true);
      setJawwalBankHeaders(headers); setJawwalBankRows(rows);
      const totalAmount = autoDetect(headers, HINTS.totalAmount);
      const direction = autoDetect(headers, HINTS.direction);
      const debit = autoDetect(headers, HINTS.debit);
      const credit = autoDetect(headers, HINTS.credit);
      setJawwalBankMap({
        date: autoDetectBankDate(headers),
        desc: autoDetect(headers, ["بيان القيد", ...HINTS.desc]),
        debit: debit === direction ? "" : debit,
        credit: credit === direction ? "" : credit,
        ref: autoDetect(headers, HINTS.ref),
        totalAmount,
        movementAmount: autoDetect(headers, HINTS.movementAmount),
        direction,
        movementType: autoDetect(headers, ["نوع الحركة", "transaction type"]),
      });
      setJawwalBankFileSessionId(id => id + 1);
      setJawwalBankFileSnapshots(prev => [{ label: fileName, headers, rows, savedAt: new Date().toLocaleString("ar-SA"), fileCount: mergedFileCount }, ...prev].slice(0, 5));
    } catch (e) { setError((e as Error).message); }
  };
  const loadJawwalCash = async (files: File[], append=false) => {
    setError(null); setResults(null);
    try {
      const previous=append&&jawwalCashRows.length?{fileName:jawwalCashFileSnapshots[0]?.label??jawwalCashFile?.name??"الكشف السابق",headers:jawwalCashHeaders,rows:jawwalCashRows}:undefined;
      const { headers, rows, fileName } = await parseAndMergeFiles(files,"cash",previous);
      const mergedFileCount=(append&&jawwalCashRows.length?Math.max(1,jawwalCashFileCount):0)+files.length;
      setJawwalCashFile(files[0]);setJawwalCashFileCount(mergedFileCount);
      setJawwalCashSwap(false);
      setJawwalCashHeaders(headers); setJawwalCashRows(rows);
      setJawwalCashMap({ date:autoDetect(headers,HINTS.date), name:autoDetectInvoiceName(headers), debit:autoDetect(headers,HINTS.debit), credit:autoDetect(headers,HINTS.credit), ref:autoDetect(headers,HINTS.ref), userId:autoDetect(headers,HINTS.userId) });
      setJawwalCashFileSessionId(id => id + 1);
      setJawwalCashFileSnapshots(prev => [{ label: fileName, headers, rows, savedAt: new Date().toLocaleString("ar-SA"), fileCount: mergedFileCount }, ...prev].slice(0, 5));
    } catch (e) { setError((e as Error).message); }
  };
  const loadCustomSource = async (platformId: number, slot: "transfer" | "invoice", files: File[]) => {
    setError(null);
    setResults(null);
    try {
      const { headers, rows, fileName } = await parseAndMergeFiles(files,slot === "invoice" ? "cash" : "bank");
      const source: PlatformSource = {
        fileName,
        headers,
        rows,
        map: {
          date: slot === "invoice" ? autoDetect(headers, HINTS.date) : autoDetectBankDate(headers),
          description: slot === "invoice" ? autoDetectInvoiceName(headers) : autoDetect(headers, HINTS.desc) || autoDetectInvoiceName(headers),
          debit: autoDetect(headers, HINTS.debit),
          credit: autoDetect(headers, HINTS.credit),
          ref: autoDetect(headers, HINTS.ref),
          userId: autoDetect(headers, HINTS.userId),
        },
        role: slot,
        otherName: "",
        otherRole: slot,
        fileSessionId: nextCustomSourceSessionId.current++,
      };
      setCustomPlatforms((current) => current.map((platform) => platform.id === platformId ? { ...platform, [slot]: source } : platform));
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const parsedBank = useMemo(():BankRow[]=>{
    return bankRowsRaw.map((r,i): BankRow | null => {
      const { debit, credit, rawAmount, type } = resolveDebitCredit(r, bankMap.debit, bankMap.credit, bankSwap);
      if (rawAmount === 0) return null;
      return {
        id:i, date:fmtDate(bankMap.date?r[bankMap.date]:""),
        description:String(cleanStatementText(String(bankMap.desc?r[bankMap.desc]:"")) ?? "").trim(),
        debit, credit, rawAmount, type,
        accountType: platformNames.bank,
        ref: String((bankMap as any).ref ? r[(bankMap as any).ref] : "").trim(),
        defaultRef: generateDefaultReference(platformNames.bank, "transfer", i + 1),
        orig:r, _fileSessionId: bankFileSessionId
      };
    }).filter((r): r is BankRow => r !== null);
  },[bankRowsRaw,bankMap,bankSwap,platformNames.bank,bankFileSessionId]);

  const parseWalletBankRows = useMemo(():BankRow[] => walletBankRows.map((r,i): BankRow | null => {
    const { debit, credit, rawAmount, type } = resolveDebitCredit(r, walletBankMap.debit, walletBankMap.credit, walletBankSwap);
    if (!rawAmount) return null;
    return {
      id: 1000000 + i,
      date:fmtDate(walletBankMap.date ? r[walletBankMap.date] : ""),
      description:String(cleanStatementText(String(walletBankMap.desc ? r[walletBankMap.desc] : "")) ?? "").trim(),
      debit, credit, rawAmount, type,
      accountType:platformNames.wallet,
      ref:String(walletBankMap.ref ? r[walletBankMap.ref] : "").trim(),
      defaultRef: generateDefaultReference(platformNames.wallet, "transfer", i + 1),
      orig:r,
      _fileSessionId:walletBankFileSessionId
    };
  }).filter((r): r is BankRow => r !== null), [walletBankRows, walletBankMap, walletBankSwap, platformNames.wallet, walletBankFileSessionId]);

  const parsedCashier = useMemo(():CashierRow[]=>{
    const rows = cashRowsRaw.map((r,i): CashierRow | null => {
      const rawName=String(cleanStatementText(String(cashMap.name?r[cashMap.name]:"")) ?? "").trim();
      const {name,notes}=parseCashierName(rawName);
      const {splitExpr,matchAmount:ma}=parseNotes(notes);
      const { debit, credit, rawAmount, type } = resolveDebitCredit(r, cashMap.debit, cashMap.credit, cashSwap);
      const amount = rawAmount;
      if (amount === 0 && !ma) return null;
      return {
        id:i, rawName, name, notes, splitExpr, debit, credit, amount,
        matchAmount: ma ?? amount, type,
        accountType: platformNames.bank,
        userId: String(cashMap.userId ? r[cashMap.userId] ?? "" : "").trim(),
        ref: String((cashMap as any).ref ? r[(cashMap as any).ref] : "").trim(),
        defaultRef: generateDefaultReference(platformNames.bank, "invoice", i + 1),
        date:fmtDate(cashMap.date?r[cashMap.date]:""), orig:r, _fileSessionId:cashFileSessionId
      };
    }).filter((r): r is CashierRow => r !== null);

    return rows;
  },[cashRowsRaw,cashMap,cashSwap,platformNames.bank,cashFileSessionId]);

  const parsedWalletCashier = useMemo(():CashierRow[] => walletCashRows.map((r,i): CashierRow | null => {
    const rawName = String(cleanStatementText(String(walletCashMap.name ? r[walletCashMap.name] : "")) ?? "").trim();
    const { name, notes } = parseCashierName(rawName);
    const { splitExpr, matchAmount: ma } = parseNotes(notes);
    const { debit, credit, rawAmount, type } = resolveDebitCredit(r, walletCashMap.debit, walletCashMap.credit, walletCashSwap);
    const amount = rawAmount || ma || 0;
    if (!amount) return null;
    return {
      id: 2000000 + i,
      rawName, name, notes, splitExpr, debit, credit, amount,
      matchAmount:ma ?? amount,
      type,
      accountType:platformNames.wallet,
      userId:String(walletCashMap.userId ? r[walletCashMap.userId] ?? "" : "").trim(),
      ref:String(walletCashMap.ref ? r[walletCashMap.ref] : "").trim(),
      defaultRef: generateDefaultReference(platformNames.wallet, "invoice", i + 1),
      date:fmtDate(walletCashMap.date ? r[walletCashMap.date] : ""),
      orig:r,
      _fileSessionId:walletCashFileSessionId
    };
  }).filter((r): r is CashierRow => r !== null), [walletCashRows, walletCashMap, walletCashSwap, platformNames.wallet, walletCashFileSessionId]);

  const cleanedJawwalBankDataset = useMemo(
    () => cleanJawwalPayLedger(jawwalBankHeaders, jawwalBankRows),
    [jawwalBankHeaders, jawwalBankRows]
  );
  const parseJawwalBankRows = useMemo(():BankRow[] => jawwalBankRows.map((r,i): BankRow | null => {
    if (!isJawwalPayLedgerMovement(jawwalBankHeaders, r)) return null;
    const original = resolveDebitCredit(r, jawwalBankMap.debit, jawwalBankMap.credit, false);
    const amountColumn = jawwalBankOptions.useTotalAmount ? jawwalBankMap.totalAmount : jawwalBankMap.movementAmount;
    const useMappedAmount = !!amountColumn;
    const amount = useMappedAmount ? Math.abs(toNum(r[amountColumn])) : original.rawAmount;
    const direction = jawwalBankOptions.readDebitCredit && jawwalBankMap.direction
      ? String(r[jawwalBankMap.direction] ?? "").trim().toLowerCase()
      : "";
    const transactionTypeColumn = jawwalBankMap.movementType || autoDetect(jawwalBankHeaders, ["نوع الحركة", "transaction type"]);
    const transactionType = transactionTypeColumn ? String(r[transactionTypeColumn] ?? "").trim().toLowerCase() : "";
    const isCredit = /دائن|credit/.test(direction);
    const isDebit = /مدين|debit/.test(direction);
    const isMerchantPayment = /pay\s+to\s+merchant|دفع(?:ة)?\s+(?:إلى|الى|لدى|للتاجر)/i.test(transactionType);
    const type: BankRow["type"] = isMerchantPayment ? "مستلم" : original.rawAmount > 0
      ? original.type
      : isCredit ? "مدفوع" : isDebit ? "مستلم" : original.type;
    let debit = type === "مستلم" ? amount : 0;
    let credit = type === "مدفوع" ? amount : 0;
    let finalType = type;
    if (jawwalBankSwap) {
      [debit, credit] = [credit, debit];
      finalType = type === "مدفوع" ? "مستلم" : "مدفوع";
    }
    const rawAmount = amount;
    if (!rawAmount) return null;
    return {
      id: 4000000 + i,
      date:fmtDate(jawwalBankMap.date ? r[jawwalBankMap.date] : ""),
      description:String(cleanStatementText(String(jawwalBankMap.desc ? r[jawwalBankMap.desc] : "")) ?? "").trim(),
      debit, credit, rawAmount,
      totalAmount: jawwalBankMap.totalAmount ? Math.abs(toNum(r[jawwalBankMap.totalAmount])) : undefined,
      movementType: transactionType,
      type:finalType,
      accountType:platformNames.jawwal,
      ref:String(jawwalBankMap.ref ? r[jawwalBankMap.ref] : "").trim(),
      defaultRef: generateDefaultReference(platformNames.jawwal, "transfer", i + 1),
      orig:r,
      _fileSessionId:jawwalBankFileSessionId
    };
  }).filter((r): r is BankRow => r !== null), [jawwalBankRows, jawwalBankHeaders, jawwalBankMap, jawwalBankSwap, jawwalBankOptions, jawwalBankFileSessionId, platformNames.jawwal]);

  const parsedJawwalCashier = useMemo(():CashierRow[] => jawwalCashRows.map((r,i): CashierRow | null => {
    const rawName = String(cleanStatementText(String(jawwalCashMap.name ? r[jawwalCashMap.name] : "")) ?? "").trim();
    const { name, notes } = parseCashierName(rawName);
    const { splitExpr, matchAmount: ma } = parseNotes(notes);
    const { debit, credit, rawAmount, type } = resolveDebitCredit(r, jawwalCashMap.debit, jawwalCashMap.credit, jawwalCashSwap);
    const amount = rawAmount || ma || 0;
    if (!amount) return null;
    return {
      id: 5000000 + i,
      rawName, name, notes, splitExpr, debit, credit, amount, matchAmount:ma ?? amount,
      type,
      accountType:platformNames.jawwal,
      userId:String(jawwalCashMap.userId ? r[jawwalCashMap.userId] ?? "" : "").trim(),
      ref:String(jawwalCashMap.ref ? r[jawwalCashMap.ref] : "").trim(),
      defaultRef: generateDefaultReference(platformNames.jawwal, "invoice", i + 1),
      date:fmtDate(jawwalCashMap.date ? r[jawwalCashMap.date] : ""),
      orig:r,
      _fileSessionId:jawwalCashFileSessionId
    };
  }).filter((r): r is CashierRow => r !== null), [jawwalCashRows, jawwalCashMap, jawwalCashSwap, jawwalCashFileSessionId, platformNames.jawwal]);

  const parsedCustomRows = useMemo(() => {
    const bank: BankRow[] = [];
    const cashier: CashierRow[] = [];
    for (const platform of customPlatforms) {
      for (const slot of ["transfer", "invoice"] as const) {
        const source = platform[slot];
        if (!source) continue;
        const role = source.role === "other" ? source.otherRole : source.role;
        source.rows.forEach((row, index) => {
          const { debit, credit, rawAmount, type } = resolveDebitCredit(row, source.map.debit, source.map.credit, false);
          const date = fmtDate(source.map.date ? row[source.map.date] : "");
          const ref = String(source.map.ref ? row[source.map.ref] : "").trim();
          const description = String(cleanStatementText(String(source.map.description ? row[source.map.description] : "")) ?? "").trim();
          const baseId = platform.id * 100000 + index;
          if (role === "transfer") {
            if (!rawAmount) return;
            bank.push({
              id: 10000000 + baseId,
              date,
              description,
              debit,
              credit,
              rawAmount,
              type,
              accountType: platform.name.trim() || `منصة مخصصة ${platform.id + 1}`,
              ref,
              defaultRef: generateDefaultReference(platform.name.trim() || `منصة مخصصة ${platform.id + 1}`, "transfer", index + 1),
              orig: row,
              _fileSessionId: source.fileSessionId,
            });
            return;
          }
          const rawName = description;
          const { name, notes } = parseCashierName(rawName);
          const { splitExpr, matchAmount: notedAmount } = parseNotes(notes);
          const amount = rawAmount || notedAmount || 0;
          if (amount === 0 && !notedAmount) return;
          cashier.push({
            id: 20000000 + baseId,
            rawName,
            name,
            notes,
            splitExpr,
            debit,
            credit,
            amount,
            matchAmount: notedAmount ?? amount,
            type,
            accountType: platform.name.trim() || `منصة مخصصة ${platform.id + 1}`,
            userId: String(source.map.userId ? row[source.map.userId] ?? "" : "").trim(),
            ref,
            defaultRef: generateDefaultReference(platform.name.trim() || `منصة مخصصة ${platform.id + 1}`, "invoice", index + 1),
            date,
            orig: row,
            _fileSessionId: source.fileSessionId,
          });
        });
      }
    }
    return { bank, cashier };
  }, [customPlatforms]);

  const heldBankKeys = useMemo(() => new Set(heldItems.filter(h=>h.kind==="bank").map(h => `${h.fileSessionId}:${h.refId}`)), [heldItems]);
  const heldCashierKeys = useMemo(() => new Set(heldItems.filter(h=>h.kind==="cashier").map(h => `${h.fileSessionId}:${h.refId}`)), [heldItems]);
  const bankSourceRows = useMemo(
    () => [...parsedBank, ...parseWalletBankRows, ...parseJawwalBankRows, ...parsedCustomRows.bank],
    [parsedBank, parseWalletBankRows, parseJawwalBankRows, parsedCustomRows]
  );
  const cashierSourceRows = useMemo(
    () => [...stageAInvoices.map((invoice, index) => ({
      id: 3000000 + index, rawName: invoice.originalName, name: invoice.name, notes: invoice.issue, splitExpr: "",
      debit: invoice.type === "مستلم" ? invoice.amount : 0, credit: invoice.type === "مدفوع" ? invoice.amount : 0,
      amount: invoice.amount, matchAmount: invoice.amount, type: invoice.type, accountType: invoice.source, userId: invoice.userId,
      registrationTime: invoice.registrationTime, invoiceTotal: invoice.invoiceTotal, cashAmount: invoice.cashAmount,
      checkAmount: invoice.checkAmount, debtAmount: invoice.debtAmount, discountAmount: invoice.discountAmount,
      cardAmount: invoice.databaseAmount, customerNumber: invoice.customerNumber, invoiceNumber: invoice.invoiceNumber, category: invoice.category,
      ref: invoice.invoiceNumber, defaultRef: generateDefaultReference(invoice.source, "invoice", invoice.id + 1), date: fmtDate(invoice.invoiceDate), orig: invoice.invoice, _fileSessionId: cashFileSessionId
    })), ...parsedCashier, ...parsedWalletCashier, ...parsedJawwalCashier, ...parsedCustomRows.cashier],
    [stageAInvoices, parsedCashier, parsedWalletCashier, parsedJawwalCashier, parsedCustomRows, cashFileSessionId, platformNames]
  );
  const normalizedReturnedBank = useMemo(() => {
    const usedSourceKeys = new Set<string>();
    const normalized = returnedHeldBank.map((returned) => {
      const sessionId = returned._fileSessionId ?? bankFileSessionId;
      const source = bankSourceRows.find(row => {
        const key = `${row._fileSessionId ?? bankFileSessionId}:${row.id}`;
        if (usedSourceKeys.has(key) || (row._fileSessionId ?? bankFileSessionId) !== sessionId) return false;
        return row.id === returned.id || (returned.id < 0 && sameBankMovement(row, returned));
      });
      if (!source) return { ...returned, _fileSessionId: sessionId, _fromHeld: true };
      usedSourceKeys.add(`${source._fileSessionId ?? bankFileSessionId}:${source.id}`);
      return { ...returned, id: source.id, _fileSessionId: sessionId, _fromHeld: true };
    });
    return normalized.filter((row, index) =>
      normalized.findIndex(candidate =>
        candidate.id === row.id && (candidate._fileSessionId ?? bankFileSessionId) === (row._fileSessionId ?? bankFileSessionId)
      ) === index
    );
  }, [returnedHeldBank, bankSourceRows, bankFileSessionId]);
  const normalizedReturnedCashier = useMemo(() => {
    const usedSourceKeys = new Set<string>();
    const normalized = returnedHeldCashier.map((returned) => {
      const sessionId = returned._fileSessionId ?? cashFileSessionId;
      const source = cashierSourceRows.find(row => {
        const key = `${row._fileSessionId ?? cashFileSessionId}:${row.id}`;
        if (usedSourceKeys.has(key) || (row._fileSessionId ?? cashFileSessionId) !== sessionId) return false;
        return row.id === returned.id || (returned.id < 0 && sameCashierMovement(row, returned));
      });
      if (!source) return { ...returned, _fileSessionId: sessionId, _fromHeld: true };
      usedSourceKeys.add(`${source._fileSessionId ?? cashFileSessionId}:${source.id}`);
      return { ...returned, id: source.id, _fileSessionId: sessionId, _fromHeld: true };
    });
    return normalized.filter((row, index) =>
      normalized.findIndex(candidate =>
        candidate.id === row.id && (candidate._fileSessionId ?? cashFileSessionId) === (row._fileSessionId ?? cashFileSessionId)
      ) === index
    );
  }, [returnedHeldCashier, cashierSourceRows, cashFileSessionId]);
  const returnedBankKeys = useMemo(() => new Set(normalizedReturnedBank.map(row => `${row._fileSessionId ?? bankFileSessionId}:${row.id}`)), [normalizedReturnedBank, bankFileSessionId]);
  const returnedCashierKeys = useMemo(() => new Set(normalizedReturnedCashier.map(row => `${row._fileSessionId ?? cashFileSessionId}:${row.id}`)), [normalizedReturnedCashier, cashFileSessionId]);
  const activeBank = useMemo(
    () => bankSourceRows
      .filter(row => {
        const key = `${row._fileSessionId ?? bankFileSessionId}:${row.id}`;
        return !heldBankKeys.has(key) && !returnedBankKeys.has(key);
      })
      .concat(normalizedReturnedBank.filter(row => !heldBankKeys.has(`${row._fileSessionId ?? bankFileSessionId}:${row.id}`))),
    [bankSourceRows, heldBankKeys, returnedBankKeys, bankFileSessionId, normalizedReturnedBank]
  );
  const activeCashier = useMemo(
    () => cashierSourceRows
      .filter(row => {
        const key = `${row._fileSessionId ?? cashFileSessionId}:${row.id}`;
        return !heldCashierKeys.has(key) && !returnedCashierKeys.has(key);
      })
      .concat(normalizedReturnedCashier.filter(row => !heldCashierKeys.has(`${row._fileSessionId ?? cashFileSessionId}:${row.id}`))),
    [cashierSourceRows, heldCashierKeys, returnedCashierKeys, cashFileSessionId, normalizedReturnedCashier]
  );
  const nonReconciliationBankRows = useMemo(
    () => activeBank.flatMap((bank) => {
      const reason = getNonReconciliationBankReason(bank);
      return reason ? [{ bank, reason }] : [];
    }),
    [activeBank]
  );
  const nonReconciliationBankKeys = useMemo(
    () => new Set(nonReconciliationBankRows.map(({ bank }) => `${bank._fileSessionId ?? bankFileSessionId}:${bank.id}`)),
    [nonReconciliationBankRows, bankFileSessionId]
  );
  const matchableBank = useMemo(
    () => activeBank.filter((bank) => !nonReconciliationBankKeys.has(`${bank._fileSessionId ?? bankFileSessionId}:${bank.id}`)),
    [activeBank, nonReconciliationBankKeys, bankFileSessionId]
  );
  const splitAdjustedBank = useMemo(() => {
    const reportRows: BankRow[] = [];
    const matchingRows: BankRow[] = [];
    for (const bank of matchableBank) {
      const split = bankSplits.find(entry => entry.bankId === bank.id && entry.bankSessionId === (bank._fileSessionId ?? 0));
      if (!split) {
        reportRows.push(bank);
        matchingRows.push(bank);
        continue;
      }
      bankSplitAllocationGroups(split).forEach((allocations, cashierKey) => {
        const amount = allocations.reduce((sum, allocation) => sum + allocation.amount, 0);
        const row: BankRow = {
          ...bank,
          id: bankSplitVirtualId(split, cashierKey, allocations),
          description: bank.description,
          rawAmount: amount,
          debit: bank.type === "مستلم" ? amount : 0,
          credit: bank.type === "مدفوع" ? amount : 0,
          orig: {
            ...bank.orig,
            _splitId: split.id,
            _splitPart: allocations.map(allocation => split.allocations.indexOf(allocation) + 1).join(", "),
            _splitNote: allocations.map(allocation => allocation.note).filter(Boolean).join(" · "),
          },
        };
        reportRows.push(row);
        matchingRows.push(row);
      });
    }
    return { reportRows, matchingRows };
  }, [matchableBank, bankSplits]);

  const clearedCashierKeys = useMemo(() => new Set(clearingGroups.flatMap(group =>
    [...group.debitRows, ...group.creditRows].map(reference => `${reference.sessionId}:${reference.id}`)
  )), [clearingGroups]);
  const cutoffAdjustedCashier = useMemo(() => {
    const replacedKeys = new Set<string>();
    const syntheticRows: CashierRow[] = [];
    for (const batch of cutoffBatches) {
      const target = activeCashier.find(row =>
        row.id === batch.aggregateId && (row._fileSessionId ?? 0) === batch.aggregateSessionId
      );
      const sourceRows = batch.sourceRows.map(reference => activeCashier.find(row =>
        row.id === reference.id && (row._fileSessionId ?? 0) === reference.sessionId
      ));
      if (!target || sourceRows.some(row => !row)) continue;
      replacedKeys.add(transactionKey(target));
      sourceRows.forEach(row => { if (row) replacedKeys.add(transactionKey(row)); });
      const selectedRows = sourceRows.filter((row): row is CashierRow => row !== undefined);
      const label = `${target.name} · مجموعة تعطل الكاشير ${batch.dateFrom} إلى ${batch.dateTo}`;
      syntheticRows.push({
        ...target,
        id: batch.syntheticId,
        rawName: label,
        name: label,
        notes: batch.note,
        amount: batch.total,
        matchAmount: batch.total,
        debit: selectedRows.reduce((sum, row) => sum + row.debit, 0),
        credit: selectedRows.reduce((sum, row) => sum + row.credit, 0),
        date: `${batch.dateFrom} — ${batch.dateTo}`,
        orig: { _cutoffBatchId: batch.id, sourceCount: selectedRows.length },
      });
    }
    const openRows = activeCashier.filter(row => !clearedCashierKeys.has(transactionKey(row)));
    return {
      reportRows: openRows.filter(row => !replacedKeys.has(transactionKey(row))).concat(syntheticRows),
      matchingRows: openRows.filter(row => !replacedKeys.has(transactionKey(row))).concat(syntheticRows),
    };
  }, [activeCashier, cutoffBatches, clearedCashierKeys]);
  const reconciliationBankRows = splitAdjustedBank.matchingRows;
  const reconciliationCashierRows = cutoffAdjustedCashier.matchingRows;
  const reportBankRows = splitAdjustedBank.reportRows;
  const reportCashierRows = cutoffAdjustedCashier.reportRows;

  const savedKeys = useMemo(() => new Set(savedMatches.flatMap(saved => {
    const cashier = resolveSavedCashierRow(saved, activeCashier);
    const bank = resolveSavedBankRow(saved, activeBank);
    return cashier && bank ? [`${transactionKey(cashier)}-${transactionKey(bank)}`] : [];
  })), [savedMatches, activeCashier, activeBank]);
  const savedCashierIds = useMemo(() => new Set(savedMatches.flatMap(saved => {
    const cashier = resolveSavedCashierRow(saved, activeCashier);
    const bank = resolveSavedBankRow(saved, activeBank);
    return cashier && bank ? [transactionKey(cashier)] : [];
  })), [savedMatches, activeCashier, activeBank]);
  const savedBankIds = useMemo(() => new Set(savedMatches.flatMap(saved => {
    const cashier = resolveSavedCashierRow(saved, activeCashier);
    const bank = resolveSavedBankRow(saved, activeBank);
    return cashier && bank ? [transactionKey(bank)] : [];
  })), [savedMatches, activeCashier, activeBank]);
  const savedMatchIssues = useMemo(() => {
    if (!bankSourceRows.length || !cashierSourceRows.length) return [];
    const claimedBanks = new Set<string>();
    const claimedCashiers = new Set<string>();
    return savedMatches.flatMap(match => {
      const bankCandidates = savedBankCandidates(match, bankSourceRows);
      const cashierCandidates = savedCashierCandidates(match, cashierSourceRows);
      const bank = resolveSavedBankRow(match, bankSourceRows);
      const cashier = resolveSavedCashierRow(match, cashierSourceRows);
      const bankKey = bank ? transactionKey(bank) : "";
      const cashierKey = cashier ? transactionKey(cashier) : "";
      const duplicateBank = !!bankKey && claimedBanks.has(bankKey);
      const duplicateCashier = !!cashierKey && claimedCashiers.has(cashierKey);
      if (bankKey && !duplicateBank) claimedBanks.add(bankKey);
      if (cashierKey && !duplicateCashier) claimedCashiers.add(cashierKey);
      if (bank && cashier && !duplicateBank && !duplicateCashier) return [];
      return [{
        match,
        bankCandidates,
        cashierCandidates,
        reasons: [
          !bank ? bankCandidates.length > 1 ? "وجدت أكثر من حوالة مرشحة؛ لم يتم اختيار واحدة تلقائياً" : "تعذّر التحقق من الحوالة المحفوظة" : "",
          !cashier ? cashierCandidates.length > 1 ? "وجدت أكثر من فاتورة مرشحة؛ لم يتم اختيار واحدة تلقائياً" : "تعذّر التحقق من الفاتورة المحفوظة" : "",
          duplicateBank ? "الحوالة مرتبطة بمطابقة محفوظة أخرى" : "",
          duplicateCashier ? "الفاتورة مرتبطة بمطابقة محفوظة أخرى" : "",
        ].filter(Boolean),
      }];
    });
  }, [savedMatches, bankSourceRows, cashierSourceRows]);
  const pendingCashierIds = useMemo(
    () => new Set((results?.filter(r => r.type === "pending") as any[] ?? []).map((r:any) => r.cashier.id as number)),
    [results]
  );
  const pendingBankIds = useMemo(
    () => new Set((results?.filter(r => r.type === "pending") as any[] ?? []).map((r:any) => r.bank.id as number)),
    [results]
  );
  const claimedNames = useMemo(() => {
    const s = new Set<string>();
    savedMatches.forEach(sm => {
      const cn = normName(sm.cashierName);
      const bn = normName(sm.bankDesc);
      if (cn) s.add(cn);
      if (bn) s.add(bn);
    });
    return s;
  }, [savedMatches]);

  const visaCashierIds = useMemo(() => new Set(activeCashier
    .filter(row => visaItems.some(item => item.cashierId === row.id &&
      (item.cashierSessionId == null || item.cashierSessionId === (row._fileSessionId ?? 0))))
    .map(transactionKey)), [visaItems, activeCashier]);
  const jawwalPayCashierIds = useMemo(() => new Set(
    activeCashier.filter(c => !wasSpecialCashierRejected(rejectedSpecialCashierIds, c) && /جوال\s*بي|jawwal\s*pay/i.test(`${c.rawName} ${c.name}`)).map(transactionKey)
  ), [activeCashier, rejectedSpecialCashierIds]);
  const mahmoudWalletCashierIds = useMemo(() => new Set(
    activeCashier.filter(c => mahmoudWalletSelectedIds.has(transactionKey(c))).map(transactionKey)
  ), [activeCashier, mahmoudWalletSelectedIds]);
  const specialCashierIds = useMemo(() => new Set([...jawwalPayCashierIds, ...mahmoudWalletCashierIds]), [jawwalPayCashierIds, mahmoudWalletCashierIds]);

  const savedRows = useMemo(() => {
    const savedCashierCandidates = uniqueTransactionRows([
      ...reconciliationCashierRows,
      ...activeCashier,
      ...cashierSourceRows,
      ...heldItems.filter(item => item.kind === "cashier").map(item => item.data as CashierRow),
      ...returnedHeldCashier,
    ]);
    const savedBankCandidates = uniqueTransactionRows([
      ...reconciliationBankRows,
      ...activeBank,
      ...bankSourceRows,
      ...heldItems.filter(item => item.kind === "bank").map(item => item.data as BankRow),
      ...returnedHeldBank,
    ]);
    return savedMatches.flatMap(savedMatch => {
      const cashier = resolveSavedCashierRow(savedMatch, savedCashierCandidates);
      const bank = resolveSavedBankRow(savedMatch, savedBankCandidates);
      if (!cashier || !bank) return [];
      return [{ type: "saved" as const, cashier, bank, savedMatch }];
    }).sort((a, b) => b.bank.rawAmount - a.bank.rawAmount);
  }, [
    savedMatches, reconciliationCashierRows, activeCashier, cashierSourceRows, heldItems, returnedHeldCashier,
    reconciliationBankRows, activeBank, bankSourceRows, returnedHeldBank,
  ]);
  const pendingRows = useMemo(
    () => (results?.filter(r => r.type === "pending") as any[] ?? []).sort((a,b) =>
      Number(b.matchType === "phone") - Number(a.matchType === "phone") ||
      b.bank.rawAmount - a.bank.rawAmount
    ),
    [results]
  );
  const visaRows = useMemo(() => results?.filter(r => r.type === "visa") as any[] ?? [], [results]);
  const uCashierRows = useMemo(
    () => (results?.filter(r => r.type === "unmatchedCashier") as any[] ?? []).sort((a,b) => b.cashier.amount - a.cashier.amount),
    [results]
  );
  const uBankRows = useMemo(
    () => (results?.filter(r => r.type === "unmatchedBank") as any[] ?? []).sort((a,b) => b.bank.rawAmount - a.bank.rawAmount),
    [results]
  );

  const unmatchedBankForSuggestions = useMemo(() =>
    results?.filter(r => r.type === "unmatchedBank").map((r: any) => r.bank as BankRow) ?? []
  , [results]);
  const suggestionOwnerMap = useMemo(() => {
    const map = new Map<string, {cashierKey:string; score:number}>();
    uCashierRows.forEach((r:any) => {
      const cashierRow = r.cashier as CashierRow;
      unmatchedBankForSuggestions.forEach(b => {
        if (cashierRow.type !== b.type) return;
        const pairKey = transactionPairKey(cashierRow, b);
        if (wasPairRejected(rejectedPairs, cashierRow, b) || savedKeys.has(pairKey)) return;
        const sc = scoreCandidate(cashierRow, b, claimedNames, nameAliases);
        if (!sc) return;
        const bankKey = transactionKey(b);
        const cur = map.get(bankKey);
        if (!cur || sc.score > cur.score) map.set(bankKey, { cashierKey: transactionKey(cashierRow), score: sc.score });
      });
    });
    return map;
  }, [uCashierRows, unmatchedBankForSuggestions, rejectedPairs, savedKeys, claimedNames, nameAliases]);

  const rerun = useCallback(() => {
    if (!reconciliationBankRows.length || !reconciliationCashierRows.length) return;
    const res = reconcile(reconciliationBankRows, reconciliationCashierRows, manualGroups, savedMatches, rejectedPairs, visaCashierIds, jawwalPayCashierIds, mahmoudWalletCashierIds, amountTolerancePercent, nameAliases, platformNames.jawwal);
    setResults(res);
  }, [reconciliationBankRows, reconciliationCashierRows, manualGroups, savedMatches, rejectedPairs, visaCashierIds, jawwalPayCashierIds, mahmoudWalletCashierIds, amountTolerancePercent, nameAliases, platformNames.jawwal]);

  const run = async() => {
    setError(null);
    if (!bankRowsRaw.length) {
      setError("ارفع ملف الحوالات أولاً قبل تشغيل المطابقة.");
      return;
    }
    if (!cashRowsRaw.length) {
      setError("ارفع ملف فواتير الكاشير أولاً قبل تشغيل المطابقة.");
      return;
    }
    if (!bankMap.desc || !cashMap.name || (!cashMap.debit && !cashMap.credit)) {
      setError("حدد أعمدة البيان والاسم ومبلغ الفاتورة قبل تشغيل المطابقة.");
      return;
    }
    setLoading(true);
    try {
      await new Promise<void>(resolve => setTimeout(resolve, 0));
      if (!reconciliationBankRows.length) throw new Error("كل حركات البنك مصنّفة كمعلّقات غير تابعة للمطابقة.");
      const startedAt = performance.now();
      const metrics: ReconcilePerformanceMetrics = { possiblePairs: 0, amountCandidates: 0, nameChecks: 0 };
      const res = reconcile(reconciliationBankRows, reconciliationCashierRows, manualGroups, savedMatches, rejectedPairs, visaCashierIds, jawwalPayCashierIds, mahmoudWalletCashierIds, amountTolerancePercent, nameAliases, platformNames.jawwal, metrics);
      const elapsedMs = performance.now() - startedAt;
      setResults(res);
      setTab("saved");
      setToast(`تمت مطابقة ${res.filter(r => r.type === "pending" || r.type === "saved").length} سجل خلال ${elapsedMs.toFixed(0)} مللي ثانية · ${metrics.amountCandidates.toLocaleString("ar")} مرشح مبلغ و${metrics.nameChecks.toLocaleString("ar")} فحص اسم من ${metrics.possiblePairs.toLocaleString("ar")} زوج نظري`);
    } catch(e){setError((e as Error).message);}
    finally{setLoading(false);}
  };

  const handleMoveToVisa = (cashierRow: CashierRow, source?: "pending" | "unmatched" | "saved" | "manual") => {
    const visaNum = extractVisaNumber(cashierRow.rawName);
    if (!visaNum) {
      setToast("هذه الفاتورة لا تحتوي على رقم فيزا من 4 أرقام بعد / أو -.");
      return;
    }
    if (visaItems.some(v => v.cashierId === cashierRow.id &&
      (v.cashierSessionId == null || v.cashierSessionId === (cashierRow._fileSessionId ?? 0)))) {
      setToast("هذه الفاتورة موجودة بالفعل في قائمة الفيزا.");
      return;
    }
    clearMatchUndo();
    const cleanName = extractVisaName(cashierRow.rawName);
    const item: VisaItem = {
      id: `visa-${Date.now()}-${cashierRow.id}`, cashierId: cashierRow.id,
      cashierSessionId: cashierRow._fileSessionId ?? 0,
      defaultRef: defaultReferenceForRow(cashierRow),
      rawName: cashierRow.rawName, name: cleanName, visaNumber: visaNum,
      amount: cashierRow.amount, type: cashierRow.type, date: cashierRow.date,
      movedAt: new Date().toLocaleString("ar-SA"), source
    };
    setVisaItems(prev => [...prev, item]);
    if (source === "saved") {
      setSavedMatches(prev => prev.filter(saved => {
        const owner = resolveSavedCashierRow(saved, activeCashier);
        return !owner || transactionKey(owner) !== transactionKey(cashierRow);
      }));
    }
    setExpandedMatchKey(null);
    setExpandedUnmatchedCashier(null);
  };

  const handleRemoveVisaItem = (id: string) => {
    setVisaItems(prev => prev.filter(v => v.id !== id));
  };

  const handleRejectSpecialCashier = (cashierRow: CashierRow) => {
    const previousRejected = new Set(rejectedSpecialCashierIds);
    const previousWallet = new Set(mahmoudWalletSelectedIds);
    pushUndoAction({
      type: "custom",
      label: "إرجاع الفاتورة للمطابقة",
      undo: () => {
        setRejectedSpecialCashierIds(previousRejected);
        setMahmoudWalletSelectedIds(previousWallet);
      },
    });
    setRejectedSpecialCashierIds(prev => {
      const next = new Set(prev);
      next.add(transactionKey(cashierRow));
      return next;
    });
    setMahmoudWalletSelectedIds(prev => {
      const next = new Set(prev);
      next.delete(transactionKey(cashierRow));
      return next;
    });
    setToast("تم ��فض التصنيف وإرجاع الفاتورة إلى المطابقة.");
  };

  const handleMoveToMahmoudWallet = (cashierRow: CashierRow) => {
    const previousWallet = new Set(mahmoudWalletSelectedIds);
    const previousRejected = new Set(rejectedSpecialCashierIds);
    pushUndoAction({
      type: "custom",
      label: "ترحيل الفاتورة إلى محفظة شخصية",
      undo: () => {
        setMahmoudWalletSelectedIds(previousWallet);
        setRejectedSpecialCashierIds(previousRejected);
      },
    });
    setMahmoudWalletSelectedIds(previous => new Set(previous).add(transactionKey(cashierRow)));
    setRejectedSpecialCashierIds(previous => {
      const next = new Set(previous);
      next.delete(transactionKey(cashierRow));
      return next;
    });
    setToast("تم ترحيل الفاتورة إلى قائمة محفظة شخصية.");
  };

  const handleToggleReviewSnapshot = (
    cashier?: CashierRow,
    bank?: BankRow,
    detail?: string,
    savedMatch?: SavedMatch,
  ) => {
    if (!cashier && !bank) return;
    const id = savedMatch?.id ?? (cashier && bank
      ? `review-pair:${transactionPairKey(cashier, bank)}`
      : cashier
        ? `review-cashier:${transactionKey(cashier)}`
        : `review-bank:${transactionKey(bank!)}`);
    const existing = reviewItems.some(item => item.id === id);
    if (existing) {
      setReviewItems(previous => previous.filter(item => item.id !== id));
      setToast("أزيلت النسخة من قائمة المراجعة فقط؛ لم تتغير الحركات أو المطابقات.");
      return;
    }
    setReviewItems(previous => [...previous, {
      id,
      savedMatch: savedMatch ? { ...savedMatch } : undefined,
      cashier: cashier ? { ...cashier, orig: { ...cashier.orig } } : undefined,
      bank: bank ? { ...bank, orig: { ...bank.orig } } : undefined,
      title: cashier && bank ? undefined : cashier ? cashier.name : bank?.description,
      detail,
      addedAt: new Date().toLocaleString("ar-SA"),
    }]);
    setToast("أُضيفت نسخة مستقلة إلى قائمة المراجعة؛ لن تتغير الحركة الأصلية.");
  };

  const handleAddReviewItem = (savedMatch: SavedMatch, cashier: CashierRow, bank: BankRow) =>
    handleToggleReviewSnapshot(cashier, bank, undefined, savedMatch);

  const handleToolReview = (
    cashierTool?: ToolCashierRow,
    bankTool?: ToolBankRow,
    detail?: string,
  ) => {
    const cashier = cashierTool
      ? activeCashier.find(row => transactionKey(row) === transactionKey(cashierTool))
      : undefined;
    const bank = bankTool
      ? activeBank.find(row => transactionKey(row) === transactionKey(bankTool))
      : undefined;
    handleToggleReviewSnapshot(cashier, bank, detail);
  };

  const handleToolCashierReview = (cashierTool?: ToolCashierRow, detail?: string) =>
    handleToolReview(cashierTool, undefined, detail);

  const handleCompleteReviewItem = (id: string) => {
    setReviewItems(previous => previous.filter(item => item.id !== id));
    setToast("تم حذف النسخة من قائمة المراجعة؛ المطابقة الأصلية لم تتغير.");
  };

  const handleSaveRowNote = (row: BankRow | CashierRow, note: string) => {
    const key = transactionKey(row);
    setRowNotes(previous => {
      const next = { ...previous };
      if (note) next[key] = note;
      else delete next[key];
      return next;
    });
    setToast(note ? "تم حفظ ملاحظة الحركة." : "تم حذف ملاحظة الحركة.");
  };

  const handleSaveMatch = (
    cashierRow: CashierRow,
    bankRow: BankRow,
    note?: string,
    options?: { force?: boolean; replacePair?: { cashier: CashierRow; bank: BankRow } }
  ): SavedMatch | null => {
    const savedPairKey = `${transactionKey(cashierRow)}-${transactionKey(bankRow)}`;
    const rejectedPairKey = transactionPairKey(cashierRow, bankRow);
    if (savedKeys.has(savedPairKey) && !options?.force) return null;
    const cashierPhone = extractMobileNumber(`${cashierRow.rawName} ${cashierRow.name} ${cashierRow.ref}`);
    const bankPhone = extractMobileNumber(`${bankRow.description} ${bankRow.ref}`);
    const hasPhoneMatch = !!cashierPhone && cashierPhone === bankPhone;
    const existingOwners = savedMatches.filter(match => {
      const ownerCashier = resolveSavedCashierRow(match, activeCashier);
      const ownerBank = resolveSavedBankRow(match, activeBank);
      return (ownerCashier && transactionKey(ownerCashier) === transactionKey(cashierRow)) ||
        (ownerBank && transactionKey(ownerBank) === transactionKey(bankRow));
    });
    const replacePairKey = options?.replacePair
      ? `${transactionKey(options.replacePair.cashier)}-${transactionKey(options.replacePair.bank)}`
      : "";
    if (existingOwners.some(match => {
      const ownerCashier = resolveSavedCashierRow(match, activeCashier);
      const ownerBank = resolveSavedBankRow(match, activeBank);
      return !ownerCashier || !ownerBank || `${transactionKey(ownerCashier)}-${transactionKey(ownerBank)}` !== replacePairKey;
    })) {
      setToast("لا يمكن تكرار الفاتورة أو الحوالة؛ إحدى الحركتين مستخدمة في مطابقة محفوظة.");
      return null;
    }
    pushUndoAction({ type: "match", snapshot: {
      manualGroups: [...manualGroups],
      savedMatches: [...savedMatches],
      rejectedPairs: new Set(rejectedPairs),
      nameAliases: { ...nameAliases },
      selectedPendingKeys: new Set(selectedPendingKeys),
    } });
    skipNextAutoReconcile.current = true;

    const isAmountDiff = Math.abs(bankRow.rawAmount - cashierRow.matchAmount) > 0.01;
    const ms = advancedMatchCheck(cashierRow.name, bankRow.description);
    const isNameDiff = !hasPhoneMatch && (ms.isApprox || ms.matchType === "none");
    const isPhoneDiff = hasPhoneDifference(cashierRow, bankRow);
    const isAccountTypeDiff = accountTypesDiffer(bankRow.accountType, cashierRow.accountType);
    const isDateDiff = transactionDatesDiffer(bankRow.date, cashierRow.date);

    const autoNote = note || [
      hasPhoneMatch ? `مطابقة برقم الجوال: ${cashierPhone}` : "",
      isPhoneDiff ? `الرقم بالفاتورة مختلف؛ تعديل الرقم إلى ${bankPhone || "الرقم الموجود في الحوالة"}` : "",
      isAmountDiff ? `اختلاف مبلغ: ${fmtNum(Math.abs(bankRow.rawAmount - cashierRow.matchAmount))}` : "",
      isNameDiff ? `اختلاف اسم: ${cashierRow.name} ↔ ${bankRow.description}` : "",
      isAccountTypeDiff ? `نوع الحساب مختلف: بنك (${bankRow.accountType || "—"}) ↔ كاشير (${cashierRow.accountType || "—"})` : "",
      isDateDiff ? `تاريخ الفاتورة (${cashierRow.date}) مختلف عن تاريخ الحوالة (${bankRow.date})` : ""
    ].filter(Boolean).join(" | ") || undefined;

    const newSaved: SavedMatch = {
      id: `saved-${Date.now()}-${cashierRow.id}-${bankRow.id}`,
      cashierId: cashierRow.id, bankId: bankRow.id,
      cashierSessionId: cashierRow._fileSessionId ?? 0, bankSessionId: bankRow._fileSessionId ?? 0,
      cashierDefaultRef: cashierRow.defaultRef || transactionKey(cashierRow),
      bankDefaultRef: bankRow.defaultRef || transactionKey(bankRow),
      cashierName: cashierRow.name, bankDesc: bankRow.description,
      cashierIdentity: movementIdentity(cashierRow), bankIdentity: movementIdentity(bankRow),
      cashierDate: cashierRow.date, bankDate: bankRow.date,
      cashierAmount: cashierRow.amount, cashierRawName: cashierRow.rawName,
      amount: bankRow.rawAmount, type: cashierRow.type,
      date: new Date().toLocaleDateString("ar-SA"),
      savedAt: new Date().toLocaleString("ar-SA"),
      note: autoNote, isAmountDiff, isPhoneDiff, isNameDiff, isManual: false, isAccountTypeDiff, isDateDiff,
      matchScore: hasPhoneMatch ? 100 : Math.round(nameSim(cashierRow.name, bankRow.description) * 100),
      editorNotes: note,
      bankAccountType: bankRow.accountType, cashierAccountType: cashierRow.accountType
    };

    setSavedMatches(prev => dedupeSavedMatches([...prev, newSaved]));
    setResults(prev => {
      if (!prev) return prev;
      const savedResult: MatchResult = { type: "saved", cashier: cashierRow, bank: bankRow, savedMatch: newSaved };
      const next = prev.filter(result => {
        const sameCashier = "cashier" in result && transactionKey(result.cashier) === transactionKey(cashierRow);
        const sameBank = "bank" in result && transactionKey(result.bank) === transactionKey(bankRow);
        return !sameCashier && !sameBank;
      });
      return [savedResult, ...next];
    });
    if (isNameDiff) {
      setNameAliases(prev => ({ ...prev, [smartNameKey(cashierRow.name)]: smartNameKey(smartAliasTarget(cashierRow.name, bankRow.description)) }));
    }
    const nextRejected = new Set(rejectedPairs);
    nextRejected.delete(rejectedPairKey);
    if ((cashierRow._fileSessionId ?? 0) === 0 && (bankRow._fileSessionId ?? 0) === 0) {
      nextRejected.delete(`${cashierRow.id}-${bankRow.id}`);
    }
    setRejected(nextRejected);
    onSafeActivities?.([`${cashierRow.name} ↔ ${bankRow.description} · ${fmtNum(bankRow.rawAmount)} ₪`]);
    setExpandedMatchKey(null);
    return newSaved;
  };

  const clearMatchUndo = () => {
    matchUndoRef.current = [];
    setCanUndoMatch(false);
    setUndoMatchCount(0);
  };

  const handleUndoMatch = () => {
    const action = matchUndoRef.current.pop();
    if (!action) return;
    setCanUndoMatch(matchUndoRef.current.length > 0);
    setUndoMatchCount(matchUndoRef.current.length);
    if (action.type === "custom") {
      action.undo();
      setResults(null);
      setToast(`تم التراجع عن: ${action.label}`);
      return;
    }
    const snapshot = action.snapshot;
    skipNextAutoReconcile.current = false;
    setManualGroups(snapshot.manualGroups);
    setSavedMatches(dedupeSavedMatches(snapshot.savedMatches));
    setRejected(new Set(snapshot.rejectedPairs));
    setNameAliases({ ...snapshot.nameAliases });
    setSelectedPendingKeys(new Set(snapshot.selectedPendingKeys));
    setResults(null);
    setToast("تم التراجع عن آخر مطابقة");
  };

  const finishReplacement = (
    oldCashier: CashierRow,
    oldBank: BankRow,
    newCashier: CashierRow,
    newBank: BankRow,
    newSaved: SavedMatch
  ) => {
    const oldPairKey = transactionPairKey(oldCashier, oldBank);
    const newPairKey = transactionPairKey(newCashier, newBank);

    // A selected row owns both sides of the match. Remove any previous
    // saved match using either the selected bank or the selected cashier.
    setSavedMatches(prev => prev.filter(saved => {
      if (saved.id === newSaved.id) return true;
      const savedCashier = resolveSavedCashierRow(saved, activeCashier);
      const savedBank = resolveSavedBankRow(saved, activeBank);
      return (!savedCashier || transactionKey(savedCashier) !== transactionKey(newCashier))
        && (!savedBank || transactionKey(savedBank) !== transactionKey(newBank));
    }));

    setRejected(prev => {
      const next = new Set(prev);
      next.add(oldPairKey);
      next.delete(newPairKey);
      if ((newCashier._fileSessionId ?? 0) === 0 && (newBank._fileSessionId ?? 0) === 0) {
        next.delete(`${newCashier.id}-${newBank.id}`);
      }
      return next;
    });

    // Replace the old row immediately and remove the same bank/cashier from
    // every other visible result, including another pending suggestion.
    setResults(prev => {
      if (!prev) return prev;
      const replacement: MatchResult = {
        type: "saved",
        cashier: newCashier,
        bank: newBank,
        savedMatch: newSaved
      };
      let inserted = false;
      const next: MatchResult[] = [];

      for (const result of prev) {
        const resultCashierKey = "cashier" in result ? transactionKey(result.cashier) : null;
        const resultBankKey = "bank" in result ? transactionKey(result.bank) : null;
        const isOldRow = resultCashierKey === transactionKey(oldCashier) && resultBankKey === transactionKey(oldBank);
        const usesClaimedCashier = resultCashierKey === transactionKey(newCashier);
        const usesClaimedBank = resultBankKey === transactionKey(newBank);

        if (usesClaimedCashier || usesClaimedBank) {
          if (!inserted && isOldRow) {
            next.push(replacement);
            inserted = true;
          }
          continue;
        }
        next.push(result);
      }

      if (!inserted) next.push(replacement);
      return next;
    });
    // Replacement changes ownership of rows, so let the normal reconciliation
    // pass rebuild the remaining suggestions immediately.
    skipNextAutoReconcile.current = false;
  };

  const handleSaveMatchesBulk = (pairs: Array<{cashier: CashierRow; bank: BankRow}>) => {
    if (!pairs.length) return;
    const newSaved: SavedMatch[] = [];
    const usedCashiers = new Set(savedMatches.flatMap(saved => {
      const row = resolveSavedCashierRow(saved, activeCashier);
      return row ? [transactionKey(row)] : [];
    }));
    const usedBanks = new Set(savedMatches.flatMap(saved => {
      const row = resolveSavedBankRow(saved, activeBank);
      return row ? [transactionKey(row)] : [];
    }));
    const usedKeysNow = new Set(savedMatches.flatMap(saved => {
      const cashier = resolveSavedCashierRow(saved, activeCashier);
      const bank = resolveSavedBankRow(saved, activeBank);
      return cashier && bank ? [`${transactionKey(cashier)}-${transactionKey(bank)}`] : [];
    }));
    let skipped = 0;
    let duplicateSkipped = 0;
    const conflictDetails: string[] = [];
    pairs.forEach(({cashier: cashierRow, bank: bankRow}) => {
      const cashierKey = transactionKey(cashierRow);
      const bankKey = transactionKey(bankRow);
      const pairKey = `${cashierKey}-${bankKey}`;
      if (usedKeysNow.has(pairKey) || usedCashiers.has(cashierKey) || usedBanks.has(bankKey)) {
        skipped++;
        duplicateSkipped++;
        const reasons = [
          usedKeysNow.has(pairKey) ? "الزوج محفوظ مسبقاً" : "",
          usedCashiers.has(cashierKey) ? "صف الكاشير مستخدم في مطابقة أخرى" : "",
          usedBanks.has(bankKey) ? "صف البنك مستخدم في مطابقة أخرى" : "",
        ].filter(Boolean).join("، ");
        conflictDetails.push(`${cashierRow.name || cashierRow.rawName} ↔ ${bankRow.description} (${reasons})`);
        return;
      }
      usedKeysNow.add(pairKey);
      usedCashiers.add(cashierKey);
      usedBanks.add(bankKey);

      const isAmountDiff = Math.abs(bankRow.rawAmount - cashierRow.amount) > 0.01;
      const isPhoneDiff = hasPhoneDifference(cashierRow, bankRow);
      const cashierPhone = extractMobileNumber(`${cashierRow.rawName} ${cashierRow.name} ${cashierRow.ref}`);
      const bankPhone = extractMobileNumber(`${bankRow.description} ${bankRow.ref}`);
      const ms = advancedMatchCheck(cashierRow.name, bankRow.description);
      const isNameDiff = !(cashierPhone && cashierPhone === bankPhone) && (ms.isApprox || ms.matchType === "none");
      const isAccountTypeDiff = accountTypesDiffer(bankRow.accountType, cashierRow.accountType);
      const isDateDiff = transactionDatesDiffer(bankRow.date, cashierRow.date);
      const autoNote = [
        isPhoneDiff ? `الجوال في الفاتورة (${cashierPhone || "غير واضح"}) مختلف عن الجوال في الحركة (${bankPhone || "غير واضح"})` : "",
        isAmountDiff ? `اختلاف مبلغ: ${fmtNum(Math.abs(bankRow.rawAmount - cashierRow.amount))}` : "",
        isNameDiff ? `اختلاف اسم: ${cashierRow.name} ↔ ${bankRow.description}` : "",
        isAccountTypeDiff ? `نوع الحساب مختلف: بنك (${bankRow.accountType || "—"}) ↔ كاشير (${cashierRow.accountType || "—"})` : "",
        isDateDiff ? `تاريخ الفاتورة (${cashierRow.date}) مختلف عن تاريخ الحوالة (${bankRow.date})` : ""
      ].filter(Boolean).join(" | ") || undefined;

      newSaved.push({
        id: `saved-${Date.now()}-${cashierRow.id}-${bankRow.id}`,
        cashierId: cashierRow.id, bankId: bankRow.id,
        cashierSessionId: cashierRow._fileSessionId ?? 0, bankSessionId: bankRow._fileSessionId ?? 0,
        cashierDefaultRef: cashierRow.defaultRef || transactionKey(cashierRow),
        bankDefaultRef: bankRow.defaultRef || transactionKey(bankRow),
        cashierName: cashierRow.name, bankDesc: bankRow.description,
        cashierIdentity: movementIdentity(cashierRow), bankIdentity: movementIdentity(bankRow),
        cashierDate: cashierRow.date, bankDate: bankRow.date,
        cashierAmount: cashierRow.amount, cashierRawName: cashierRow.rawName,
        amount: bankRow.rawAmount, type: cashierRow.type,
        date: new Date().toLocaleDateString("ar-SA"),
        savedAt: new Date().toLocaleString("ar-SA"),
        note: autoNote, isAmountDiff, isPhoneDiff, isNameDiff, isManual: true, isAccountTypeDiff, isDateDiff,
        matchScore: Math.round(nameSim(cashierRow.name, bankRow.description) * 100),
        editorNotes: undefined,
        bankAccountType: bankRow.accountType, cashierAccountType: cashierRow.accountType
      });
    });
    if (!newSaved.length) {
      if (skipped) setToast(`لم تُحفظ أي مطابقة. التعارضات: ${conflictDetails.slice(0, 3).join("؛ ")}${conflictDetails.length > 3 ? `؛ و${conflictDetails.length - 3} غيرها` : ""}.`);
      return;
    }
    onSafeActivities?.(newSaved.map(saved => `${saved.cashierName} ↔ ${saved.bankDesc} · ${fmtNum(saved.amount)} ₪`));
    pushUndoAction({ type: "match", snapshot: {
      manualGroups: [...manualGroups],
      savedMatches: [...savedMatches],
      rejectedPairs: new Set(rejectedPairs),
      nameAliases: { ...nameAliases },
      selectedPendingKeys: new Set(selectedPendingKeys),
    } });
    skipNextAutoReconcile.current = true;
    setSavedMatches(prev => dedupeSavedMatches([...prev, ...newSaved]));
    const groupCashierKeys = new Set(newSaved.flatMap(saved => {
      const row = pairs.find(({ cashier }) => cashier.id === saved.cashierId && (cashier._fileSessionId ?? 0) === (saved.cashierSessionId ?? 0))?.cashier;
      return row ? [transactionKey(row)] : [];
    }));
    const groupBankKeys = new Set(newSaved.flatMap(saved => {
      const row = pairs.find(({ bank }) => bank.id === saved.bankId && (bank._fileSessionId ?? 0) === (saved.bankSessionId ?? 0))?.bank;
      return row ? [transactionKey(row)] : [];
    }));
    setResults(prev => {
      if (!prev) return prev;
      const savedResults: MatchResult[] = newSaved.flatMap(saved => {
        const pair = pairs.find(({ cashier, bank }) =>
          cashier.id === saved.cashierId && bank.id === saved.bankId
          && (cashier._fileSessionId ?? 0) === (saved.cashierSessionId ?? 0)
          && (bank._fileSessionId ?? 0) === (saved.bankSessionId ?? 0)
        );
        return pair ? [{ type: "saved" as const, cashier: pair.cashier, bank: pair.bank, savedMatch: saved }] : [];
      });
      const next = prev.filter(result => {
        const resultCashierKey = "cashier" in result ? transactionKey(result.cashier) : "";
        const resultBankKey = "bank" in result ? transactionKey(result.bank) : "";
        return !groupCashierKeys.has(resultCashierKey) && !groupBankKeys.has(resultBankKey);
      });
      return [...savedResults, ...next];
    });
    const aliasUpdates: Record<string, string> = {};
    pairs.forEach(({ cashier, bank }) => {
      const ms = advancedMatchCheck(cashier.name, bank.description);
      if (ms.isApprox || ms.matchType === "none") {
        aliasUpdates[smartNameKey(cashier.name)] = smartNameKey(smartAliasTarget(cashier.name, bank.description));
      }
    });
    if (Object.keys(aliasUpdates).length) setNameAliases(prev => ({ ...prev, ...aliasUpdates }));
    setToast(`تم تأكيد ${newSaved.length} مطابقة يدوياً${skipped ? `. تم تجاوز ${duplicateSkipped}: ${conflictDetails.slice(0, 2).join("؛ ")}${conflictDetails.length > 2 ? "؛ وغيرها" : ""}` : ""}`);
    setSelectedPendingKeys(new Set());
  };

  const handleUnsaveMatch = (savedMatch: SavedMatch) => {
    if (!window.confirm("هل أنت متأكد من إلغاء حفظ هذه المطابقة؟")) return;
    clearMatchUndo();
    if (savedMatch.sourceGroupId) {
      setSavedMatches(prev => prev.filter(s => s.sourceGroupId !== savedMatch.sourceGroupId));
      setManualGroups(prev => prev.filter(g => g.id !== savedMatch.sourceGroupId));
    } else {
      setSavedMatches(prev => prev.filter(s => s.id !== savedMatch.id));
    }
    setSelectedSavedIds(prev => {
      const next = new Set(prev);
      next.delete(savedMatch.id);
      return next;
    });
  };

  const handleUnsaveSelected = () => {
    const selected = new Set(selectedSavedIds);
    if (!selected.size) return;
    const selectedMatches = savedMatches.filter(match => selected.has(match.id));
    if (!window.confirm(`إلغاء تأكيد ${selectedMatches.length} مطابقة محددة؟`)) return;
    clearMatchUndo();
    const groupIds = new Set(selectedMatches.flatMap(match => match.sourceGroupId ? [match.sourceGroupId] : []));
    setSavedMatches(prev => prev.filter(match =>
      !selected.has(match.id) && (!match.sourceGroupId || !groupIds.has(match.sourceGroupId))
    ));
    if (groupIds.size) {
      setManualGroups(prev => prev.filter(group => !groupIds.has(group.id)));
    }
    setSelectedSavedIds(new Set());
    setExpandedMatchKey(null);
    setToast(`تم إلغاء تأكيد ${selectedMatches.length} مطابقة`);
  };

  const handleRejectMatch = (cashierRow: CashierRow, bankRow: BankRow) => {
    const pairKey = transactionPairKey(cashierRow, bankRow);
    clearMatchUndo();
    const next = new Set(rejectedPairs);
    next.add(pairKey);
    setRejected(next);
    setExpandedMatchKey(null);
    setSavedMatches(prev => prev.filter(s =>
      !(s.cashierId === cashierRow.id && s.bankId === bankRow.id
        && (s.cashierSessionId == null || s.cashierSessionId === (cashierRow._fileSessionId ?? 0))
        && (s.bankSessionId == null || s.bankSessionId === (bankRow._fileSessionId ?? 0)))
    ));
    setSelectedPendingKeys(prev => { const n = new Set(prev); n.delete(pairKey); return n; });
  };

  const handleReplaceBank = (cashierRow: CashierRow, oldBank: BankRow, newBank: BankRow) => {
    const newSaved = handleSaveMatch(cashierRow, newBank, undefined, { force: true, replacePair: { cashier: cashierRow, bank: oldBank } });
    if (!newSaved) {
      setToast("هذه الحوالة مرتبطة مسبقاً ولا يمكن استخدامها للاستبدال.");
      return;
    }
    finishReplacement(cashierRow, oldBank, cashierRow, newBank, newSaved);
    closeDrawer();
  };

  const drawerBanks = useMemo(() => {
    if (drawerMode !== "bank" || !drawerCashier) return [];
    return reconciliationBankRows
      .filter(b => {
        if (b.type !== drawerCashier.type) return false;
        const pairKey = `${transactionKey(drawerCashier)}-${transactionKey(b)}`;
        if (savedBankIds.has(transactionKey(b)) && !savedKeys.has(pairKey)) return false;
        if (b.id === drawerOldBank?.id) return false;
        const nameMatch = !drawerNameSearch.trim() ||
          b.description.toLowerCase().includes(drawerNameSearch.trim().toLowerCase());
        const amt = b.rawAmount;
        const fromOk = !drawerAmountFrom || amt >= Number(drawerAmountFrom);
        const toOk = !drawerAmountTo || amt <= Number(drawerAmountTo);
        return nameMatch && fromOk && toOk;
      })
      .sort((a, b) => {
        const diffA = Math.abs(a.rawAmount - drawerCashier.matchAmount);
        const diffB = Math.abs(b.rawAmount - drawerCashier.matchAmount);
        if (diffA !== diffB) return diffA - diffB;
        return nameSim(drawerCashier.name, b.description) - nameSim(drawerCashier.name, a.description);
      })
      .slice(0, 50);
  }, [drawerMode, drawerCashier, drawerNameSearch, drawerAmountFrom, drawerAmountTo, reconciliationBankRows, savedBankIds, savedKeys, drawerOldBank]);

  const drawerCashiers = useMemo(() => {
    if (drawerMode !== "cashier" || !drawerOldBank) return [];
    return reconciliationCashierRows
      .filter(c => {
        if (c.type !== drawerOldBank.type) return false;
        if (visaCashierIds.has(transactionKey(c))) return false;
        const pairKey = `${transactionKey(c)}-${transactionKey(drawerOldBank)}`;
        if (savedCashierIds.has(transactionKey(c)) && !savedKeys.has(pairKey)) return false;
        if (c.id === drawerCashier?.id) return false;
        const nameMatch = !drawerNameSearch.trim() ||
          c.name.toLowerCase().includes(drawerNameSearch.trim().toLowerCase());
        const amt = c.matchAmount;
        const fromOk = !drawerAmountFrom || amt >= Number(drawerAmountFrom);
        const toOk = !drawerAmountTo || amt <= Number(drawerAmountTo);
        return nameMatch && fromOk && toOk;
      })
      .sort((a, b) => {
        const diffA = Math.abs(a.matchAmount - drawerOldBank.rawAmount);
        const diffB = Math.abs(b.matchAmount - drawerOldBank.rawAmount);
        if (diffA !== diffB) return diffA - diffB;
        return nameSim(drawerOldBank.description, b.name) - nameSim(drawerOldBank.description, a.name);
      })
      .slice(0, 50);
  }, [drawerMode, drawerOldBank, drawerCashier, drawerNameSearch, drawerAmountFrom, drawerAmountTo, reconciliationCashierRows, savedCashierIds, savedKeys, visaCashierIds]);

  const openDrawer = (cashier: CashierRow, oldBank: BankRow, mode: "bank" | "cashier" = "bank") => {
    setDrawerMode(mode);
    setDrawerCashier(cashier);
    setDrawerOldBank(oldBank);
    setDrawerNameSearch("");
    setDrawerAmountFrom("");
    setDrawerAmountTo("");
    setDrawerOpen(true);
    const pairKey = transactionPairKey(cashier, oldBank);
    window.setTimeout(() => {
      pendingRowRefs.current[pairKey]?.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
    }, 50);
  };

  const closeDrawer = () => {
    setDrawerOpen(false);
    setDrawerCashier(null);
    setDrawerOldBank(null);
    setDrawerNameSearch("");
    setDrawerAmountFrom("");
    setDrawerAmountTo("");
    setDraggedItemId(null);
    setDropTargetKey(null);
  };

  useEffect(() => {
    if (!drawerOpen) return;
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") closeDrawer(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [drawerOpen]);

  const handleDropBank = (cashier: CashierRow, oldBank: BankRow, newBankId: number) => {
    const newBank = reconciliationBankRows.find(b => b.id === newBankId);
    if (!newBank) return;
    handleReplaceBank(cashier, oldBank, newBank);
  };

  const handleReplaceCashier = (bankRow: BankRow, oldCashier: CashierRow, newCashier: CashierRow) => {
    const newSaved = handleSaveMatch(newCashier, bankRow, undefined, { force: true, replacePair: { cashier: oldCashier, bank: bankRow } });
    if (!newSaved) {
      setToast("هذه الفاتورة مرتبطة مسبقاً ولا يمكن استخدامها للاستبدال.");
      return;
    }
    finishReplacement(oldCashier, bankRow, newCashier, bankRow, newSaved);
    closeDrawer();
  };

  const handleDropCashier = (bankRow: BankRow, oldCashier: CashierRow, newCashierId: number) => {
    const newCashier = reconciliationCashierRows.find(c => c.id === newCashierId);
    if (!newCashier) return;
    handleReplaceCashier(bankRow, oldCashier, newCashier);
  };

  const handleAcceptSuggestion = (cashierRow: CashierRow, bankRow: BankRow) => {
    handleSaveMatch(cashierRow, bankRow);
    setExpandedUnmatchedCashier(null);
  };

  const handleRejectSuggestion = (cashierRow: CashierRow, bankRow: BankRow) => {
    const pairKey = transactionPairKey(cashierRow, bankRow);
    clearMatchUndo();
    const next = new Set(rejectedPairs);
    next.add(pairKey);
    setRejected(next);
    setExpandedUnmatchedCashier(cashierRow.id);
  };

  const handleAddGroup = (g: ManualMatchGroup): boolean => {
    const banks = uniqueTransactionRows(g.banks);
    const cashiers = uniqueTransactionRows(g.cashiers);
    if (banks.length !== 1 || cashiers.length !== 1) {
      setToast("لمنع تكرار الحركات، المطابقة اليدوية تقبل حوالة واحدة وفاتورة واحدة فقط. استخدم أداة التجزئة لتوزيع حوالة على عدة فواتير.");
      return false;
    }
    const bank = banks[0];
    const cashier = cashiers[0];
    const bankKey = transactionKey(bank);
    const cashierKey = transactionKey(cashier);
    if (savedBankIds.has(bankKey) || savedCashierIds.has(cashierKey)) {
      setToast("هذه الفاتورة أو الحوالة مستخدمة في مطابقة محفوظة بالفعل.");
      return false;
    }
    const amountDiff = Math.abs(bank.rawAmount - cashier.matchAmount);
    const isPhoneDiff = hasPhoneDifference(cashier, bank);
    const { cashierPhone, bankPhone } = movementPhones(cashier, bank);
    const dateDiff = transactionDatesDiffer(bank.date, cashier.date);
    const accountTypeDiff = accountTypesDiffer(bank.accountType, cashier.accountType);
    const differenceNotes = [
      g.note,
      isPhoneDiff ? `الرقم بالفاتورة ${cashierPhone ?? "غير موجود"} مختلف؛ تعديل الرقم إلى ${bankPhone ?? "الرقم الموجود في الحوالة"}.` : "",
      amountDiff > 0.01 ? `فرق المبلغ ${fmtNum(amountDiff)}.` : "",
      accountTypeDiff ? `نوع الحساب مختلف: بنك (${bank.accountType || "—"}) ↔ كاشير (${cashier.accountType || "—"}).` : "",
      dateDiff ? `تاريخ الفاتورة (${cashier.date}) مختلف عن تاريخ الحوالة (${bank.date}).` : "",
    ].filter(Boolean);
    const savedNote = differenceNotes.join(" ");
    pushUndoAction({ type: "match", snapshot: {
      manualGroups: [...manualGroups],
      savedMatches: [...savedMatches],
      rejectedPairs: new Set(rejectedPairs),
      nameAliases: { ...nameAliases },
      selectedPendingKeys: new Set(selectedPendingKeys),
    } });
    const newSaved: SavedMatch[] = [{
          id: `manual-${g.id}-${cashier.id}-${bank.id}`,
          cashierId: cashier.id, bankId: bank.id,
          cashierSessionId: cashier._fileSessionId ?? 0, bankSessionId: bank._fileSessionId ?? 0,
          cashierDefaultRef: cashier.defaultRef || transactionKey(cashier),
          bankDefaultRef: bank.defaultRef || transactionKey(bank),
          cashierName: cashier.name, bankDesc: bank.description,
          cashierIdentity: movementIdentity(cashier), bankIdentity: movementIdentity(bank),
          cashierDate: cashier.date, bankDate: bank.date,
          cashierAmount: cashier.amount, cashierRawName: cashier.rawName,
          amount: bank.rawAmount, type: cashier.type,
          date: new Date().toLocaleDateString("ar-SA"),
          savedAt: new Date().toLocaleString("ar-SA"),
          note: savedNote || "مطابقة يدوية",
          isAmountDiff: amountDiff > 0.01,
          isPhoneDiff,
          isNameDiff: true, isManual: true,
          isAccountTypeDiff: accountTypeDiff,
          isDateDiff: dateDiff,
          bankAccountType: bank.accountType, cashierAccountType: cashier.accountType,
          editorNotes: g.note,
          sourceGroupId: g.id,
          matchScore: Math.round(nameSim(cashier.name, bank.description) * 100)
        }];
    if (newSaved.length) {
      onSafeActivities?.(newSaved.map(saved => `${saved.cashierName} ↔ ${saved.bankDesc} · ${fmtNum(saved.amount)} ₪`));
    }
    setSavedMatches(prev => dedupeSavedMatches([...prev, ...newSaved]));
    const groupPairKeys = new Set([transactionPairKey(cashier, bank)]);
    setResults(prev => prev
      ? [
          { type: "saved", cashier, bank, savedMatch: newSaved[0] },
          ...prev.filter(result => {
            const sameCashier = "cashier" in result && transactionKey(result.cashier) === cashierKey;
            const sameBank = "bank" in result && transactionKey(result.bank) === bankKey;
            return !sameCashier && !sameBank;
          })
        ]
      : prev);
    setManualGroups(prev => [...prev, g]);
    setSelectedPendingKeys(prev => { const n = new Set(prev); groupPairKeys.forEach(k => n.delete(k)); return n; });
    return true;
  };

  const handleRemoveGroup = (id: string) => {
    clearMatchUndo();
    setSavedMatches(prev => prev.filter(s => s.sourceGroupId !== id));
    setManualGroups(prev => prev.filter(g => g.id !== id));
  };

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (!(event.ctrlKey || event.metaKey) || event.code !== "KeyZ" || event.shiftKey || !canUndoMatch) return;
      if (isTextEntryTarget(event.target)) return;
      event.preventDefault();
      handleUndoMatch();
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [canUndoMatch]);

  const handleHoldCashier = (c: CashierRow, note?: string) => {
    const fileSessionId = c._fileSessionId ?? cashFileSessionId;
    if (heldItems.some(h => h.kind === "cashier" && h.refId === c.id && h.fileSessionId === fileSessionId)) {
      setToast("هذه الفاتورة موجودة بالفعل في المعلقات.");
      return;
    }
    if ((c as any)._fromHeld) {
      setReturnedHeldCashier(prev => prev.filter(x =>
        (x._fileSessionId ?? cashFileSessionId) !== fileSessionId ||
        (x.id !== c.id && !sameCashierMovement(x, c))
      ));
    }
    const item: HeldItem = {
      id: `held-c-${Date.now()}-${c.id}`, kind: "cashier", refId: c.id,
      fileSessionId, data: c,
      heldAt: new Date().toLocaleString("ar-SA"), note
    };
    setHeldItems(prev => dedupeHeldItems([...prev, item]));
    setExpandedMatchKey(null);
    setExpandedUnmatchedCashier(null);
  };
  const handleHoldBank = (b: BankRow, note?: string) => {
    const fileSessionId = b._fileSessionId ?? bankFileSessionId;
    if (heldItems.some(h => h.kind === "bank" && h.refId === b.id && h.fileSessionId === fileSessionId)) {
      setToast("هذه الحوالة موجودة بالفعل في المعلقات.");
      return;
    }
    if ((b as any)._fromHeld) {
      setReturnedHeldBank(prev => prev.filter(x =>
        (x._fileSessionId ?? bankFileSessionId) !== fileSessionId ||
        (x.id !== b.id && !sameBankMovement(x, b))
      ));
    }
    const item: HeldItem = {
      id: `held-b-${Date.now()}-${b.id}`, kind: "bank", refId: b.id,
      fileSessionId, data: b,
      heldAt: new Date().toLocaleString("ar-SA"), note
    };
    setHeldItems(prev => dedupeHeldItems([...prev, item]));
    setExpandedMatchKey(null);
  };
  const handleUnhold = (h: HeldItem) => {
    if (h.kind === "bank") {
      const row: BankRow = { ...(h.data as BankRow), _fileSessionId: h.fileSessionId, _fromHeld: true };
      setReturnedHeldBank(prev => prev.some(item => item.id === row.id && item._fileSessionId === row._fileSessionId) ? prev : [...prev, row]);
    } else {
      const row: CashierRow = { ...(h.data as CashierRow), _fileSessionId: h.fileSessionId, _fromHeld: true };
      setReturnedHeldCashier(prev => prev.some(item => item.id === row.id && item._fileSessionId === row._fileSessionId) ? prev : [...prev, row]);
    }
    setHeldItems(prev => prev.filter(x => x.id !== h.id));
  };
  const handleDeleteHeld = (id: string) => {
    if (!window.confirm("حذف هذا العنصر نهائياً من المعلقات؟")) return;
    setHeldItems(prev => prev.filter(h => h.id !== id));
  };
  const handleSaveBankSplit = (split: BankSplit): boolean => {
    const bank = matchableBank.find(row => transactionKey(row) === `${split.bankSessionId}:${split.bankId}`);
    const bankAlreadySplit = bankSplits.some(existing =>
      existing.id !== split.id && existing.bankId === split.bankId && existing.bankSessionId === split.bankSessionId
    );
    if (!bank || savedBankIds.has(transactionKey(bank)) || bankAlreadySplit) {
      setToast("تعذّر حفظ التجزئة: الحوالة لم تعد متاحة. حدّث الأداة وحاول مجدداً.");
      return false;
    }
    const allocationTotal = split.allocations.reduce((sum, allocation) => sum + allocation.amount, 0);
    if (!split.allocations.length || Math.abs(allocationTotal - bank.rawAmount) > 0.01) {
      setToast("تعذّر حفظ التجزئة: مجموع الحصص لا يساوي قيمة الحوالة الأصلية.");
      return false;
    }
    const linkedMatches: SavedMatch[] = [];
    const allocatedByCashier = new Map<string, { row: CashierRow; total: number; allocations: SplitAllocation[] }>();
    const allocatedCashierKeys = new Set<string>();
    for (const allocation of split.allocations) {
      if (allocation.amount <= 0 || (allocation.kind === "external" && !allocation.note.trim())) {
        setToast("تعذّر حفظ التجزئة: كل حصة يجب أن تكون موجبة والحصص الخارجية تحتاج إلى توضيح.");
        return false;
      }
      if (allocation.kind !== "company" || allocation.cashierId == null) continue;
      const cashier = activeCashier.find(row =>
        row.id === allocation.cashierId && (row._fileSessionId ?? 0) === (allocation.cashierSessionId ?? 0)
      );
      if (!cashier || savedCashierIds.has(transactionKey(cashier)) || cashier.type !== bank.type) {
        setToast("تعذّر حفظ التجزئة: إحدى الفواتير لم تعد متاحة أو تغيّرت قيمتها.");
        return false;
      }
      const cashierKey = transactionKey(cashier);
      if (allocatedCashierKeys.has(cashierKey)) {
        setToast("تعذّر حفظ التجزئة: لا يمكن إضافة الفاتورة نفسها أكثر من مرة؛ اجمع حصتها في سطر واحد.");
        return false;
      }
      allocatedCashierKeys.add(cashierKey);
      const allocated: { row: CashierRow; total: number; allocations: SplitAllocation[] } =
        allocatedByCashier.get(cashierKey) ?? { row: cashier, total: 0, allocations: [] };
      allocated.total += allocation.amount;
      allocated.allocations.push(allocation);
      allocatedByCashier.set(cashierKey, allocated);
    }
    for (const [cashierKey, { row, total, allocations }] of allocatedByCashier) {
      if (Math.abs(row.matchAmount - total) > 0.01) {
        setToast(`تعذّر حفظ التجزئة: مجموع الحصص المخصصة إلى «${row.name}» لا يساوي مبلغ السند.`);
        return false;
      }
      const virtualBankId = bankSplitVirtualId(split, cashierKey, allocations);
      if (savedMatches.some(match => match.cashierId === row.id && match.bankId === virtualBankId)) {
        setToast("إحدى حصص هذه التجزئة مرتبطة بمطابقة سابقة؛ لم يتم حفظ التغيير.");
        return false;
      }
      linkedMatches.push({
        id: `split-match-${split.id}-${cashierKey.replace(":", "-")}`,
        cashierId: row.id,
        cashierSessionId: row._fileSessionId ?? 0,
        bankId: virtualBankId,
        bankSessionId: bank._fileSessionId ?? 0,
        cashierDefaultRef: row.defaultRef || transactionKey(row),
        bankDefaultRef: bank.defaultRef || transactionKey(bank),
        cashierName: row.name,
        bankDesc: bank.description,
        cashierIdentity: movementIdentity(row),
        bankIdentity: movementIdentity({
          ...bank,
          id: virtualBankId,
          rawAmount: total,
          debit: bank.type === "مستلم" ? total : 0,
          credit: bank.type === "مدفوع" ? total : 0,
        }),
        cashierDate: row.date,
        bankDate: bank.date,
        cashierAmount: row.amount,
        cashierRawName: row.rawName,
        amount: total,
        type: row.type,
        date: new Date().toLocaleDateString("ar-SA"),
        savedAt: new Date().toLocaleString("ar-SA"),
        note: allocations.map(allocation => allocation.note).filter(Boolean).join(" · ") || "حصة من حوالة مجزأة",
        isAmountDiff: false,
        isNameDiff: true,
        isManual: true,
        isAccountTypeDiff: accountTypesDiffer(bank.accountType, row.accountType),
        bankAccountType: bank.accountType,
        cashierAccountType: row.accountType,
        editorNotes: `تجزئة حوالة ${split.description} · الحصص ${allocations.map(allocation => split.allocations.indexOf(allocation) + 1).join(", ")}`,
      });
    }
    if (!linkedMatches.length) {
      setToast("يجب ربط حصة واحدة على الأقل بفاتورة تخص شركتكم.");
      return false;
    }
    setBankSplits(previous => [...previous, split]);
    setSavedMatches(previous => dedupeSavedMatches([...previous, ...linkedMatches]));
    setResults(null);
    setToast(`تم حفظ التجزئة وربط ${linkedMatches.length} سنداً بالفواتير.`);
    return true;
  };
  const handleDeleteBankSplit = (split: BankSplit) => {
    const childIds = new Set(Array.from(bankSplitAllocationGroups(split), ( [key, allocations]) => bankSplitVirtualId(split, key, allocations)));
    const linked = savedMatches.filter(match => childIds.has(match.bankId));
    if (!window.confirm(linked.length
      ? `إلغاء هذه التجزئة وإزالة ${linked.length} مطابقة مرتبطة بها؟`
      : "إلغاء هذه التجزئة وإعادة الحوالة الأصلية إلى قائمة المطابقة؟")) return;
    setSavedMatches(previous => previous.filter(match => !childIds.has(match.bankId)));
    setManualGroups(previous => previous.filter(group => !group.banks.some(bank => childIds.has(bank.id))));
    setBankSplits(previous => previous.filter(item => item.id !== split.id));
    setResults(null);
    setToast("تم إلغاء التجزئة وإعادة الحوالة الأصلية.");
  };
  const handleSaveCutoffBatch = (batch: CutoffBatch): boolean => {
    const unavailable = new Set([
      ...savedCashierIds,
      ...heldItems.filter(item => item.kind === "cashier").map(item => `${item.fileSessionId}:${item.refId}`),
    ]);
    const target = activeCashier.find(row =>
      row.id === batch.aggregateId && (row._fileSessionId ?? 0) === batch.aggregateSessionId
    );
    const sources = batch.sourceRows.map(reference => activeCashier.find(row =>
      row.id === reference.id && (row._fileSessionId ?? 0) === reference.sessionId
    ));
    const sourceKeys = batch.sourceRows.map(reference => `${reference.sessionId}:${reference.id}`);
    const uniqueSourceKeys = new Set(sourceKeys);
    if (!target || sources.some(row => !row) || uniqueSourceKeys.size !== sourceKeys.length ||
        sourceKeys.includes(transactionKey(target)) ||
        unavailable.has(transactionKey(target)) ||
        sources.some(row => row && (unavailable.has(transactionKey(row)) || row.type !== target.type)) ||
        sources.some((row): row is CashierRow => !!row && cutoffBatches.some(existing =>
          (existing.aggregateId === row.id && existing.aggregateSessionId === (row._fileSessionId ?? 0)) ||
          existing.sourceRows.some(reference => reference.id === row.id && reference.sessionId === (row._fileSessionId ?? 0))
        ))) {
      setToast("تعذّر حفظ المجموعة: تغيّرت إحدى الحركات أو لم تعد متاحة للمطابقة.");
      return false;
    }
    const total = sources.reduce((sum, row) => sum + (row?.matchAmount ?? 0), 0);
    if (Math.abs(total - target.matchAmount) > 0.01 || Math.abs(total - batch.total) > 0.01) {
      setToast("تعذّر حفظ المجموعة: مجموع الفواتير لم يعد مساوياً لسند الإجمالي.");
      return false;
    }
    setCutoffBatches(previous => [...previous, batch]);
    setResults(null);
    setToast(`تم إنشاء مجموعة فترة التعطل من ${sources.length} فاتورة بمبلغ ${fmtNum(total)}.`);
    return true;
  };
  const handleDeleteCutoffBatch = (batch: CutoffBatch) => {
    const linked = savedMatches.filter(match => match.cashierId === batch.syntheticId);
    if (!window.confirm(linked.length
      ? `إلغاء المجموعة وإزالة ${linked.length} مطابقة مرتبطة بها؟`
      : "إلغاء المجموعة وإعادة الحركات الأصلية إلى المطابقة؟")) return;
    setSavedMatches(previous => previous.filter(match => match.cashierId !== batch.syntheticId));
    setManualGroups(previous => previous.filter(group => !group.cashiers.some(row => row.id === batch.syntheticId)));
    setCutoffBatches(previous => previous.filter(item => item.id !== batch.id));
    setResults(null);
    setToast("تم إلغاء مجموعة فترة التعطل.");
  };
  const handleSaveClearingGroup = (group: ClearingGroup): boolean => {
    const allRefs = [...group.debitRows, ...group.creditRows];
    const selectedKeys = allRefs.map(reference => `${reference.sessionId}:${reference.id}`);
    const uniqueKeys = new Set(selectedKeys);
    const rows = allRefs.map(reference => activeCashier.find(row =>
      row.id === reference.id && (row._fileSessionId ?? 0) === reference.sessionId
    ));
    const unavailableKeys = new Set<string>([
      ...activeCashier.filter(row => savedCashierIds.has(transactionKey(row))).map(transactionKey),
      ...heldItems.filter(item => item.kind === "cashier").map(item => `${item.fileSessionId}:${item.refId}`),
      ...manualGroups.flatMap(group => group.cashiers.map(transactionKey)),
      ...cutoffBatches.flatMap(batch => [
        `${batch.aggregateSessionId}:${batch.aggregateId}`,
        ...batch.sourceRows.map(reference => `${reference.sessionId}:${reference.id}`),
      ]),
      ...bankSplits.flatMap(split => split.allocations
        .filter(allocation => allocation.kind === "company" && allocation.cashierId != null)
        .map(allocation => `${allocation.cashierSessionId ?? 0}:${allocation.cashierId}`)),
      ...clearedCashierKeys,
    ]);
    if (allRefs.length < 2 || uniqueKeys.size !== allRefs.length || rows.some(row => !row) ||
        selectedKeys.some(key => unavailableKeys.has(key))) {
      setToast("تعذّر التسكير: إحدى الحركات مكررة أو لم تعد متاحة.");
      return false;
    }
    const resolved = rows as CashierRow[];
    const debitRows = resolved.slice(0, group.debitRows.length);
    const creditRows = resolved.slice(group.debitRows.length);
    if (new Set(resolved.map(row => row.accountType.trim()).filter(Boolean)).size > 1) {
      setToast("تعذّر التسكير: لا يمكن جمع حركات من حسابات أو منصات مختلفة.");
      return false;
    }
    if (debitRows.some(row => row.debit <= 0 || row.credit > 0) ||
        creditRows.some(row => row.credit <= 0 || row.debit > 0)) {
      setToast("تعذّر التسكير: تأكد أن الطرف الأول مدين بالكامل والثاني دائن بالكامل.");
      return false;
    }
    const debitTotal = debitRows.reduce((sum, row) => sum + row.debit, 0);
    const creditTotal = creditRows.reduce((sum, row) => sum + row.credit, 0);
    if (Math.abs(debitTotal - creditTotal) > 0.01 || Math.abs(debitTotal - group.total) > 0.01) {
      setToast("تعذّر التسكير: مجموع المدين لا يساوي مجموع الدائن.");
      return false;
    }
    setClearingGroups(previous => [...previous, group]);
    setResults(null);
    setToast(`تم تسكير ${debitRows.length + creditRows.length} حركة بمبلغ ${fmtNum(debitTotal)}.`);
    return true;
  };
  const handleDeleteClearingGroup = (group: ClearingGroup) => {
    if (!window.confirm("إلغاء التسكير وإعادة الحركات إلى قائمة المراجعة؟")) return;
    setClearingGroups(previous => previous.filter(item => item.id !== group.id));
    setResults(null);
    setToast("تم إلغاء التسكير وإعادة الحركات المفتوحة.");
  };

  useEffect(() => {
    if (skipNextAutoReconcile.current) {
      skipNextAutoReconcile.current = false;
      return;
    }
    if (reconciliationBankRows.length && reconciliationCashierRows.length) {
      rerun();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [manualGroups, savedMatches, rejectedPairs, visaItems, jawwalPayCashierIds, mahmoudWalletCashierIds, rejectedSpecialCashierIds, heldItems, bankSplits, cutoffBatches, clearingGroups, reconciliationBankRows, reconciliationCashierRows, amountTolerancePercent]);

  const stats = useMemo(() => {
    if (!results) return null;
    return {
      saved: results.filter(r => r.type === "saved").length,
      pending: results.filter(r => r.type === "pending").length,
      manual: results.filter((r:any) => r.type === "saved" && (r.savedMatch as SavedMatch).isManual).length,
      review: reviewItems.length,
      visa: visaItems.length,
      jawwalPay: results.filter(r => r.type === "jawwalPay").length,
      mahmoudWallet: results.filter(r => r.type === "mahmoudWallet").length,
      held: heldItems.length,
      nonReconciliationBank: nonReconciliationBankRows.length,
      uCashier: results.filter(r => r.type === "unmatchedCashier").length,
      uBank: results.filter(r => r.type === "unmatchedBank").length,
    };
  }, [results, visaItems, heldItems, nonReconciliationBankRows, reviewItems]);

  const filterResult = (row: any) => {
    const cashier = row.cashier as CashierRow | undefined;
    const bank = row.bank as BankRow | undefined;
    const text = `${cashier?.name || ""} ${bank?.description || ""} ${cashier ? defaultReferenceForRow(cashier) : ""} ${bank ? defaultReferenceForRow(bank) : ""}`.toLowerCase();
    const amount = bank?.rawAmount ?? cashier?.amount ?? 0;
    return (!resultSearch.trim() || text.includes(resultSearch.trim().toLowerCase())) &&
      (resultType === "all" || cashier?.type === resultType || bank?.type === resultType) &&
      (!minAmount || amount >= Number(minAmount)) && (!maxAmount || amount <= Number(maxAmount));
  };
  const visibleSavedRows = savedRows.filter(filterResult);
  const visiblePendingRows = pendingRows.filter(filterResult);
  const pagedPendingRows = useMemo(() => {
    const totalPages = Math.max(1, Math.ceil(visiblePendingRows.length / PENDING_PAGE_SIZE));
    const page = Math.min(pendingPage, totalPages);
    const start = (page - 1) * PENDING_PAGE_SIZE;
    return visiblePendingRows.slice(start, start + PENDING_PAGE_SIZE);
  }, [visiblePendingRows, pendingPage]);
  useEffect(() => {
    const row = pagedPendingRows[Math.min(keyboardPendingIndex, pagedPendingRows.length - 1)] as { cashier: CashierRow; bank: BankRow } | undefined;
    if (row) pendingRowRefs.current[transactionPairKey(row.cashier, row.bank)]?.scrollIntoView({ block: "nearest" });
  }, [pagedPendingRows, keyboardPendingIndex]);
  const visibleUCashierRows = uCashierRows.filter(filterResult);
  const visibleUBankRows = uBankRows.filter(filterResult);
  const visibleNonReconciliationBankRows = nonReconciliationBankRows.filter(({ bank }) => filterResult({ bank }));
  const visibleFilteredNonReconciliationBankRows = visibleNonReconciliationBankRows.filter(({ reason }) =>
    nonReconBankTypeFilter === "all" ||
    (nonReconBankTypeFilter === "merchant" && reason === "عمولات المشتريات والتجار") ||
    (nonReconBankTypeFilter === "other" && reason !== "عمولات المشتريات والتجار")
  );
  const merchantHeldBankItems = heldItems.filter(item =>
    item.kind === "bank" && isMerchantPurchaseDescription((item.data as BankRow).description)
  );
  const handleMoveSelectedNonReconBankToHeld = () => {
    const selectedRows = visibleFilteredNonReconciliationBankRows
      .filter(({ bank }) => selectedNonReconBankKeys.has(transactionKey(bank)));
    if (!selectedRows.length) return;
    selectedRows.forEach(({ bank, reason }) => handleHoldBank(bank, reason));
    setSelectedNonReconBankKeys(new Set());
  };
  const handleTransferHeldMerchantToC = (items: HeldItem[]) => {
    const bankRows = items
      .filter(item => item.kind === "bank" && isMerchantPurchaseDescription((item.data as BankRow).description))
      .map(item => {
        const bank = item.data as BankRow;
        return { id: item.id, description: bank.description, amount: bank.rawAmount, date: bank.date };
      });
    if (!bankRows.length) return;
    setStageDBankInputs(bankRows);
    setPage("recon2Visa");
  };

  const handleOpenReference = (candidate: { id: number; _fileSessionId?: number; defaultRef?: string; rawAmount?: number }) => {
    const key = transactionKey(candidate);
    const isBank = typeof candidate.rawAmount === "number";
    const row = isBank
      ? [...bankSourceRows, ...returnedHeldBank].find(item => transactionKey(item) === key)
      : [...cashierSourceRows, ...returnedHeldCashier].find(item => transactionKey(item) === key);
    if (!row) {
      setError("تعذّر تحديد الصف الأصلي لهذا الرقم. حدّث الملفات أو استرجع جلسة التسوية ثم حاول مجدداً.");
      return;
    }

    const isResolvedBank = "rawAmount" in row;
    const locations: ReferenceLocation[] = [];
    const addLocation = (location: ReferenceLocation) => {
      if (!locations.some(item => item.id === location.id)) locations.push(location);
    };
    const result = results?.find(item => {
      if (item.type === "manualGroup") {
        return [...item.group.banks, ...item.group.cashiers].some(itemRow => transactionKey(itemRow) === key);
      }
      return ("bank" in item && transactionKey(item.bank) === key)
        || ("cashier" in item && transactionKey(item.cashier) === key);
    });

    if (result) {
      if (result.type === "saved") {
        addLocation({ id: "main-saved", label: "المطابقة العامة · المؤكدات", detail: `${result.cashier.name} ↔ ${result.bank.description}`, page: "main", tab: "saved" });
      } else if (result.type === "pending") {
        addLocation({ id: "main-pending", label: "المطابقة العامة · المنتظر", detail: `${result.cashier.name} ↔ ${result.bank.description}`, page: "main", tab: "pending" });
      } else if (result.type === "manualGroup") {
        addLocation({ id: "main-manual", label: "المطابقة العامة · الربط اليدوي", detail: "موجود ضمن مجموعة مطابقة يدوية", page: "main", tab: "manual" });
      } else if (result.type === "unmatchedCashier") {
        addLocation({ id: "main-u-cashier", label: "المطابقة العامة · فواتير بلا حوالات", detail: result.cashier.name, page: "main", tab: "uCashier" });
      } else if (result.type === "unmatchedBank") {
        addLocation({ id: "main-u-bank", label: "المطابقة العامة · حوالات بلا فواتير", detail: result.bank.description, page: "main", tab: "uBank" });
      } else if (result.type === "mahmoudWallet") {
        addLocation({ id: "main-mahmoud", label: "المطابقة العامة · محفظة شخصية", detail: result.cashier.name, page: "main", tab: "mahmoudWallet" });
      } else if (result.type === "jawwalPay") {
        addLocation({ id: "main-jawwal", label: "المطابقة العامة · جوال بي", detail: result.cashier.name, page: "main", tab: "jawwalPay" });
      } else if (result.type === "visa") {
        addLocation({ id: "main-visa", label: "المطابقة العامة · الفيزا", detail: result.cashier.name, page: "main", tab: "visa" });
      }
    }

    if (heldItems.some(item => item.kind === (isResolvedBank ? "bank" : "cashier") && `${item.fileSessionId}:${item.refId}` === key)) {
      addLocation({ id: "main-held", label: "المطابقة العامة · المعلّقات", detail: isResolvedBank ? (row as BankRow).description : (row as CashierRow).name, page: "main", tab: "held" });
    }
    if (reviewItems.some(item =>
      (item.cashier && transactionKey(item.cashier) === key) || (item.bank && transactionKey(item.bank) === key)
    )) {
      addLocation({ id: "main-review", label: "المطابقة العامة · قائمة المراجعة", detail: "نسخة محفوظة للمراجعة", page: "main", tab: "review" });
    }

    const split = bankSplits.find(item =>
      (isResolvedBank && item.bankId === row.id && item.bankSessionId === (row._fileSessionId ?? 0))
      || (!isResolvedBank && item.allocations.some(allocation =>
        allocation.kind === "company" && allocation.cashierId === row.id && (allocation.cashierSessionId ?? 0) === (row._fileSessionId ?? 0)
      ))
    );
    if (split) addLocation({ id: "split-tool", label: "أداة تجزئة الحوالات", detail: `تجزئة محفوظة · ${split.description}`, page: "splitTool" });

    const cutoff = cutoffBatches.find(item =>
      (!isResolvedBank && item.aggregateId === row.id && item.aggregateSessionId === (row._fileSessionId ?? 0))
      || (!isResolvedBank && item.sourceRows.some(reference => reference.id === row.id && reference.sessionId === (row._fileSessionId ?? 0)))
    );
    if (cutoff) addLocation({ id: "cutoff-tool", label: "أداة تجميع فترة التعطل", detail: `مجموعة ${cutoff.id}`, page: "batchCutoff" });

    const clearing = !isResolvedBank && clearingGroups.find(group =>
      [...group.debitRows, ...group.creditRows].some(reference => reference.id === row.id && reference.sessionId === (row._fileSessionId ?? 0))
    );
    if (clearing) addLocation({ id: "clearing-tool", label: "أداة تسكير السندات", detail: `مجموعة ${clearing.id}`, page: "clearing" });

    const isSaved = savedCashierIds.has(key) || savedBankIds.has(key);
    const isHeld = heldItems.some(item => `${item.fileSessionId}:${item.refId}` === key);
    const isManualGroupMember = manualGroups.some(group => [...group.banks, ...group.cashiers].some(item => transactionKey(item) === key));
    const isMahmoudWallet = !isResolvedBank && mahmoudWalletCashierIds.has(key);
    if (!isResolvedBank && visaItems.some(item => item.cashierId === row.id && (item.cashierSessionId ?? 0) === (row._fileSessionId ?? 0))) {
      addLocation({ id: "visa-items", label: "قائمة الفيزا", detail: (row as CashierRow).name, page: "main", tab: "visa" });
    }
    if (isMahmoudWallet) {
      addLocation({ id: "mahmoud-wallet", label: "المطابقة العامة · محفظة شخصية", detail: (row as CashierRow).name, page: "main", tab: "mahmoudWallet" });
    }
    const isNonReconciliationBank = isResolvedBank && nonReconciliationBankRows.some(item => transactionKey(item.bank) === key);
    if (isNonReconciliationBank) {
      addLocation({ id: "non-reconciliation-bank", label: "حركات بنك مستبعدة من المطابقة", detail: (row as BankRow).description, page: "main", tab: "nonReconciliationBank" });
    }
    const isAssignedToTool = !!split || !!cutoff || !!clearing;
    if (!isSaved && !isHeld && !isManualGroupMember && !isMahmoudWallet
      && !isAssignedToTool && !isNonReconciliationBank
      && !(!isResolvedBank && (jawwalPayCashierIds.has(key) || visaCashierIds.has(key)))) {
      addLocation({ id: "manual-workbench", label: "المطابقة اليدوية", detail: isResolvedBank ? (row as BankRow).description : (row as CashierRow).name, page: "manual" });
    }

    setReferenceLocator({ row, locations });
  };

  const handleOpenStageAInvoiceReference = (invoice: StageAResult) => {
    let index = stageAInvoices.findIndex(item => item === invoice);
    if (index < 0) index = stageAInvoices.findIndex(item => item.id === invoice.id && item.source === invoice.source);
    const row = index < 0 ? undefined : cashierSourceRows.find(item => item.id === 3000000 + index);
    if (!row) return false;
    handleOpenReference(row);
    return true;
  };

  useEffect(() => {
    const onPendingShortcut = (event: KeyboardEvent) => {
      const target = event.target;
      const focusedCheckbox = target instanceof HTMLElement
        ? target.closest<HTMLElement>("[data-pending-row-checkbox]")
        : null;
      if (page !== "main" || tab !== "pending" || !pagedPendingRows.length ||
        event.ctrlKey || event.metaKey || event.altKey || isTextEntryTarget(target) ||
        (target instanceof HTMLElement && target.closest("button, a, input, select, textarea, [contenteditable='true']") && !focusedCheckbox)) return;

      if (focusedCheckbox && ["ArrowDown", "ArrowUp", "Tab"].includes(event.code)) {
        const focusedKey = focusedCheckbox.dataset.pendingRowCheckbox;
        const focusedIndex = pagedPendingRows.findIndex((row: any) =>
          transactionPairKey(row.cashier, row.bank) === focusedKey
        );
        const direction = event.code === "ArrowUp" || (event.code === "Tab" && event.shiftKey) ? -1 : 1;
        const nextIndex = focusedIndex + direction;
        if (nextIndex >= 0 && nextIndex < pagedPendingRows.length) {
          event.preventDefault();
          setKeyboardPendingIndex(nextIndex);
          const nextRow = pagedPendingRows[nextIndex] as { cashier: CashierRow; bank: BankRow };
          pendingRowCheckboxRefs.current[transactionPairKey(nextRow.cashier, nextRow.bank)]?.focus();
        }
        return;
      }

      const currentIndex = Math.min(keyboardPendingIndex, pagedPendingRows.length - 1);
      const current = pagedPendingRows[currentIndex] as { cashier: CashierRow; bank: BankRow } | undefined;
      if (!current) return;
      const key = event.code;
      if (key === "ArrowDown" || key === "KeyN") {
        event.preventDefault();
        setKeyboardPendingIndex(Math.min(pagedPendingRows.length - 1, currentIndex + 1));
      } else if (key === "ArrowUp") {
        event.preventDefault();
        setKeyboardPendingIndex(Math.max(0, currentIndex - 1));
      } else if (key === "KeyB" || key === "Enter") {
        event.preventDefault();
        handleSaveMatch(current.cashier, current.bank);
        setKeyboardPendingIndex(Math.min(currentIndex, pagedPendingRows.length - 1));
      } else if (key === "KeyR") {
        event.preventDefault();
        handleOpenReference(current.cashier);
      } else if (key === "KeyT") {
        event.preventDefault();
        handleOpenReference(current.bank);
      } else if (key === "KeyF" || key === "KeyH") {
        event.preventDefault();
        openDrawer(current.cashier, current.bank, key === "KeyF" ? "cashier" : "bank");
      } else if (key === "KeyX") {
        event.preventDefault();
        handleRejectMatch(current.cashier, current.bank);
        setKeyboardPendingIndex(Math.min(currentIndex, pagedPendingRows.length - 1));
      } else if (key === "Space" && focusedCheckbox) {
        event.preventDefault();
        const pairKey = transactionPairKey(current.cashier, current.bank);
        setSelectedPendingKeys(previous => {
          const next = new Set(previous);
          if (next.has(pairKey)) next.delete(pairKey);
          else next.add(pairKey);
          return next;
        });
      }
    };
    window.addEventListener("keydown", onPendingShortcut);
    return () => window.removeEventListener("keydown", onPendingShortcut);
  }, [page, tab, pagedPendingRows, keyboardPendingIndex, handleSaveMatch, handleRejectMatch, handleOpenReference, openDrawer]);

  const followReferenceLocation = (location: ReferenceLocation) => {
    if (!referenceLocator) return;
    const reference = defaultReferenceForRow(referenceLocator.row);
    setReferenceFocus(reference);
    setReferenceLocator(null);
    setPage(location.page);
    if (location.tab) {
      setTab(location.tab);
      setResultSearch(reference);
      setPendingPage(1);
    }
  };

  useEffect(() => {
    if (!referenceFocus) return;
    const firstFrame = window.requestAnimationFrame(() => {
      window.requestAnimationFrame(() => {
        const target = Array.from(document.querySelectorAll<HTMLElement>("[data-default-reference]"))
          .find(element => (element.dataset.defaultReference ?? "").split("|").includes(referenceFocus));
        target?.scrollIntoView({ behavior: "smooth", block: "center" });
        target?.classList.add("ring-2", "ring-indigo-400");
        window.setTimeout(() => target?.classList.remove("ring-2", "ring-indigo-400"), 1800);
      });
    });
    return () => window.cancelAnimationFrame(firstFrame);
  }, [referenceFocus, page, tab, results]);

  const referenceInvoiceRow = referenceLocator && !("rawAmount" in referenceLocator.row)
    ? referenceLocator.row
    : null;
  const referenceLocatorDialog = referenceLocator && (
    <div className="fixed inset-0 z-[140] flex items-center justify-center bg-black/50 p-4" onMouseDown={event => {
      if (event.target === event.currentTarget) setReferenceLocator(null);
    }}>
      <section role="dialog" aria-modal="true" aria-labelledby="reference-locator-title" className="w-full max-w-3xl overflow-hidden rounded-2xl border bg-card shadow-2xl">
        <header className="flex items-start justify-between gap-4 border-b p-4">
          <div>
            <h2 id="reference-locator-title" className="font-bold">أماكن الحركة</h2>
            <p className="mt-1 font-mono text-sm font-bold text-indigo-700">{defaultReferenceForRow(referenceLocator.row)}</p>
            <p className="mt-1 text-xs text-muted-foreground">
              {"rawAmount" in referenceLocator.row ? referenceLocator.row.description : referenceLocator.row.name}
              {" · "}
              {fmtNum("rawAmount" in referenceLocator.row ? referenceLocator.row.rawAmount : referenceLocator.row.matchAmount)} ₪
            </p>
          </div>
          <button type="button" onClick={() => setReferenceLocator(null)} aria-label="إغلاق" className="rounded-lg p-2 hover:bg-muted"><X className="h-4 w-4" /></button>
        </header>
        {!("rawAmount" in referenceLocator.row) && <div className="grid grid-cols-2 gap-2 border-b bg-indigo-50/40 p-4 text-xs sm:grid-cols-3">
          {([
            ["رقم المستخدم", referenceLocator.row.userId || "—"],
            ["تاريخ الفاتورة", referenceLocator.row.date || "—"],
            ["وقت التسجيل", referenceLocator.row.registrationTime || "—"],
            ["رقم الفاتورة", referenceLocator.row.invoiceNumber || referenceLocator.row.ref || "—"],
            ["رقم الزبون", referenceLocator.row.customerNumber || "—"],
            ["منصة / طريقة الدفع", referenceLocator.row.category || referenceLocator.row.accountType || "—"],
            ["إجمالي الفاتورة", referenceLocator.row.invoiceTotal == null ? "—" : `${fmtNum(referenceLocator.row.invoiceTotal)} ₪`],
            ["نقدي", referenceLocator.row.cashAmount == null ? "—" : `${fmtNum(referenceLocator.row.cashAmount)} ₪`],
            ["شيك", referenceLocator.row.checkAmount == null ? "—" : `${fmtNum(referenceLocator.row.checkAmount)} ₪`],
            ["مدفوع بالكارد / التطبيق", referenceLocator.row.cardAmount == null ? "—" : `${fmtNum(referenceLocator.row.cardAmount)} ₪`],
            ["متبقي على الحساب", referenceLocator.row.debtAmount == null ? "—" : `${fmtNum(referenceLocator.row.debtAmount)} ₪`],
            ["الخصم", referenceLocator.row.discountAmount == null ? "—" : referenceLocator.row.discountAmount > 0 ? `${fmtNum(referenceLocator.row.discountAmount)} ₪` : "لا يوجد خصم"],
          ] as const).map(([label, value]) => <div key={label} className="rounded-lg border bg-white p-2"><span className="block text-[10px] text-muted-foreground">{label}</span><strong className="mt-1 block">{value}</strong></div>)}
        </div>}
        <div className="space-y-2 p-4">
          {referenceLocator.locations.map(location => (
            <button key={location.id} type="button" onClick={() => followReferenceLocation(location)}
              className="flex w-full items-center justify-between gap-3 rounded-xl border p-3 text-right hover:border-indigo-300 hover:bg-indigo-50">
              <span><b className="block text-sm">{location.label}</b><span className="mt-1 block text-xs text-muted-foreground">{location.detail}</span></span>
              <ArrowUpRight className="h-4 w-4 shrink-0 text-indigo-700" />
            </button>
          ))}
          {!referenceLocator.locations.length && <p className="rounded-lg bg-muted/50 p-3 text-sm text-muted-foreground">لم أجد لهذه الحركة موقعاً محفوظاً حالياً. جرّب المطابقة اليدوية أو أعد تشغيل المطابقة العامة.</p>}
          {page === "assist" && referenceInvoiceRow && (
            <button
              type="button"
              onClick={() => {
                if (!assistantSkipRef.current?.(referenceInvoiceRow)) {
                  setToast("هذه ليست الفاتورة الحالية في مساعد المطابقة.");
                  return;
                }
                setReferenceLocator(null);
              }}
              className="w-full rounded-xl border border-amber-300 bg-amber-50 px-4 py-3 text-sm font-semibold text-amber-900 hover:bg-amber-100"
            >
              تخطي هذه الفاتورة في مساعد المطابقة
            </button>
          )}
        </div>
      </section>
    </div>
  );

  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(null), 2800);
    return () => window.clearTimeout(timer);
  }, [toast]);

  useEffect(() => { setPendingPage(1); }, [resultSearch, resultType, minAmount, maxAmount]);

  const saveProject = async (name: string) => {
    const projects = await loadProjectsList();
    const stageASession = await storageGet<Record<string, unknown> | null>("stage_a_session", null);
    let stageASourceFiles: Awaited<ReturnType<typeof getStageASourceFiles>> = [];
    let sourceFilesReadFailed = false;
    if (centralStorageEnabled) {
      try {
        stageASourceFiles = await getStageASourceFiles();
      } catch (cause) {
        console.error("تعذر قراءة ملفات المرحلة A المؤقتة لحفظها مع المشروع:", cause);
        sourceFilesReadFailed = true;
      }
    }
    const uploadedSourceFiles: NonNullable<SavedProject["stageASourceFiles"]> = [];
    const uploadFailures: string[] = sourceFilesReadFailed ? ["تعذر قراءة الملفات المؤقتة"] : [];
    for (const group of stageASourceFiles) {
      for (const file of group.files) {
        const fileName = `${file.name}`;
        try {
          const id = await uploadStoredFile(file, fileName);
          uploadedSourceFiles.push({ id, stage: group.stage, source: group.source, fileName, mimeType: file.type || "application/octet-stream" });
        } catch (cause) {
          console.error(`تعذر حفظ ملف المشروع «${file.name}» في مكتبة الملفات:`, cause);
          uploadFailures.push(file.name);
        }
      }
    }
    const project: SavedProject = {
      id: `proj-${Date.now()}`, name,
      savedAt: new Date().toLocaleString("ar-SA"),
      stageASourceFiles: uploadedSourceFiles,
      bankHeaders, bankRowsRaw, bankMap, bankSwap, bankFileSessionId,
      cashHeaders, cashRowsRaw, cashMap, cashSwap, cashFileSessionId,
      manualGroups, savedMatches, reviewItems,
      bankSplits, cutoffBatches, clearingGroups,
      nameAliases,
      rejectedSpecialCashierIds: Array.from(rejectedSpecialCashierIds),
      mahmoudWalletCashierIds: Array.from(mahmoudWalletSelectedIds),
      rowNotes,
      amountTolerancePercent,
      bankFileSnapshots, cashFileSnapshots, walletBankFileSnapshots, walletCashFileSnapshots,
      jawwalBankFileSnapshots, jawwalCashFileSnapshots,
      stageAInvoices,
      stageADatabaseSummary,
      stageASession,
      walletBankHeaders, walletBankRows, walletBankMap, walletBankSwap, walletBankFileSessionId,
      walletCashHeaders, walletCashRows, walletCashMap, walletCashSwap, walletCashFileSessionId,
      jawwalBankHeaders, jawwalBankRows, jawwalBankMap, jawwalBankSwap, jawwalBankOptions, jawwalBankFileSessionId,
      jawwalCashHeaders, jawwalCashRows, jawwalCashMap, jawwalCashSwap, jawwalCashFileSessionId,
      platformNames, customPlatforms,
      rejectedPairs: Array.from(rejectedPairs),
      visaItems, heldItems, returnedHeldBank, returnedHeldCashier
    };
    await persistProjectsList([...projects, project]);
    setActiveProjectSourceFiles(uploadedSourceFiles);
    setToast(uploadFailures.length
      ? `تم حفظ المشروع، لكن تعذر رفع ${uploadFailures.length} ملف${uploadFailures.length === 1 ? "" : "ات"} إلى ملفاتي.`
      : `تم حفظ المشروع "${name}"${uploadedSourceFiles.length ? ` وحفظ ${uploadedSourceFiles.length} ملفاً في ملفاتي` : ""} ✓`);
  };

  const loadProject = (p: SavedProject) => {
    clearMatchUndo();
    setActiveProjectSourceFiles(p.stageASourceFiles ?? []);
    setBankH(p.bankHeaders); setBankRows(p.bankRowsRaw); setBankMap({ date: p.bankMap?.date ?? "", desc: p.bankMap?.desc ?? "", debit: p.bankMap?.debit ?? "", credit: p.bankMap?.credit ?? "", accountType: p.bankMap?.accountType ?? "", ref: (p.bankMap as any)?.ref ?? "" }); setBankSwap(p.bankSwap ?? false);
    setBankFileSessionId(p.bankFileSessionId ?? 0);
    setCashH(p.cashHeaders); setCashRows(p.cashRowsRaw); setCashMap({ date: p.cashMap?.date ?? "", name: p.cashMap?.name ?? "", debit: p.cashMap?.debit ?? "", credit: p.cashMap?.credit ?? "", accountType: p.cashMap?.accountType ?? "", ref: p.cashMap?.ref ?? "", userId: p.cashMap?.userId ?? "" }); setCashSwap(p.cashSwap ?? false);
    setCashFileSessionId(p.cashFileSessionId ?? 0);
    setStageAInvoices(p.stageAInvoices ?? []);
    setStageADatabaseSummary(p.stageADatabaseSummary ?? { databaseInvoiceCount: 0, databaseInvoiceTotal: 0, databaseCashTotal: 0, databaseDebtTotal: 0, databaseCardTotal: 0 });
    void storageSet("stage_a_session", p.stageASession ?? null);
    setWalletBankHeaders(p.walletBankHeaders ?? []); setWalletBankRows(p.walletBankRows ?? []);
    setWalletBankMap({ date: p.walletBankMap?.date ?? "", desc: p.walletBankMap?.desc ?? "", debit: p.walletBankMap?.debit ?? "", credit: p.walletBankMap?.credit ?? "", ref: p.walletBankMap?.ref ?? "" });
    setWalletBankSwap(p.walletBankSwap ?? false); setWalletBankFileSessionId(p.walletBankFileSessionId ?? 0);
    setWalletCashHeaders(p.walletCashHeaders ?? []); setWalletCashRows(p.walletCashRows ?? []);
    setWalletCashMap({ date: p.walletCashMap?.date ?? "", name: p.walletCashMap?.name ?? "", debit: p.walletCashMap?.debit ?? "", credit: p.walletCashMap?.credit ?? "", ref: p.walletCashMap?.ref ?? "", userId: p.walletCashMap?.userId ?? "" });
    setWalletCashSwap(p.walletCashSwap ?? false); setWalletCashFileSessionId(p.walletCashFileSessionId ?? 0);
    setJawwalBankHeaders(p.jawwalBankHeaders ?? []); setJawwalBankRows(p.jawwalBankRows ?? []);
    setJawwalBankMap({ date: p.jawwalBankMap?.date ?? "", desc: p.jawwalBankMap?.desc ?? "", debit: p.jawwalBankMap?.debit ?? "", credit: p.jawwalBankMap?.credit ?? "", ref: p.jawwalBankMap?.ref ?? "", totalAmount: p.jawwalBankMap?.totalAmount ?? "", movementAmount: p.jawwalBankMap?.movementAmount ?? "", direction: p.jawwalBankMap?.direction ?? "", movementType: p.jawwalBankMap?.movementType ?? "" });
    setJawwalBankSwap(p.jawwalBankSwap ?? true); setJawwalBankOptions({ useTotalAmount: true, readDebitCredit: true, ...(p.jawwalBankOptions ?? {}) });
    setJawwalBankFileSessionId(p.jawwalBankFileSessionId ?? 0);
    setJawwalCashHeaders(p.jawwalCashHeaders ?? []); setJawwalCashRows(p.jawwalCashRows ?? []);
    setJawwalCashMap({ date: p.jawwalCashMap?.date ?? "", name: p.jawwalCashMap?.name ?? "", debit: p.jawwalCashMap?.debit ?? "", credit: p.jawwalCashMap?.credit ?? "", ref: p.jawwalCashMap?.ref ?? "", userId: p.jawwalCashMap?.userId ?? "" });
    setJawwalCashSwap(p.jawwalCashSwap ?? false); setJawwalCashFileSessionId(p.jawwalCashFileSessionId ?? 0);
    setManualGroups(p.manualGroups); setSavedMatches(dedupeSavedMatches(p.savedMatches)); setReviewItems(p.reviewItems ?? []);
    setBankSplits(p.bankSplits ?? []);
    setCutoffBatches(p.cutoffBatches ?? []);
    setClearingGroups(p.clearingGroups ?? []);
    setNameAliases(p.nameAliases ?? {});
    setRejectedSpecialCashierIds(normalizeSpecialKeySet(p.rejectedSpecialCashierIds));
    setMahmoudWalletSelectedIds(new Set(p.mahmoudWalletCashierIds ?? []));
    setRowNotes(p.rowNotes ?? {});
    setAmountTolerancePercent(Math.min(10, Math.max(0, p.amountTolerancePercent ?? 0.5)));
    setBankFileSnapshots(p.bankFileSnapshots ?? []);
    setCashFileSnapshots(p.cashFileSnapshots ?? []);
    setWalletBankFileSnapshots(p.walletBankFileSnapshots ?? []);
    setWalletCashFileSnapshots(p.walletCashFileSnapshots ?? []);
    setJawwalBankFileSnapshots(p.jawwalBankFileSnapshots ?? []);
    setJawwalCashFileSnapshots(p.jawwalCashFileSnapshots ?? []);
    setBankFileCount(p.bankFileSnapshots?.[0]?.fileCount ?? 0);
    setCashFileCount(p.cashFileSnapshots?.[0]?.fileCount ?? 0);
    setWalletBankFileCount(p.walletBankFileSnapshots?.[0]?.fileCount ?? 0);
    setWalletCashFileCount(p.walletCashFileSnapshots?.[0]?.fileCount ?? 0);
    setJawwalBankFileCount(p.jawwalBankFileSnapshots?.[0]?.fileCount ?? 0);
    setJawwalCashFileCount(p.jawwalCashFileSnapshots?.[0]?.fileCount ?? 0);
    setRejected(new Set(p.rejectedPairs));
    setVisaItems(p.visaItems ?? []);
    setHeldItems(dedupeHeldItems(p.heldItems ?? []));
    setReturnedHeldBank((p as any).returnedHeldBank ?? []);
    setReturnedHeldCashier((p as any).returnedHeldCashier ?? []);
    setPlatformNames({ bank: "بنك فلسطين", wallet: "بال بي", jawwal: "جوال بي", ...(p.platformNames ?? {}) });
    setCustomPlatforms(p.customPlatforms ?? []);
    nextCustomPlatformId.current = Math.max(0, ...(p.customPlatforms ?? []).map(platform => platform.id + 1));
    nextCustomSourceSessionId.current = Math.max(1, ...(p.customPlatforms ?? []).flatMap(platform => [platform.transfer?.fileSessionId ?? 0, platform.invoice?.fileSessionId ?? 0]).map(id => id + 1));
    setWalletBankFileSessionId(p.walletBankFileSessionId ?? 0);
    setWalletCashFileSessionId(p.walletCashFileSessionId ?? 0);
    setJawwalBankFileSessionId(p.jawwalBankFileSessionId ?? 0);
    setJawwalCashFileSessionId(p.jawwalCashFileSessionId ?? 0);
    setJawwalBankOptions({ useTotalAmount: true, readDebitCredit: true, ...(p.jawwalBankOptions ?? {}) });
    setBankFile(null); setCashFile(null);
    setResults(null);
    setPage("main");
    setTab("saved");
  };

  const deleteProject = async (id: string) => {
    const projects = (await loadProjectsList()).filter(x => x.id !== id);
    await persistProjectsList(projects);
  };

  const handleClearSavedMatches = () => {
    if (!savedMatches.length) return;
    if (!window.confirm(`مسح جميع المطابقات المؤكدة (${savedMatches.length})؟`)) return;
    clearMatchUndo();
    setSavedMatches([]);
    setManualGroups([]);
    setSelectedSavedIds(new Set());
    setExpandedMatchKey(null);
    setShowDeleteMenu(false);
    setToast("تم إلغاء جميع المطابقات المؤكدة، بما فيها المطابقات اليدوية");
  };

  const handleClearSession = (preserveHeldItems = false) => {
    const heldCount = heldItems.length;
    const message = preserveHeldItems
      ? `مسح ملفات وبيانات الجلسة الحالية مع الإبقاء على المعلّقات (${heldCount})؟\nسيتم حذف الكشوف والملفات المحفوظة للاسترجاع والمطابقات والفيزا والرفض، وستبقى الحركات المعلّقة لمتابعتها مع كشف اليوم التالي.`
      : "مسح الجلسة الحالية كاملة؟\nهاد بيحذف: الملفات (بكل المنصات)، الملفات المحفوظة للاسترجاع، المطابقات، الفيزا، المعلقات، والرفوض.\nالمشاريع المحفوظة رح تضل محفوظة.";
    if (!window.confirm(message)) return;
    clearMatchUndo();
    setBankH([]); setBankRows([]); setBankMap({date:"",desc:"",debit:"",credit:"",accountType:"",ref:""}); setBankSwap(false); setBankFile(null); setBankFileCount(0); setBankFileSessionId(0);
    setCashH([]); setCashRows([]); setCashMap({date:"",name:"",debit:"",credit:"",accountType:"",ref:"",userId:""}); setCashSwap(false); setCashFile(null); setCashFileCount(0); setCashFileSessionId(0); setStageAInvoices([]);
    setStageADatabaseSummary({ databaseInvoiceCount: 0, databaseInvoiceTotal: 0, databaseCashTotal: 0, databaseDebtTotal: 0, databaseCardTotal: 0 });
    for (const key of ["stage_a_session", "stage_b_session", "stage_b_aliases", "stage_c_platform_session", "stage_c_invoices"]) {
      void storageDelete(key).catch(error => console.error(`تعذر مسح بيانات الجلسة «${key}»:`, error));
    }
    void clearStageASourceFiles().catch(error => console.error("تعذر حذف ملفات الجلسة المؤقتة للمرحلة A:", error));
    setWalletBankFile(null); setWalletBankFileCount(0); setWalletBankHeaders([]); setWalletBankRows([]); setWalletBankMap({date:"",desc:"",debit:"",credit:"",ref:""}); setWalletBankSwap(false);
    setWalletCashFile(null); setWalletCashFileCount(0); setWalletCashHeaders([]); setWalletCashRows([]); setWalletCashMap({date:"",name:"",debit:"",credit:"",ref:"",userId:""}); setWalletCashSwap(false);
    setWalletBankFileSessionId(id => id + 1); setWalletCashFileSessionId(id => id + 1); setJawwalBankFileSessionId(id => id + 1); setJawwalCashFileSessionId(id => id + 1);
    setJawwalBankFile(null); setJawwalBankFileCount(0); setJawwalBankHeaders([]); setJawwalBankRows([]); setJawwalBankMap({date:"",desc:"",debit:"",credit:"",ref:"",totalAmount:"",movementAmount:"",direction:"",movementType:""}); setJawwalBankSwap(true); setJawwalBankOptions({useTotalAmount:true,readDebitCredit:true});
    setJawwalCashFile(null); setJawwalCashFileCount(0); setJawwalCashHeaders([]); setJawwalCashRows([]); setJawwalCashMap({date:"",name:"",debit:"",credit:"",ref:"",userId:""}); setJawwalCashSwap(false);
    setPlatformNames({ bank: "بنك فلسطين", wallet: "بال بي", jawwal: "جوال بي" }); setCustomPlatforms([]); nextCustomPlatformId.current = 0;
    setBankFileSnapshots([]); setCashFileSnapshots([]);
    setWalletBankFileSnapshots([]); setWalletCashFileSnapshots([]);
    setJawwalBankFileSnapshots([]); setJawwalCashFileSnapshots([]);
    setManualGroups([]); setSavedMatches([]); setRejected(new Set());
    setReviewItems([]);
    setBankSplits([]); setCutoffBatches([]); setClearingGroups([]);
    setVisaItems([]);
    if (!preserveHeldItems) setHeldItems([]);
    setReturnedHeldBank([]); setReturnedHeldCashier([]);
    setRejectedSpecialCashierIds(new Set());
    setResults(null); setTab("saved");
    setSessionRestoredNotice(false);
    setShowDeleteMenu(false);
    setToast(preserveHeldItems
      ? `تم مسح الجلسة والإبقاء على ${heldCount} حركة معلّقة. ارفع كشف اليوم التالي لمتابعتها.`
      : "تم مسح الجلسة الحالية كاملة.");
  };

  const handleClearPlatform = (plat: "bank" | "cash" | "walletBank" | "walletCash" | "jawwalBank" | "jawwalCash") => {
    if (!window.confirm("مسح ملف هذه المنصة والملفات المحفوظة للاسترجاع الخاصة بها فقط؟\nباقي المنصات والمطابقات المؤكدة رح تضل كما هي.")) return;
    if (plat === "bank") { setBankFile(null); setBankFileCount(0); setBankH([]); setBankRows([]); setBankMap({date:"",desc:"",debit:"",credit:"",accountType:"",ref:""}); setBankSwap(false); setBankFileSessionId(id => id + 1); setBankFileSnapshots([]); }
    if (plat === "cash") { setCashFile(null); setCashFileCount(0); setCashH([]); setCashRows([]); setCashMap({date:"",name:"",debit:"",credit:"",accountType:"",ref:"",userId:""}); setCashSwap(false); setCashFileSessionId(id => id + 1); setCashFileSnapshots([]); }
    if (plat === "walletBank") { setWalletBankFile(null); setWalletBankFileCount(0); setWalletBankHeaders([]); setWalletBankRows([]); setWalletBankMap({date:"",desc:"",debit:"",credit:"",ref:""}); setWalletBankSwap(false); setWalletBankFileSessionId(id => id + 1); setWalletBankFileSnapshots([]); }
    if (plat === "walletCash") { setWalletCashFile(null); setWalletCashFileCount(0); setWalletCashHeaders([]); setWalletCashRows([]); setWalletCashMap({date:"",name:"",debit:"",credit:"",ref:"",userId:""}); setWalletCashSwap(false); setWalletCashFileSessionId(id => id + 1); setWalletCashFileSnapshots([]); }
    if (plat === "jawwalBank") { setJawwalBankFile(null); setJawwalBankFileCount(0); setJawwalBankHeaders([]); setJawwalBankRows([]); setJawwalBankMap({date:"",desc:"",debit:"",credit:"",ref:"",totalAmount:"",movementAmount:"",direction:"",movementType:""}); setJawwalBankSwap(true); setJawwalBankOptions({useTotalAmount:true,readDebitCredit:true}); setJawwalBankFileSessionId(id => id + 1); setJawwalBankFileSnapshots([]); }
    if (plat === "jawwalCash") { setJawwalCashFile(null); setJawwalCashFileCount(0); setJawwalCashHeaders([]); setJawwalCashRows([]); setJawwalCashMap({date:"",name:"",debit:"",credit:"",ref:"",userId:""}); setJawwalCashSwap(false); setJawwalCashFileSessionId(id => id + 1); setJawwalCashFileSnapshots([]); }
    setResults(null);
    setToast("تم مسح بيانات هذه المنصة ✓");
  };

  const handleDeleteAllProjects = async () => {
    try {
      const projects = await loadProjectsList();
      if (!projects.length) return;
      if (!window.confirm(`حذف كل المشاريع المحفوظة (${projects.length})؟\nما رح تقدر ترجعها!`)) return;
      await persistProjectsList([]);
      setShowDeleteMenu(false);
      setToast("تم حذف كل المشاريع المحفوظة.");
    } catch (cause) {
      console.error("تعذر حذف المشاريع المحفوظة:", cause);
      setToast(cause instanceof Error ? cause.message : "تعذر حذف المشاريع المحفوظة.");
    }
  };

  const buildResumeData = (): ResumeData => ({
    bankHeaders, bankRowsRaw, bankMap, bankSwap, bankFileSessionId,
    cashHeaders, cashRowsRaw, cashMap, cashSwap, cashFileSessionId,
    manualGroups, savedMatches, reviewItems,
    bankSplits, cutoffBatches, clearingGroups,
    nameAliases,
    rejectedSpecialCashierIds: Array.from(rejectedSpecialCashierIds),
    mahmoudWalletCashierIds: Array.from(mahmoudWalletSelectedIds),
    rowNotes,
    amountTolerancePercent,
    bankFileSnapshots, cashFileSnapshots, walletBankFileSnapshots, walletCashFileSnapshots,
    jawwalBankFileSnapshots, jawwalCashFileSnapshots,
    stageAInvoices,
    walletBankHeaders, walletBankRows, walletBankMap, walletBankSwap, walletBankFileSessionId,
    walletCashHeaders, walletCashRows, walletCashMap, walletCashSwap, walletCashFileSessionId,
    jawwalBankHeaders, jawwalBankRows, jawwalBankMap, jawwalBankSwap, jawwalBankOptions, jawwalBankFileSessionId,
    jawwalCashHeaders, jawwalCashRows, jawwalCashMap, jawwalCashSwap, jawwalCashFileSessionId,
    platformNames, customPlatforms,
    rejectedPairs: Array.from(rejectedPairs),
    visaItems, heldItems, returnedHeldBank, returnedHeldCashier
  });

  const loadResumeData = (p: ResumeData) => {
    setBankH(p.bankHeaders || []); setBankRows(p.bankRowsRaw || []);
    setBankMap({ date: p.bankMap?.date ?? "", desc: p.bankMap?.desc ?? "", debit: p.bankMap?.debit ?? "", credit: p.bankMap?.credit ?? "", accountType: p.bankMap?.accountType ?? "", ref: (p.bankMap as any)?.ref ?? "" }); setBankSwap(!!p.bankSwap);
    setBankFileSessionId(p.bankFileSessionId ?? 0);
    setCashH(p.cashHeaders || []); setCashRows(p.cashRowsRaw || []);
    setCashMap({ date: p.cashMap?.date ?? "", name: p.cashMap?.name ?? "", debit: p.cashMap?.debit ?? "", credit: p.cashMap?.credit ?? "", accountType: p.cashMap?.accountType ?? "", ref: p.cashMap?.ref ?? "", userId: p.cashMap?.userId ?? "" }); setCashSwap(!!p.cashSwap);
    setCashFileSessionId(p.cashFileSessionId ?? 0);
    setStageAInvoices(p.stageAInvoices ?? []);
    setStageADatabaseSummary(p.stageADatabaseSummary ?? { databaseInvoiceCount: 0, databaseInvoiceTotal: 0, databaseCashTotal: 0, databaseDebtTotal: 0, databaseCardTotal: 0 });
    setWalletBankHeaders(p.walletBankHeaders ?? []); setWalletBankRows(p.walletBankRows ?? []);
    setWalletBankMap({ date: p.walletBankMap?.date ?? "", desc: p.walletBankMap?.desc ?? "", debit: p.walletBankMap?.debit ?? "", credit: p.walletBankMap?.credit ?? "", ref: p.walletBankMap?.ref ?? "" });
    setWalletBankSwap(p.walletBankSwap ?? false); setWalletBankFileSessionId(p.walletBankFileSessionId ?? 0);
    setWalletCashHeaders(p.walletCashHeaders ?? []); setWalletCashRows(p.walletCashRows ?? []);
    setWalletCashMap({ date: p.walletCashMap?.date ?? "", name: p.walletCashMap?.name ?? "", debit: p.walletCashMap?.debit ?? "", credit: p.walletCashMap?.credit ?? "", ref: p.walletCashMap?.ref ?? "", userId: p.walletCashMap?.userId ?? "" });
    setWalletCashSwap(p.walletCashSwap ?? false); setWalletCashFileSessionId(p.walletCashFileSessionId ?? 0);
    setJawwalBankHeaders(p.jawwalBankHeaders ?? []); setJawwalBankRows(p.jawwalBankRows ?? []);
    setJawwalBankMap({ date: p.jawwalBankMap?.date ?? "", desc: p.jawwalBankMap?.desc ?? "", debit: p.jawwalBankMap?.debit ?? "", credit: p.jawwalBankMap?.credit ?? "", ref: p.jawwalBankMap?.ref ?? "", totalAmount: p.jawwalBankMap?.totalAmount ?? "", movementAmount: p.jawwalBankMap?.movementAmount ?? "", direction: p.jawwalBankMap?.direction ?? "", movementType: p.jawwalBankMap?.movementType ?? "" });
    setJawwalBankSwap(p.jawwalBankSwap ?? true); setJawwalBankOptions({ useTotalAmount: true, readDebitCredit: true, ...(p.jawwalBankOptions ?? {}) });
    setJawwalBankFileSessionId(p.jawwalBankFileSessionId ?? 0);
    setJawwalCashHeaders(p.jawwalCashHeaders ?? []); setJawwalCashRows(p.jawwalCashRows ?? []);
    setJawwalCashMap({ date: p.jawwalCashMap?.date ?? "", name: p.jawwalCashMap?.name ?? "", debit: p.jawwalCashMap?.debit ?? "", credit: p.jawwalCashMap?.credit ?? "", ref: p.jawwalCashMap?.ref ?? "", userId: p.jawwalCashMap?.userId ?? "" });
    setJawwalCashSwap(p.jawwalCashSwap ?? false); setJawwalCashFileSessionId(p.jawwalCashFileSessionId ?? 0);
    setPlatformNames({ bank: "بنك فلسطين", wallet: "بال بي", jawwal: "جوال بي", ...(p.platformNames ?? {}) });
    setCustomPlatforms(p.customPlatforms ?? []);
    nextCustomPlatformId.current = Math.max(0, ...(p.customPlatforms ?? []).map(platform => platform.id + 1));
    nextCustomSourceSessionId.current = Math.max(1, ...(p.customPlatforms ?? []).flatMap(platform => [platform.transfer?.fileSessionId ?? 0, platform.invoice?.fileSessionId ?? 0]).map(id => id + 1));
    setManualGroups(p.manualGroups || []); setSavedMatches(dedupeSavedMatches(p.savedMatches || [])); setReviewItems(p.reviewItems ?? []);
    setBankSplits(p.bankSplits || []);
    setCutoffBatches(p.cutoffBatches || []);
    setClearingGroups(p.clearingGroups || []);
    setNameAliases(p.nameAliases ?? {});
    setRejectedSpecialCashierIds(normalizeSpecialKeySet(p.rejectedSpecialCashierIds));
    setMahmoudWalletSelectedIds(new Set(p.mahmoudWalletCashierIds ?? []));
    setRowNotes(p.rowNotes ?? {});
    setAmountTolerancePercent(Math.min(10, Math.max(0, p.amountTolerancePercent ?? 0.5)));
    setBankFileSnapshots(p.bankFileSnapshots ?? []);
    setCashFileSnapshots(p.cashFileSnapshots ?? []);
    setWalletBankFileSnapshots(p.walletBankFileSnapshots ?? []);
    setWalletCashFileSnapshots(p.walletCashFileSnapshots ?? []);
    setJawwalBankFileSnapshots(p.jawwalBankFileSnapshots ?? []);
    setJawwalCashFileSnapshots(p.jawwalCashFileSnapshots ?? []);
    setBankFileCount(p.bankFileSnapshots?.[0]?.fileCount ?? 0);
    setCashFileCount(p.cashFileSnapshots?.[0]?.fileCount ?? 0);
    setWalletBankFileCount(p.walletBankFileSnapshots?.[0]?.fileCount ?? 0);
    setWalletCashFileCount(p.walletCashFileSnapshots?.[0]?.fileCount ?? 0);
    setJawwalBankFileCount(p.jawwalBankFileSnapshots?.[0]?.fileCount ?? 0);
    setJawwalCashFileCount(p.jawwalCashFileSnapshots?.[0]?.fileCount ?? 0);
    setRejected(new Set(p.rejectedPairs || []));
    setVisaItems(p.visaItems || []);
    setHeldItems(dedupeHeldItems(p.heldItems || []));
    setReturnedHeldBank((p.returnedHeldBank as BankRow[]) ?? []);
    setReturnedHeldCashier((p.returnedHeldCashier as CashierRow[]) ?? []);
    setBankFile(null); setCashFile(null);
    setWalletBankFile(null); setWalletCashFile(null);
    setJawwalBankFile(null); setJawwalCashFile(null);
    setShowBankSnapshots(false); setShowCashSnapshots(false);
    setShowWalletBankSnapshots(false); setShowWalletCashSnapshots(false);
    setShowJawwalBankSnapshots(false); setShowJawwalCashSnapshots(false);
    setResults(null); setTab("saved");
  };

  const handleImportResumeFile = async (f: File) => {
    setError(null);
    try {
      const data = await extractResumeData(f);
      loadResumeData(data);
      const rowCount = (rows?: unknown[]) => Array.isArray(rows) ? rows.length : 0;
      const customTransferCount = (data.customPlatforms ?? []).reduce((sum, platform) => sum + rowCount(platform.transfer?.rows), 0);
      const customInvoiceCount = (data.customPlatforms ?? []).reduce((sum, platform) => sum + rowCount(platform.invoice?.rows), 0);
      const bankCount = rowCount(data.bankRowsRaw) + rowCount(data.walletBankRows) + rowCount(data.jawwalBankRows) + customTransferCount;
      const invoiceCount = rowCount(data.cashRowsRaw) + rowCount(data.walletCashRows) + rowCount(data.jawwalCashRows) + customInvoiceCount;
      setToast(`تم استرجاع الملف: ${invoiceCount} صف فاتورة و${bankCount} صف حوالة من البيانات الخام. الصفوف ذات المبلغ غير الصالح أو الصفري لا تدخل المطابقة، لكنها محفوظة ضمن بيانات المصدر.`);
    } catch (error) {
      console.error("تعذّر استيراد ملف الاستكمال:", error);
      setError(error instanceof Error ? error.message : "تعذّر استيراد الملف. تحقق من سلامته ثم حاول مجدداً.");
      return;
    }
  };

  if (page === "splitTool") {
    const unavailableBankKeys = new Set(matchableBank
      .filter(row => savedBankIds.has(transactionKey(row)))
      .map(transactionKey));
    const unavailableCashierKeys = new Set(activeCashier
      .filter(row => savedCashierIds.has(transactionKey(row)))
      .map(transactionKey));
    heldItems.filter(item => item.kind === "cashier").forEach(item => unavailableCashierKeys.add(`${item.fileSessionId}:${item.refId}`));
    cutoffBatches.forEach(batch => {
      unavailableCashierKeys.add(`${batch.aggregateSessionId}:${batch.aggregateId}`);
      batch.sourceRows.forEach(row => unavailableCashierKeys.add(`${row.sessionId}:${row.id}`));
    });
    bankSplits.forEach(split => split.allocations.forEach(allocation => {
      if (allocation.kind === "company" && allocation.cashierId != null) {
        unavailableCashierKeys.add(`${allocation.cashierSessionId ?? 0}:${allocation.cashierId}`);
      }
    }));
    return <>{referenceLocatorDialog}<SplitTool
      bankRows={matchableBank}
      cashierRows={activeCashier}
      splits={bankSplits}
      unavailableBankKeys={unavailableBankKeys}
      unavailableCashierKeys={unavailableCashierKeys}
      onSave={handleSaveBankSplit}
      onDelete={handleDeleteBankSplit}
      onReview={handleToolReview}
      onOpenReference={handleOpenReference}
      onBack={() => setPage("main")}
    /></>;
  }

  if (page === "batchCutoff") {
    const unavailableCashierKeys = new Set(activeCashier
      .filter(row => savedCashierIds.has(transactionKey(row)))
      .map(transactionKey));
    heldItems.filter(item => item.kind === "cashier").forEach(item => unavailableCashierKeys.add(`${item.fileSessionId}:${item.refId}`));
    cutoffBatches.forEach(batch => {
      unavailableCashierKeys.add(`${batch.aggregateSessionId}:${batch.aggregateId}`);
      batch.sourceRows.forEach(row => unavailableCashierKeys.add(`${row.sessionId}:${row.id}`));
    });
    bankSplits.forEach(split => split.allocations.forEach(allocation => {
      if (allocation.kind === "company" && allocation.cashierId != null) {
        unavailableCashierKeys.add(`${allocation.cashierSessionId ?? 0}:${allocation.cashierId}`);
      }
    }));
    return <>{referenceLocatorDialog}<BatchCutoffTool
      cashierRows={activeCashier}
      batches={cutoffBatches}
      unavailableCashierKeys={unavailableCashierKeys}
      onSave={handleSaveCutoffBatch}
      onDelete={handleDeleteCutoffBatch}
      onReview={handleToolCashierReview}
      onOpenReference={handleOpenReference}
      onBack={() => setPage("main")}
    /></>;
  }

  if (page === "clearing") {
    const unavailableCashierKeys = new Set<string>([
      ...activeCashier.filter(row => savedCashierIds.has(transactionKey(row))).map(transactionKey),
      ...manualGroups.flatMap(group => group.cashiers.map(transactionKey)),
      ...heldItems.filter(item => item.kind === "cashier").map(item => `${item.fileSessionId}:${item.refId}`),
      ...cutoffBatches.flatMap(batch => [
        `${batch.aggregateSessionId}:${batch.aggregateId}`,
        ...batch.sourceRows.map(reference => `${reference.sessionId}:${reference.id}`),
      ]),
      ...bankSplits.flatMap(split => split.allocations
        .filter(allocation => allocation.kind === "company" && allocation.cashierId != null)
        .map(allocation => `${allocation.cashierSessionId ?? 0}:${allocation.cashierId}`)),
      ...clearedCashierKeys,
    ]);
    return <>{referenceLocatorDialog}<ClearingTool
      cashierRows={activeCashier}
      groups={clearingGroups}
      unavailableCashierKeys={unavailableCashierKeys}
      onSave={handleSaveClearingGroup}
      onDelete={handleDeleteClearingGroup}
      onReview={handleToolCashierReview}
      onOpenReference={handleOpenReference}
      onBack={() => setPage("main")}
    /></>;
  }

  if (page === "manual") {
    const manualCashierRows = reconciliationCashierRows.filter(c => !visaCashierIds.has(transactionKey(c)) && !specialCashierIds.has(transactionKey(c)));
    const manualBankRows = reconciliationBankRows;
    return (
      <>{referenceLocatorDialog}<ManualWorkbench
        bankRows={manualBankRows}
        cashierRows={manualCashierRows}
        manualGroups={manualGroups}
        onAddGroup={handleAddGroup}
        onRemoveGroup={handleRemoveGroup}
        onBack={() => setPage("main")}
        savedCashierIds={savedCashierIds}
        savedBankIds={savedBankIds}
        onHoldBank={handleHoldBank}
        onHoldCashier={handleHoldCashier}
        onToggleReview={(cashier, bank, detail) => handleToggleReviewSnapshot(cashier, bank, detail)}
        rowNotes={rowNotes}
        onSaveNote={handleSaveRowNote}
        isInReview={(cashier, bank) => reviewItems.some(item =>
          item.id === `review-pair:${transactionPairKey(cashier, bank)}`
        )}
        onOpenReference={handleOpenReference}
        focusReference={referenceFocus}
      /></>
    );
  }

  if (page === "report") {
    return <ReportPage
      bank={reportBankRows}
      cashier={reportCashierRows}
      results={results}
      stageAInvoices={stageAInvoices}
      stageADatabaseSummary={stageADatabaseSummary}
      bankSplits={bankSplits}
      cutoffBatches={cutoffBatches}
      clearingGroups={clearingGroups}
      sourceCashierRows={activeCashier}
      onBack={() => setPage("main")}
    />;
  }

  if (page === "assist") {
    return (
      <>{referenceLocatorDialog}<MatchAssistant
        cashierRows={uCashierRows.map((r:any) => r.cashier as CashierRow)}
        bankRows={uBankRows.map((r:any) => r.bank as BankRow)}
        rejectedPairs={rejectedPairs}
        onMatch={(c, b) => handleSaveMatch(c, b)}
        onHold={(c) => handleHoldCashier(c)}
        onHoldBank={(b) => handleHoldBank(b)}
        onReject={(c, b) => handleRejectMatch(c, b)}
        onBack={() => setPage("main")}
        onOpenReference={handleOpenReference}
        assistantSkipRef={assistantSkipRef}
      /></>
    );
  }

  if (page === "stageBChoice") {
    return (
      <div dir="rtl" className="min-h-screen bg-slate-50 p-6 text-slate-900">
        <div className="mx-auto max-w-3xl space-y-6">
          <button type="button" onClick={() => setPage("recon2")} className="flex items-center gap-1 text-sm text-slate-600 hover:text-slate-900">
            <ChevronRight className="h-4 w-4" />العودة إلى منصة المراحل
          </button>
          <header className="rounded-2xl border border-slate-200 bg-white p-6 text-center shadow-sm">
            <h1 className="text-2xl font-bold">منصة المطابقة الذكية للفواتير والحوالات</h1>
            <p className="mt-2 text-sm text-slate-600">اختر طريقة البدء: جهّز ملفاتك أولاً أو انتقل مباشرة إلى منصة التسوية.</p>
          </header>
          <section className="grid gap-4 sm:grid-cols-2">
            <button type="button" onClick={() => setPage("upload")} className="rounded-2xl border border-blue-200 bg-white p-6 text-right shadow-sm transition hover:-translate-y-0.5 hover:shadow-md">
              <Upload className="mb-4 h-8 w-8 text-blue-600" />
              <span className="block text-lg font-bold text-blue-700">رفع الملفات</span>
              <span className="mt-2 block text-sm leading-6 text-slate-600">ارفع كشوفات البنك والفواتير وانتظر اكتمال تجهيزها قبل المطابقة.</span>
            </button>
            <button type="button" onClick={() => setPage("main")} className="rounded-2xl border border-emerald-200 bg-white p-6 text-right shadow-sm transition hover:-translate-y-0.5 hover:shadow-md">
              <Sparkles className="mb-4 h-8 w-8 text-emerald-600" />
              <span className="block text-lg font-bold text-emerald-700">منصة التسوية</span>
              <span className="mt-2 block text-sm leading-6 text-slate-600">افتح شاشة المطابقة والعمل على البيانات المتاحة في جلستك.</span>
            </button>
          </section>
        </div>
      </div>
    );
  }

  if (page === "recon2") {
    return <Suspense fallback={<PlatformLoading />}><Reconciliation2 projectSourceFiles={activeProjectSourceFiles} onSafeActivities={onSafeActivities} onOpenInvoiceReference={handleOpenStageAInvoiceReference} onBack={() => setPage("main")} onStageBToMain={() => setPage("stageBChoice")} onStageAToMain={(results, summary) => {
      const transferred = results;
      setStageAInvoices(previous => {
        const transferredKeys = new Set(transferred.map(result => `${result.source}:${result.id}`));
        return [...previous.filter(result => !transferredKeys.has(`${result.source}:${result.id}`)), ...transferred];
      });
      if (summary) setStageADatabaseSummary(summary);
      setPage("main");
    }} /></Suspense>;
  }
  if (page === "recon2StageA") {
    return <Suspense fallback={<PlatformLoading />}><Reconciliation2 projectSourceFiles={activeProjectSourceFiles} initialStage="stageA" onSafeActivities={onSafeActivities} onOpenInvoiceReference={handleOpenStageAInvoiceReference} onBack={() => setPage("main")} onStageBToMain={() => setPage("stageBChoice")} onStageAToMain={(results, summary) => {
      const transferred = results;
      setStageAInvoices(previous => {
        const transferredKeys = new Set(transferred.map(result => `${result.source}:${result.id}`));
        return [...previous.filter(result => !transferredKeys.has(`${result.source}:${result.id}`)), ...transferred];
      });
      if (summary) setStageADatabaseSummary(summary);
      setPage("main");
    }} /></Suspense>;
  }
  if (page === "recon2Visa") {
    const visaInvoices: StageAResult[] = visaItems.map(item => ({
      id: item.cashierId,
      source: "المنصة الرئيسية",
      name: item.name,
      originalName: item.rawName,
      amount: item.amount,
      databaseName: "",
      databaseAmount: 0,
      invoiceTotal: 0,
      cashAmount: 0,
      checkAmount: 0,
      debtAmount: 0,
      discountAmount: 0,
      type: item.type,
      userId: "",
      registrationTime: item.date,
      customerNumber: "",
      invoiceDate: item.date,
      invoiceNumber: item.visaNumber,
      accountType: "",
      invoice: { ...item },
      matched: true,
      issue: item.note || "",
      category: "فيزا",
      categoryIssue: ""
    }));
    return <Suspense fallback={<PlatformLoading />}><Reconciliation2 initialStage="stageC" visaInvoices={visaInvoices} initialStageDBankInputs={stageDBankInputs} onSafeActivities={onSafeActivities} onBack={() => setPage("main")} /></Suspense>;
  }
  const togglePendingSelect = (key: string) => {
    setSelectedPendingKeys(prev => {
      const n = new Set(prev);
      if (n.has(key)) n.delete(key); else n.add(key);
      return n;
    });
  };
  const allPendingSelected = pendingRows.length > 0 && pendingRows.every((r:any) => selectedPendingKeys.has(transactionPairKey(r.cashier, r.bank)));
  const currentPagePendingKeys = pagedPendingRows.map((row: any) => transactionPairKey(row.cashier, row.bank));
  const allCurrentPagePendingSelected = currentPagePendingKeys.length > 0 &&
    currentPagePendingKeys.every(key => selectedPendingKeys.has(key));
  const togglePendingSelectAll = () => {
    if (allPendingSelected) { setSelectedPendingKeys(new Set()); return; }
    setSelectedPendingKeys(new Set(pendingRows.map((r:any) => transactionPairKey(r.cashier, r.bank))));
  };
  const toggleCurrentPagePendingSelection = () => {
    setSelectedPendingKeys(previous => {
      const next = new Set(previous);
      if (allCurrentPagePendingSelected) currentPagePendingKeys.forEach(key => next.delete(key));
      else currentPagePendingKeys.forEach(key => next.add(key));
      return next;
    });
  };
  const handleSaveSelectedPending = () => {
    const pairs = pendingRows
      .filter((r:any) => selectedPendingKeys.has(transactionPairKey(r.cashier, r.bank)))
      .map((r:any) => ({ cashier: r.cashier as CashierRow, bank: r.bank as BankRow }));
    handleSaveMatchesBulk(pairs);
  };

  const handleRetryMatch = () => {
    if (!pendingRows.length) {
      setToast("ما في مطابقات منتظرة لإعادة فرزها.");
      return;
    }

    if (!window.confirm(
      `إعادة فرز دقيقة لـ ${pendingRows.length} مطابقة منتظرة.\n` +
      `رح نفك كل الروابط الحالية ونعيد المطابقة من الصفر بسعر تنافسي لكل فاتورة وحوالة.\n` +
      `النتيجة بتضل بقائمة "منتظر" — ما رح تنتقل للمؤكدة إلا إذا ضغطتي حفظ.`
    )) return;

    clearMatchUndo();
    const pendingPairKeys = new Set(pendingRows.map((r:any) => transactionPairKey(r.cashier, r.bank)));

    setRejected(prev => {
      const next = new Set(prev);
      pendingPairKeys.forEach(key => next.delete(key));
      return next;
    });

    setToast(`تمت إعادة الفرز التنافسي لـ ${pendingRows.length} مطابقة منتظرة`);
  };

  const handleMoveToManual = (pairs: Array<{ cashier: CashierRow; bank: BankRow }>) => {
    if (!pairs.length) return;
    clearMatchUndo();
    setRejected(prev => {
      const next = new Set(prev);
      pairs.forEach(({ cashier, bank }) => next.add(transactionPairKey(cashier, bank)));
      return next;
    });
    setSelectedPendingKeys(new Set());
    setPage("manual");
  };
  const handleMoveSelectedPendingToManual = () => {
    const pairs = pendingRows
      .filter((r:any) => selectedPendingKeys.has(transactionPairKey(r.cashier, r.bank)))
      .map((r:any) => ({ cashier: r.cashier as CashierRow, bank: r.bank as BankRow }));
    handleMoveToManual(pairs);
  };
  const handleMoveAllPendingToManual = () => {
    if (!pendingRows.length) return;
    if (!window.confirm(`نقل كل المطابقات المنتظرة (${pendingRows.length}) للمطابقة اليدوية؟`)) return;
    handleMoveToManual(pendingRows.map((r:any) => ({ cashier: r.cashier as CashierRow, bank: r.bank as BankRow })));
  };

  const hasLoadedFiles = Boolean(
    bankHeaders.length || cashHeaders.length ||
    walletBankHeaders.length || walletCashHeaders.length ||
    jawwalBankHeaders.length || jawwalCashHeaders.length ||
    customPlatforms.some(platform => platform.transfer || platform.invoice) ||
    stageAInvoices.length
  );

  return (
    <div dir="rtl" className="min-h-screen bg-background text-foreground p-6">
      {referenceLocatorDialog}
      <div className="max-w-6xl mx-auto space-y-5">

        {toast && <div className="fixed bottom-5 left-5 z-50 flex items-center gap-2 rounded-lg bg-slate-900 px-4 py-3 text-sm text-white shadow-xl"><Check className="h-4 w-4 text-green-300" />{toast}</div>}

        {page === "upload" && (
          <header className="flex flex-wrap items-center justify-between gap-4 rounded-2xl border bg-white p-5 shadow-sm">
            <div>
              <p className="text-xs font-semibold text-blue-700">الخطوة ١ من ٢</p>
              <h1 className="mt-1 text-xl font-bold">رفع الملفات وتجهيزها</h1>
              <p className="mt-1 text-sm text-muted-foreground">ارفع ملفات المنصات المطلوبة؛ ستتم قراءتها وتجهيزها قبل فتح شاشة المطابقة.</p>
            </div>
            <button type="button" disabled={activeFileImports > 0} onClick={() => setPage("main")} className="rounded-lg border px-4 py-2 text-sm font-medium hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-50">
              العودة دون متابعة
            </button>
          </header>
        )}

        {page !== "upload" && <div className="flex flex-wrap items-center gap-3 justify-between">
          <div><h1 className="text-xl font-bold">Taswiyah Hub Smart Reconciliation Platform</h1><p className="mt-1 text-xs text-muted-foreground">مراجعة مالية أسرع، بقرارات قابلة للتتبع</p></div>
          <div className="flex max-w-5xl flex-wrap items-center justify-end gap-2">
            <button type="button" onClick={() => setPage("recon2")} className="flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50">
              <ChevronLeft className="h-4 w-4" />العودة إلى المنصة الأولى
            </button>
            <button onClick={() => setPage("assist")}
              className="flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700 transition-colors">
              <Sparkles className="h-4 w-4"/>مساعد المطابقة
            </button>
            <button onClick={() => setPage("manual")}
              className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg text-sm font-medium hover:bg-green-700 transition-colors">
              <Save className="w-4 h-4"/>المطابقة اليدوية (تؤكد فوراً)
            </button>
            <button onClick={() => setPage("upload")}
              className="flex items-center gap-2 rounded-lg border border-blue-200 bg-blue-50 px-4 py-2 text-sm font-medium text-blue-800 hover:bg-blue-100">
              <Upload className="h-4 w-4"/>رفع وإدارة الملفات
            </button>
            <details className="relative">
              <summary className="flex cursor-pointer list-none items-center gap-2 rounded-lg border border-slate-200 bg-white px-4 py-2 text-sm font-medium text-slate-700 shadow-sm hover:bg-slate-50">
                خيارات إضافية <ChevronDown className="h-4 w-4" />
              </summary>
              <div className="absolute left-0 z-30 mt-2 flex w-72 flex-col gap-2 rounded-xl border bg-white p-3 shadow-xl">
                <button onClick={() => setShowSettings(value => !value)} aria-expanded={showSettings} className={`rounded-lg border px-3 py-2 text-right text-xs font-semibold transition-colors ${showSettings ? "border-indigo-200 bg-indigo-50 text-indigo-800" : "text-slate-700 hover:bg-slate-50"}`}>إعدادات التسامح للمطابقة</button>
                <button onClick={handleUndoMatch} disabled={!canUndoMatch} title="تراجع عن آخر إجراء (Ctrl+Z)" aria-label="تراجع عن آخر إجراء"
                  className="flex items-center gap-1.5 rounded-lg border bg-white px-3 py-2 text-right text-xs font-medium hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-40">
                  <RotateCcw className="h-3.5 w-3.5"/>تراجع عن آخر إجراء{undoMatchCount > 1 ? ` (${undoMatchCount})` : ""}
                </button>
                <ProjectsBar onSave={saveProject} onLoad={loadProject} onDelete={deleteProject} />
                <button type="button" onClick={() => { setPage("main"); setResults(previous => previous ?? []); setTab("review"); }}
                  className="flex items-center gap-1.5 rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 text-xs font-medium text-amber-800 hover:bg-amber-100">
                  <Star className="h-3.5 w-3.5" />المراجعة ({reviewItems.length})
                </button>
                <div className="relative">
                  <button onClick={() => setShowDeleteMenu(value => !value)}
                    className="flex w-full items-center gap-1 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-right text-xs text-red-600 hover:bg-red-100">
                    <Trash2 className="h-3.5 w-3.5"/>أدوات الحذف <ChevronDown className="h-3 w-3"/>
                  </button>
                  {showDeleteMenu && <div className="mt-1 rounded-lg border bg-white py-1 text-xs shadow-lg">
                    <button onClick={handleClearSavedMatches} disabled={!savedMatches.length} className="flex w-full items-center gap-2 px-3 py-2.5 text-right transition-colors hover:bg-red-50 disabled:opacity-40"><Trash2 className="h-3.5 w-3.5 text-red-600"/><span>مسح المطابقات المؤكدة</span><span className="mr-auto text-muted-foreground">{savedMatches.length}</span></button>
                    <button onClick={() => handleClearSession()} className="flex w-full items-center gap-2 border-t px-3 py-2.5 text-right transition-colors hover:bg-amber-50"><RotateCcw className="h-3.5 w-3.5 text-amber-600"/><span>مسح الجلسة الحالية كاملة</span></button>
                    <button onClick={() => handleClearSession(true)} disabled={!heldItems.length} className="flex w-full items-center gap-2 border-t px-3 py-2.5 text-right transition-colors hover:bg-blue-50 disabled:opacity-40"><RotateCcw className="h-3.5 w-3.5 text-blue-600"/><span>مسح الجلسة مع إبقاء المعلّقات</span><span className="mr-auto text-muted-foreground">{heldItems.length}</span></button>
                    <button onClick={handleDeleteAllProjects} className="flex w-full items-center gap-2 border-t px-3 py-2.5 text-right transition-colors hover:bg-red-50"><FolderMinus className="h-3.5 w-3.5 text-red-600"/><span>حذف كل المشاريع المحفوظة</span></button>
                  </div>}
                </div>
              </div>
            </details>
          </div>
        </div>}

        {page !== "upload" && sessionRestoredNotice && (
          <div className="flex items-center gap-3 bg-green-50 border border-green-200 rounded-xl px-4 py-3">
            <Shield className="w-5 h-5 text-green-600 shrink-0"/>
            <p className="flex-1 text-sm font-medium text-green-800">
              تم استرجاع شغلك من آخر جلسة تلقائياً ✓ (الملفات، المطابقات، الفيزا، المعلقات)
            </p>
            <button onClick={() => setSessionRestoredNotice(false)} className="p-1 text-green-500 hover:text-green-800 transition-colors">
              <X className="w-4 h-4"/>
            </button>
          </div>
        )}

        {page !== "upload" && visaItems.length > 0 && (
          <div className="flex items-start gap-3 bg-teal-50 border border-teal-200 rounded-xl px-4 py-3">
            <CreditCard className="w-5 h-5 text-teal-600 shrink-0 mt-0.5"/>
            <div className="flex-1">
              <p className="text-sm font-semibold text-teal-800">{visaItems.length} عنصر منقول للفيزا</p>
              <p className="text-xs text-teal-600 mt-0.5">اضغط على تبويب "الفيزا" بالأسفل لمراجعتها</p>
            </div>
            <button onClick={() => setPage("recon2Visa")} className="shrink-0 rounded-lg bg-teal-600 px-3 py-2 text-xs font-medium text-white hover:bg-teal-700">
              ترحيل على C
            </button>
          </div>
        )}

        {error && <div className="p-3 bg-red-100 text-red-800 rounded-lg text-sm">{error}</div>}

        {page === "upload" && activeFileImports > 0 && (
          <div role="status" aria-live="polite" className="flex items-center gap-3 rounded-xl border border-blue-200 bg-blue-50 px-4 py-3 text-sm text-blue-900">
            <span className="h-5 w-5 animate-spin rounded-full border-2 border-blue-200 border-t-blue-700" />
            <span>جارٍ قراءة الملفات وتجهيز البيانات في الخلفية… ({activeFileImports} عملية)</span>
          </div>
        )}
        {page === "upload" && activeFileImports === 0 && hasLoadedFiles && (
          <div role="status" className="flex items-center gap-3 rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-900">
            <Check className="h-5 w-5 shrink-0" />
            <span className="flex-1">انتهى تجهيز الملفات المحمّلة. يمكنك الآن الانتقال إلى المطابقة العامة.</span>
            <button type="button" onClick={() => { setTab("saved"); setPage("main"); }} className="shrink-0 rounded-lg bg-emerald-700 px-4 py-2 font-semibold text-white hover:bg-emerald-800">
              اذهب للمطابقة العامة
            </button>
          </div>
        )}

        {page !== "upload" && showSettings && <section className="rounded-2xl border border-indigo-100 bg-gradient-to-l from-indigo-50/80 via-white to-white p-5 shadow-sm">
          <div className="flex flex-wrap items-center gap-5">
            <div className="min-w-56 flex-1">
              <p className="text-sm font-bold text-slate-900">إعدادات التسامح للمطابقة</p>
              <p className="mt-1 text-xs leading-5 text-slate-600">حدد الحد النسبي لفرق رسوم التحويل الذي تسمح به المطابقة الآلية.</p>
            </div>
            <label className="flex items-center gap-3 rounded-xl border border-indigo-100 bg-white px-4 py-3 text-sm font-semibold text-slate-700 shadow-sm">
              <span>نسبة التسامح</span>
              <span className="flex items-center gap-1 rounded-lg border bg-slate-50 px-2 py-1">
                <input type="number" min="0" max="10" step="0.1" value={amountTolerancePercent} onChange={e => setAmountTolerancePercent(Math.min(10, Math.max(0, Number(e.target.value) || 0)))} className="w-16 bg-transparent text-center font-mono outline-none" aria-label="نسبة التسامح للمطابقة" />
                <span className="text-muted-foreground">%</span>
              </span>
            </label>
            <span className="text-[11px] text-muted-foreground">من 0% إلى 10%</span>
          </div>
        </section>}

        {page === "upload" && <>
        <section className="rounded-xl border border-slate-200 bg-white p-4">
          <h2 className="text-sm font-semibold">أسماء المنصات</h2>
          <p className="mt-1 text-xs text-muted-foreground">الاسم نفسه يربط فواتير المنصة بحوالاتها عند المطابقة.</p>
          <div className="mt-3 grid gap-3 sm:grid-cols-3">
            {([
              ["bank", "المنصة الرئيسية (بنك فلسطين سابقاً)"],
              ["wallet", "كشف بال بي"],
              ["jawwal", "كشف جوال بي"],
            ] as const).map(([key, label]) => <label key={key} className="text-xs font-medium">
              {label}
              <input value={platformNames[key]} onChange={(event) => setPlatformNames((current) => ({ ...current, [key]: event.target.value }))} className="mt-1 w-full rounded-lg border bg-slate-50 px-3 py-2 text-sm" />
            </label>)}
          </div>
        </section>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
          <div className="bg-card border p-4 rounded-xl space-y-4">
            <div className="flex items-center justify-between">
               <h2 className="font-semibold text-sm">كشف حوالات {platformNames.bank}</h2>
              {(bankHeaders.length > 0 || bankFileSnapshots.length > 0) && (
                <button onClick={() => handleClearPlatform("bank")} title="مسح بيانات هذه المنصة فقط"
                  className="flex items-center gap-1 text-[11px] text-muted-foreground hover:text-red-600 transition-colors">
                  <Trash2 className="w-3 h-3"/>مسح هذه المنصة
                </button>
              )}
            </div>
            <DropZone file={bankFile} fileCount={bankFileCount} onFiles={(files, append) => trackFileImport(() => loadBank(files, append))} onClear={() => { setBankFile(null); setBankFileCount(0); setBankH([]); setBankRows([]); setResults(null); }} headers={bankHeaders} rows={bankRowsRaw} />
            {bankHeaders.length > 0 && (
              <div className="grid grid-cols-2 gap-2">
                <Sel label="التاريخ" headers={bankHeaders} value={bankMap.date} onChange={v => setBankMap(m => ({ ...m, date: v }))} />
                <Sel label="الإيضاحات / البيان" headers={bankHeaders} value={bankMap.desc} onChange={v => setBankMap(m => ({ ...m, desc: v }))} />
                <Sel label="المبالغ المستلمة (مدين)" headers={bankHeaders} value={bankMap.debit} onChange={v => setBankMap(m => ({ ...m, debit: v }))} />
                <Sel label="المبالغ المدفوعة (دائن)" headers={bankHeaders} value={bankMap.credit} onChange={v => setBankMap(m => ({ ...m, credit: v }))} />
                <Sel label="نوع الحساب (بنك فلسطين / محفظة...)" headers={bankHeaders} value={bankMap.accountType} onChange={v => setBankMap(m => ({ ...m, accountType: v }))} />
                <Sel label="رقم المرجع / الحوالة" headers={bankHeaders} value={bankMap.ref} onChange={v => setBankMap(m => ({ ...m, ref: v }))} />
              </div>
            )}
            {bankHeaders.length > 0 && (
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <input type="checkbox" checked={bankSwap} onChange={e=>setBankSwap(e.target.checked)} className="rounded"/>
                عكس المدين والدائن (إذا جاءت الأنواع مقلوبة)
              </label>
            )}
            {bankFileSnapshots.length > 0 && (
              <div className="relative">
                <button onClick={() => setShowBankSnapshots(v => !v)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors">
                  <RotateCcw className="w-3.5 h-3.5"/>استرجاع ملف سابق ({bankFileSnapshots.length})
                </button>
                {showBankSnapshots && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setShowBankSnapshots(false)} />
                    <div className="absolute left-0 mt-1 w-72 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs max-h-72 overflow-y-auto">
                      {bankFileSnapshots.map((s, i) => (
                        <button key={i} onClick={() => {
                          setBankFileCount(s.fileCount ?? 1);
                          setBankH(s.headers); setBankRows(s.rows);
                          setBankMap({ date:autoDetectBankDate(s.headers), desc:autoDetect(s.headers,HINTS.desc), debit:autoDetect(s.headers,HINTS.debit), credit:autoDetect(s.headers,HINTS.credit), accountType:autoDetect(s.headers,HINTS.accountType), ref:autoDetect(s.headers,HINTS.ref) });
                          setBankFileSessionId(id => id + 1);
                          setShowBankSnapshots(false); setResults(null);
                        }} className="w-full text-right px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 transition-colors">
                          <RotateCcw className="w-3 h-3 text-slate-400"/>
                          <div className="flex-1 min-w-0">
                            <div className="truncate font-medium">{s.label}</div>
                            <div className="text-[10px] text-muted-foreground">{s.savedAt} · {s.rows.length} صف</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>
          {!stageAInvoices.length && <div className="bg-card border p-4 rounded-xl space-y-4">
            <div className="flex items-center justify-between">
               <h2 className="font-semibold text-sm">كشف فواتير {platformNames.bank}</h2>
              {(cashHeaders.length > 0 || cashFileSnapshots.length > 0) && (
                <button onClick={() => handleClearPlatform("cash")} title="مسح بيانات هذه المنصة فقط"
                  className="flex items-center gap-1 text-[11px] text-muted-foreground hover:text-red-600 transition-colors">
                  <Trash2 className="w-3 h-3"/>مسح هذه المنصة
                </button>
              )}
            </div>
            <DropZone file={cashFile} fileCount={cashFileCount} onFiles={(files, append) => trackFileImport(() => loadCash(files, append))} onClear={() => { setCashFile(null); setCashFileCount(0); setCashH([]); setCashRows([]); setResults(null); }} headers={cashHeaders} rows={cashRowsRaw} />
            {cashHeaders.length > 0 && (
              <div className="grid grid-cols-2 gap-2">
                <Sel label="التاريخ" headers={cashHeaders} value={cashMap.date} onChange={v => setCashMap(m => ({ ...m, date: v }))} />
                <Sel label="البيان" headers={cashHeaders} value={cashMap.name} onChange={v => setCashMap(m => ({ ...m, name: v }))} />
                <Sel label="المبالغ المستلمة (مدين)" headers={cashHeaders} value={cashMap.debit} onChange={v => setCashMap(m => ({ ...m, debit: v }))} />
                <Sel label="المبالغ المدفوعة (دائن)" headers={cashHeaders} value={cashMap.credit} onChange={v => setCashMap(m => ({ ...m, credit: v }))} />
                <Sel label="نوع الحساب (بنك فلسطين / محفظة...)" headers={cashHeaders} value={cashMap.accountType} onChange={v => setCashMap(m => ({ ...m, accountType: v }))} />
                <Sel label="رقم المرجع / الحوالة" headers={cashHeaders} value={cashMap.ref} onChange={v => setCashMap(m => ({ ...m, ref: v }))} />
                <Sel label="رقم المستخدم" headers={cashHeaders} value={cashMap.userId} onChange={v => setCashMap(m => ({ ...m, userId: v }))} />
              </div>
            )}
            {cashHeaders.length > 0 && (
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <input type="checkbox" checked={cashSwap} onChange={e=>setCashSwap(e.target.checked)} className="rounded"/>
                عكس المدين والدائن (إذا جاءت الأنواع مقلوبة)
              </label>
            )}
            {cashFileSnapshots.length > 0 && (
              <div className="relative">
                <button onClick={() => setShowCashSnapshots(v => !v)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors">
                  <RotateCcw className="w-3.5 h-3.5"/>استرجاع ملف سابق ({cashFileSnapshots.length})
                </button>
                {showCashSnapshots && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setShowCashSnapshots(false)} />
                    <div className="absolute left-0 mt-1 w-72 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs max-h-72 overflow-y-auto">
                      {cashFileSnapshots.map((s, i) => (
                        <button key={i} onClick={() => {
                          setCashFileCount(s.fileCount ?? 1);
                          setCashH(s.headers); setCashRows(s.rows);
                          setCashMap({ date:autoDetect(s.headers,HINTS.date), name:autoDetectInvoiceName(s.headers), debit:autoDetect(s.headers,HINTS.debit), credit:autoDetect(s.headers,HINTS.credit), accountType:autoDetect(s.headers,HINTS.accountType), ref:autoDetect(s.headers,HINTS.ref), userId:autoDetect(s.headers,HINTS.userId) });
                          setCashFileSessionId(id => id + 1);
                          setShowCashSnapshots(false); setResults(null);
                        }} className="w-full text-right px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 transition-colors">
                          <RotateCcw className="w-3 h-3 text-slate-400"/>
                          <div className="flex-1 min-w-0">
                            <div className="truncate font-medium">{s.label}</div>
                            <div className="text-[10px] text-muted-foreground">{s.savedAt} · {s.rows.length} صف</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>}
          <div className="bg-card border p-4 rounded-xl space-y-4">
            <div className="flex items-center justify-between">
               <h2 className="font-semibold text-sm">كشف حوالات {platformNames.wallet}</h2>
              {(walletBankHeaders.length > 0 || walletBankFileSnapshots.length > 0) && (
                <button onClick={() => handleClearPlatform("walletBank")} title="مسح بيانات هذه المنصة فقط"
                  className="flex items-center gap-1 text-[11px] text-muted-foreground hover:text-red-600 transition-colors">
                  <Trash2 className="w-3 h-3"/>مسح هذه المنصة
                </button>
              )}
            </div>
            <DropZone file={walletBankFile} fileCount={walletBankFileCount} onFiles={(files, append) => trackFileImport(() => loadWalletBank(files, append))} onClear={() => { setWalletBankFile(null); setWalletBankFileCount(0); setWalletBankHeaders([]); setWalletBankRows([]); setResults(null); }} headers={walletBankHeaders} rows={walletBankRows} />
            {walletBankHeaders.length > 0 && <div className="grid grid-cols-2 gap-2">
              <Sel label="التاريخ" headers={walletBankHeaders} value={walletBankMap.date} onChange={v => setWalletBankMap(m => ({ ...m, date:v }))} />
              <Sel label="الإيضاحات / البيان" headers={walletBankHeaders} value={walletBankMap.desc} onChange={v => setWalletBankMap(m => ({ ...m, desc:v }))} />
              <Sel label="المبالغ المستلمة (مدين)" headers={walletBankHeaders} value={walletBankMap.debit} onChange={v => setWalletBankMap(m => ({ ...m, debit:v }))} />
              <Sel label="المبالغ المدفوعة (دائن)" headers={walletBankHeaders} value={walletBankMap.credit} onChange={v => setWalletBankMap(m => ({ ...m, credit:v }))} />
              <Sel label="رقم المرجع / الحوالة" headers={walletBankHeaders} value={walletBankMap.ref} onChange={v => setWalletBankMap(m => ({ ...m, ref:v }))} />
            </div>}
            {walletBankHeaders.length > 0 && (
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <input type="checkbox" checked={walletBankSwap} onChange={e=>setWalletBankSwap(e.target.checked)} className="rounded"/>
                عكس المدين والدائن (إذا جاءت الأنواع مقلوبة)
              </label>
            )}
            {walletBankFileSnapshots.length > 0 && (
              <div className="relative">
                <button onClick={() => setShowWalletBankSnapshots(v => !v)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors">
                  <RotateCcw className="w-3.5 h-3.5"/>استرجاع ملف سابق ({walletBankFileSnapshots.length})
                </button>
                {showWalletBankSnapshots && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setShowWalletBankSnapshots(false)} />
                    <div className="absolute left-0 mt-1 w-72 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs max-h-72 overflow-y-auto">
                      {walletBankFileSnapshots.map((s, i) => (
                        <button key={i} onClick={() => {
                          setWalletBankFileCount(s.fileCount ?? 1);
                          setWalletBankHeaders(s.headers); setWalletBankRows(s.rows);
                          setWalletBankMap({ date:autoDetectBankDate(s.headers), desc:autoDetect(s.headers,HINTS.desc), debit:autoDetect(s.headers,HINTS.debit), credit:autoDetect(s.headers,HINTS.credit), ref:autoDetect(s.headers,HINTS.ref) });
                          setWalletBankFileSessionId(id => id + 1);
                          setShowWalletBankSnapshots(false); setResults(null);
                        }} className="w-full text-right px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 transition-colors">
                          <RotateCcw className="w-3 h-3 text-slate-400"/>
                          <div className="flex-1 min-w-0">
                            <div className="truncate font-medium">{s.label}</div>
                            <div className="text-[10px] text-muted-foreground">{s.savedAt} · {s.rows.length} صف</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>
          {!stageAInvoices.length && <div className="bg-card border p-4 rounded-xl space-y-4">
            <div className="flex items-center justify-between">
               <h2 className="font-semibold text-sm">كشف فواتير {platformNames.wallet}</h2>
              {(walletCashHeaders.length > 0 || walletCashFileSnapshots.length > 0) && (
                <button onClick={() => handleClearPlatform("walletCash")} title="مسح بيانات هذه المنصة فقط"
                  className="flex items-center gap-1 text-[11px] text-muted-foreground hover:text-red-600 transition-colors">
                  <Trash2 className="w-3 h-3"/>مسح هذه المنصة
                </button>
              )}
            </div>
            <DropZone file={walletCashFile} fileCount={walletCashFileCount} onFiles={(files, append) => trackFileImport(() => loadWalletCash(files, append))} onClear={() => { setWalletCashFile(null); setWalletCashFileCount(0); setWalletCashHeaders([]); setWalletCashRows([]); setResults(null); }} headers={walletCashHeaders} rows={walletCashRows} />
            {walletCashHeaders.length > 0 && <div className="grid grid-cols-2 gap-2">
              <Sel label="التاريخ" headers={walletCashHeaders} value={walletCashMap.date} onChange={v => setWalletCashMap(m => ({ ...m, date:v }))} />
              <Sel label="البيان" headers={walletCashHeaders} value={walletCashMap.name} onChange={v => setWalletCashMap(m => ({ ...m, name:v }))} />
              <Sel label="المبالغ المستلمة (مدين)" headers={walletCashHeaders} value={walletCashMap.debit} onChange={v => setWalletCashMap(m => ({ ...m, debit:v }))} />
              <Sel label="المبالغ المدفوعة (دائن)" headers={walletCashHeaders} value={walletCashMap.credit} onChange={v => setWalletCashMap(m => ({ ...m, credit:v }))} />
              <Sel label="رقم المرجع / الحوالة" headers={walletCashHeaders} value={walletCashMap.ref} onChange={v => setWalletCashMap(m => ({ ...m, ref:v }))} />
              <Sel label="رقم المستخدم" headers={walletCashHeaders} value={walletCashMap.userId} onChange={v => setWalletCashMap(m => ({ ...m, userId:v }))} />
            </div>}
            {walletCashHeaders.length > 0 && (
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <input type="checkbox" checked={walletCashSwap} onChange={e=>setWalletCashSwap(e.target.checked)} className="rounded"/>
                عكس المدين والدائن (إذا جاءت الأنواع مقلوبة)
              </label>
            )}
            {walletCashFileSnapshots.length > 0 && (
              <div className="relative">
                <button onClick={() => setShowWalletCashSnapshots(v => !v)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors">
                  <RotateCcw className="w-3.5 h-3.5"/>استرجاع ملف سابق ({walletCashFileSnapshots.length})
                </button>
                {showWalletCashSnapshots && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setShowWalletCashSnapshots(false)} />
                    <div className="absolute left-0 mt-1 w-72 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs max-h-72 overflow-y-auto">
                      {walletCashFileSnapshots.map((s, i) => (
                        <button key={i} onClick={() => {
                          setWalletCashFileCount(s.fileCount ?? 1);
                          setWalletCashHeaders(s.headers); setWalletCashRows(s.rows);
                          setWalletCashMap({ date:autoDetect(s.headers,HINTS.date), name:autoDetectInvoiceName(s.headers), debit:autoDetect(s.headers,HINTS.debit), credit:autoDetect(s.headers,HINTS.credit), ref:autoDetect(s.headers,HINTS.ref), userId:autoDetect(s.headers,HINTS.userId) });
                          setWalletCashFileSessionId(id => id + 1);
                          setShowWalletCashSnapshots(false); setResults(null);
                        }} className="w-full text-right px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 transition-colors">
                          <RotateCcw className="w-3 h-3 text-slate-400"/>
                          <div className="flex-1 min-w-0">
                            <div className="truncate font-medium">{s.label}</div>
                            <div className="text-[10px] text-muted-foreground">{s.savedAt} · {s.rows.length} صف</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>}
          <div className="bg-card border p-4 rounded-xl space-y-4">
            <div className="flex items-center justify-between">
               <h2 className="font-semibold text-sm">كشف حوالات {platformNames.jawwal}</h2>
              {(jawwalBankHeaders.length > 0 || jawwalBankFileSnapshots.length > 0) && (
                <button onClick={() => handleClearPlatform("jawwalBank")} title="مسح بيانات هذه المنصة فقط"
                  className="flex items-center gap-1 text-[11px] text-muted-foreground hover:text-red-600 transition-colors">
                  <Trash2 className="w-3 h-3"/>مسح هذه المنصة
                </button>
              )}
            </div>
            <DropZone file={jawwalBankFile} fileCount={jawwalBankFileCount} onFiles={(files, append) => trackFileImport(() => loadJawwalBank(files, append))} onClear={() => { setJawwalBankFile(null); setJawwalBankFileCount(0); setJawwalBankHeaders([]); setJawwalBankRows([]); setResults(null); }} headers={cleanedJawwalBankDataset.headers} rows={cleanedJawwalBankDataset.rows} />
            {jawwalBankHeaders.length > 0 && <div className="grid grid-cols-2 gap-2">
              <Sel label="التاريخ" headers={jawwalBankHeaders} value={jawwalBankMap.date} onChange={v => setJawwalBankMap(m => ({ ...m, date:v }))} />
              <Sel label="الإيضاحات / البيان" headers={jawwalBankHeaders} value={jawwalBankMap.desc} onChange={v => setJawwalBankMap(m => ({ ...m, desc:v }))} />
              <Sel label="عمود المبلغ الكلي" headers={jawwalBankHeaders} value={jawwalBankMap.totalAmount} onChange={v => setJawwalBankMap(m => ({ ...m, totalAmount:v }))} />
              <Sel label="المبلغ البديل عند إيقاف الكلي" headers={jawwalBankHeaders} value={jawwalBankMap.movementAmount} onChange={v => setJawwalBankMap(m => ({ ...m, movementAmount:v }))} />
              <Sel label="نوع الحركة" headers={jawwalBankHeaders} value={jawwalBankMap.movementType} onChange={v => setJawwalBankMap(m => ({ ...m, movementType:v }))} />
              <Sel label="عمود تحديد مدين / دائن" headers={jawwalBankHeaders} value={jawwalBankMap.direction} onChange={v => setJawwalBankMap(m => ({ ...m, direction:v }))} />
              <Sel label="المبالغ المستلمة (مدين)" headers={jawwalBankHeaders} value={jawwalBankMap.debit} onChange={v => setJawwalBankMap(m => ({ ...m, debit:v }))} />
              <Sel label="المبالغ المدفوعة (دائن)" headers={jawwalBankHeaders} value={jawwalBankMap.credit} onChange={v => setJawwalBankMap(m => ({ ...m, credit:v }))} />
              <Sel label="رقم المرجع / الحوالة" headers={jawwalBankHeaders} value={jawwalBankMap.ref} onChange={v => setJawwalBankMap(m => ({ ...m, ref:v }))} />
            </div>}
            {jawwalBankHeaders.length > 0 && <div className="flex flex-wrap gap-4">
              <label className="flex items-center gap-2 text-xs text-muted-foreground">
                <input type="checkbox" checked={jawwalBankOptions.useTotalAmount} onChange={e=>setJawwalBankOptions(options=>({...options,useTotalAmount:e.target.checked}))} className="rounded"/>
                استخدام المبلغ الكلي
              </label>
              <label className="flex items-center gap-2 text-xs text-muted-foreground">
                <input type="checkbox" checked={jawwalBankOptions.readDebitCredit} onChange={e=>setJawwalBankOptions(options=>({...options,readDebitCredit:e.target.checked}))} className="rounded"/>
                تحديد مستلم / مدفوع من عمود مدين / دائن
              </label>
            </div>}
            {jawwalBankHeaders.length > 0 && (
              <div className="space-y-1">
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <input type="checkbox" checked={jawwalBankSwap} onChange={e=>setJawwalBankSwap(e.target.checked)} className="rounded"/>
                عكس المدين والدائن تلقائيًا لهذا كشف جوال بي
              </label>
              <p className="text-[11px] text-muted-foreground">مفعّل افتراضيًا لعكس دلالة جوال بي؛ أزل التحديد إذا كان هذا الكشف بالاتجاه المعتاد. لا تتغير بيانات الملف الأصلي.</p>
              </div>
            )}
            {jawwalBankHeaders.length > 0 && jawwalCashHeaders.length > 0 && (
              <p className="rounded-lg border border-cyan-200 bg-cyan-50 px-3 py-2 text-xs leading-5 text-cyan-900">
                تتم مطابقة ملفي جوال بي عند تطابق رقم الجوال والمبلغ معاً؛ ولا يُستخدم الاسم في المطابقة. الحركات التي لا يتحقق فيها الشرطان تبقى للمراجعة.
              </p>
            )}
            {jawwalBankFileSnapshots.length > 0 && (
              <div className="relative">
                <button onClick={() => setShowJawwalBankSnapshots(v => !v)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors">
                  <RotateCcw className="w-3.5 h-3.5"/>استرجاع ملف سابق ({jawwalBankFileSnapshots.length})
                </button>
                {showJawwalBankSnapshots && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setShowJawwalBankSnapshots(false)} />
                    <div className="absolute left-0 mt-1 w-72 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs max-h-72 overflow-y-auto">
                      {jawwalBankFileSnapshots.map((s, i) => (
                        <button key={i} onClick={() => {
                          setJawwalBankFileCount(s.fileCount ?? 1);
                          setJawwalBankHeaders(s.headers); setJawwalBankRows(s.rows);
                          setJawwalBankSwap(true);
                          const totalAmount = autoDetect(s.headers, HINTS.totalAmount);
                          const direction = autoDetect(s.headers, HINTS.direction);
                          const debit = autoDetect(s.headers, HINTS.debit);
                          const credit = autoDetect(s.headers, HINTS.credit);
                          setJawwalBankMap({
                            date: autoDetectBankDate(s.headers),
                            desc: autoDetect(s.headers, ["بيان القيد", ...HINTS.desc]),
                            debit: debit === direction ? "" : debit,
                            credit: credit === direction ? "" : credit,
                            ref: autoDetect(s.headers, HINTS.ref),
                            totalAmount,
                            movementAmount: autoDetect(s.headers, HINTS.movementAmount),
                            direction,
                            movementType: autoDetect(s.headers, ["نوع الحركة", "transaction type"]),
                          });
                          setJawwalBankFileSessionId(id => id + 1);
                          setShowJawwalBankSnapshots(false); setResults(null);
                        }} className="w-full text-right px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 transition-colors">
                          <RotateCcw className="w-3 h-3 text-slate-400"/>
                          <div className="flex-1 min-w-0">
                            <div className="truncate font-medium">{s.label}</div>
                            <div className="text-[10px] text-muted-foreground">{s.savedAt} · {s.rows.length} صف</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>
          {!stageAInvoices.length && <div className="bg-card border p-4 rounded-xl space-y-4">
            <div className="flex items-center justify-between">
               <h2 className="font-semibold text-sm">كشف فواتير {platformNames.jawwal}</h2>
              {(jawwalCashHeaders.length > 0 || jawwalCashFileSnapshots.length > 0) && (
                <button onClick={() => handleClearPlatform("jawwalCash")} title="مسح بيانات هذه المنصة فقط"
                  className="flex items-center gap-1 text-[11px] text-muted-foreground hover:text-red-600 transition-colors">
                  <Trash2 className="w-3 h-3"/>مسح هذه المنصة
                </button>
              )}
            </div>
            <DropZone file={jawwalCashFile} fileCount={jawwalCashFileCount} onFiles={(files, append) => trackFileImport(() => loadJawwalCash(files, append))} onClear={() => { setJawwalCashFile(null); setJawwalCashFileCount(0); setJawwalCashHeaders([]); setJawwalCashRows([]); setResults(null); }} headers={jawwalCashHeaders} rows={jawwalCashRows} />
            {jawwalCashHeaders.length > 0 && <div className="grid grid-cols-2 gap-2">
              <Sel label="التاريخ" headers={jawwalCashHeaders} value={jawwalCashMap.date} onChange={v => setJawwalCashMap(m => ({ ...m, date:v }))} />
              <Sel label="البيان" headers={jawwalCashHeaders} value={jawwalCashMap.name} onChange={v => setJawwalCashMap(m => ({ ...m, name:v }))} />
              <Sel label="المبالغ المستلمة (مدين)" headers={jawwalCashHeaders} value={jawwalCashMap.debit} onChange={v => setJawwalCashMap(m => ({ ...m, debit:v }))} />
              <Sel label="المبالغ المدفوعة (دائن)" headers={jawwalCashHeaders} value={jawwalCashMap.credit} onChange={v => setJawwalCashMap(m => ({ ...m, credit:v }))} />
              <Sel label="رقم المرجع / الحوالة" headers={jawwalCashHeaders} value={jawwalCashMap.ref} onChange={v => setJawwalCashMap(m => ({ ...m, ref:v }))} />
              <Sel label="رقم المستخدم" headers={jawwalCashHeaders} value={jawwalCashMap.userId} onChange={v => setJawwalCashMap(m => ({ ...m, userId:v }))} />
            </div>}
            {jawwalCashHeaders.length > 0 && (
              <label className="flex items-center gap-2 text-xs text-muted-foreground cursor-pointer w-fit">
                <input type="checkbox" checked={jawwalCashSwap} onChange={e=>setJawwalCashSwap(e.target.checked)} className="rounded"/>
                عكس المدين والدائن (إذا جاءت الأنواع مقلوبة)
              </label>
            )}
            {jawwalCashFileSnapshots.length > 0 && (
              <div className="relative">
                <button onClick={() => setShowJawwalCashSnapshots(v => !v)}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors">
                  <RotateCcw className="w-3.5 h-3.5"/>استرجاع ملف سابق ({jawwalCashFileSnapshots.length})
                </button>
                {showJawwalCashSnapshots && (
                  <>
                    <div className="fixed inset-0 z-10" onClick={() => setShowJawwalCashSnapshots(false)} />
                    <div className="absolute left-0 mt-1 w-72 bg-card border border-border rounded-lg shadow-lg z-20 py-1 text-xs max-h-72 overflow-y-auto">
                      {jawwalCashFileSnapshots.map((s, i) => (
                        <button key={i} onClick={() => {
                          setJawwalCashFileCount(s.fileCount ?? 1);
                          setJawwalCashHeaders(s.headers); setJawwalCashRows(s.rows);
                          setJawwalCashMap({ date:autoDetect(s.headers,HINTS.date), name:autoDetectInvoiceName(s.headers), debit:autoDetect(s.headers,HINTS.debit), credit:autoDetect(s.headers,HINTS.credit), ref:autoDetect(s.headers,HINTS.ref), userId:autoDetect(s.headers,HINTS.userId) });
                          setJawwalCashFileSessionId(id => id + 1);
                          setShowJawwalCashSnapshots(false); setResults(null);
                        }} className="w-full text-right px-4 py-2.5 hover:bg-slate-50 flex items-center gap-2 transition-colors">
                          <RotateCcw className="w-3 h-3 text-slate-400"/>
                          <div className="flex-1 min-w-0">
                            <div className="truncate font-medium">{s.label}</div>
                            <div className="text-[10px] text-muted-foreground">{s.savedAt} · {s.rows.length} صف</div>
                          </div>
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </div>
            )}
          </div>}
        </div>

        <section className="space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <h2 className="text-sm font-semibold">منصات وملفات إضافية</h2>
              <p className="mt-1 text-xs text-muted-foreground">أضف أي عدد من المنصات، وسمّ كل منصة واختر دور كل ملف في المطابقة.</p>
            </div>
            <button onClick={() => {
              const id = nextCustomPlatformId.current++;
              setCustomPlatforms((current) => [...current, { id, name: "", transfer: null, invoice: null }]);
            }} className="inline-flex items-center gap-2 rounded-lg bg-blue-600 px-4 py-2 text-sm font-semibold text-white hover:bg-blue-700">
              <Plus className="h-4 w-4" />إضافة منصة
            </button>
          </div>
          {customPlatforms.map((platform) => <article key={platform.id} className="rounded-xl border border-blue-100 bg-white p-4 shadow-sm">
            <div className="mb-4 flex flex-wrap items-end gap-3">
              <label className="min-w-56 flex-1 text-xs font-semibold">اسم المنصة
                <CustomPlatformNameInput
                  value={platform.name}
                  onCommit={name => setCustomPlatforms(current => current.map(item => item.id === platform.id ? { ...item, name } : item))}
                />
              </label>
              <button type="button" onClick={() => setCustomPlatforms((current) => current.filter((item) => item.id !== platform.id))} className="inline-flex items-center gap-1 rounded-lg border border-red-200 px-3 py-2 text-xs text-red-700 hover:bg-red-50"><Trash2 className="h-3.5 w-3.5" />حذف المنصة</button>
            </div>
            <div className="grid gap-4 lg:grid-cols-2">
              {(["transfer", "invoice"] as const).map((slot) => {
                const source = platform[slot];
                const updateSource = (updater: (value: PlatformSource) => PlatformSource) => setCustomPlatforms((current) => current.map((item) => {
                  if (item.id !== platform.id || !item[slot]) return item;
                  return { ...item, [slot]: updater(item[slot]!) };
                }));
                return <div key={slot} className="space-y-3 rounded-lg border border-slate-200 p-3">
                  <h3 className="text-xs font-bold">{slot === "transfer" ? "ملف الحوالات" : "ملف الفواتير"}</h3>
                  <label className="flex cursor-pointer items-center gap-2 rounded-lg border border-dashed bg-slate-50 px-3 py-3 text-xs hover:bg-slate-100">
                    <Upload className="h-4 w-4 text-blue-600" />
                    <span className="flex-1 truncate">{source?.fileName ?? "اختر ملفاً"}</span>
                    <input type="file" multiple accept=".xlsx,.xls,.csv" className="hidden" onChange={(event) => { const files = Array.from(event.target.files ?? []); if (files.length) void trackFileImport(() => loadCustomSource(platform.id, slot, files)); event.currentTarget.value = ""; }} />
                  </label>
                  {source && <>
                    {source.rows.length > 0 && <button type="button" onClick={() => downloadCleanedSheet(source.fileName, source.headers, source.rows)} className="inline-flex items-center gap-1 rounded-md bg-emerald-50 px-2 py-1.5 text-xs font-medium text-emerald-700 hover:bg-emerald-100"><Download className="h-3.5 w-3.5"/>تنزيل Excel منظف</button>}
                    <label className="block text-xs font-medium">نوع الملف
                      <select value={source.role} onChange={(event) => updateSource((value) => ({ ...value, role: event.target.value as PlatformRole }))} className="mt-1 w-full rounded-lg border bg-white px-3 py-2">
                        <option value="invoice">فواتير</option><option value="transfer">حوالات</option><option value="other">شيء آخر</option>
                      </select>
                    </label>
                    {source.role === "other" && <div className="grid gap-2 sm:grid-cols-2">
                      <label className="text-xs font-medium">اسم النوع
                        <input value={source.otherName} placeholder="مثال: سند قبض" onChange={(event) => updateSource((value) => ({ ...value, otherName: event.target.value }))} className="mt-1 w-full rounded-lg border bg-white px-3 py-2" />
                      </label>
                      <label className="text-xs font-medium">يعامل كـ
                        <select value={source.otherRole} onChange={(event) => updateSource((value) => ({ ...value, otherRole: event.target.value as "invoice" | "transfer" }))} className="mt-1 w-full rounded-lg border bg-white px-3 py-2">
                          <option value="invoice">فواتير</option><option value="transfer">حوالات</option>
                        </select>
                      </label>
                    </div>}
                    {source.headers.length > 0 && <div className="grid grid-cols-2 gap-2">
                      <Sel label="التاريخ" headers={source.headers} value={source.map.date} onChange={(value) => updateSource((item) => ({ ...item, map: { ...item.map, date: value } }))} />
                      <Sel label="البيان / الاسم" headers={source.headers} value={source.map.description} onChange={(value) => updateSource((item) => ({ ...item, map: { ...item.map, description: value } }))} />
                      <Sel label="المدين / المبلغ المستلم" headers={source.headers} value={source.map.debit} onChange={(value) => updateSource((item) => ({ ...item, map: { ...item.map, debit: value } }))} />
                      <Sel label="الدائن / المبلغ المدفوع" headers={source.headers} value={source.map.credit} onChange={(value) => updateSource((item) => ({ ...item, map: { ...item.map, credit: value } }))} />
                      <Sel label="رقم المرجع" headers={source.headers} value={source.map.ref} onChange={(value) => updateSource((item) => ({ ...item, map: { ...item.map, ref: value } }))} />
                      {source.role !== "transfer" && <Sel label="رقم المستخدم" headers={source.headers} value={source.map.userId ?? ""} onChange={(value) => updateSource((item) => ({ ...item, map: { ...item.map, userId: value } }))} />}
                    </div>}
                    {source.role === "other" && !!source.otherName.trim() && <p className="text-[11px] text-slate-500">سيتم التعامل مع «{source.otherName}» كـ{source.otherRole === "invoice" ? "فواتير" : "حوالات"}.</p>}
                  </>}
                </div>
              })}
            </div>
          </article>)}
        </section>
        </>}

        {page === "upload" ? (
          <div className="flex flex-wrap items-center gap-3">
            <button type="button" onClick={() => { setTab("saved"); setPage("main"); }} disabled={!hasLoadedFiles || activeFileImports > 0}
              className="inline-flex items-center gap-2 rounded-lg bg-blue-700 px-6 py-3 font-semibold text-white hover:bg-blue-800 disabled:cursor-not-allowed disabled:opacity-40">
              <Check className="h-4 w-4" />الانتقال إلى المطابقة العامة
            </button>
            <span className="text-xs text-muted-foreground">يمكنك رفع ملفات منصات متعددة؛ لكل ملف مؤشرات أعمدة وإعدادات مستقلة.</span>
          </div>
        ) : <div className="flex flex-wrap items-center gap-3">
          <button onClick={run} disabled={loading}
            className="px-6 py-2.5 bg-blue-600 text-white rounded-lg disabled:opacity-40 font-medium hover:bg-blue-700 transition-colors flex items-center gap-2">
             {loading ? <><span className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />جاري المعالجة...</> : "تشغيل المطابقة"}
          </button>
          {loading && <div role="status" className="flex items-center gap-2 rounded-lg border border-amber-200 bg-white px-3 py-2">
            <img src="/taswiyah-logo.png" alt="" className="h-8 w-8 rounded-lg object-contain" />
            <span className="text-sm font-semibold text-slate-900">جارٍ تشغيل المطابقة...</span>
          </div>}
          {(results || bankSplits.length > 0 || cutoffBatches.length > 0 || clearingGroups.length > 0) && (
            <button onClick={() => {
                try {
                  doExport(
                    results ?? [], savedMatches, visaItems, heldItems,
                    reportBankRows, reportCashierRows, buildResumeData(),
                    nonReconciliationBankRows.map(({ bank }) => bank), bankSplits, cutoffBatches, clearingGroups, activeCashier, activeBank, rowNotes,
                    stageAInvoices, stageADatabaseSummary
                  );
                } catch (e) {
                  console.error("فشل تصدير الملف:", e);
                  setError(e instanceof Error ? e.message : "صار خطأ أثناء تصدير الملف. جرّب تاني.");
                }
              }}
              className="px-4 py-2.5 border rounded-lg flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white border-transparent transition-colors font-medium text-sm">
              <Download className="w-4 h-4" />تصدير Excel
            </button>
          )}
          <label className="px-4 py-2.5 border border-teal-300 bg-teal-50 text-teal-700 rounded-lg flex items-center gap-1.5 hover:bg-teal-100 transition-colors font-medium text-sm cursor-pointer">
            <Upload className="w-4 h-4" />استيراد ملف سابق
            <input type="file" accept=".xlsx,.xls" className="hidden"
              onChange={e => { const f = e.target.files?.[0]; if (f) handleImportResumeFile(f); e.target.value = ""; }}/>
          </label>
        </div>}

        {page !== "upload" && results && stats && (
          <div className="space-y-4">
            <div className="flex gap-2 overflow-x-auto pb-1">
               {([
                 { id: "saved", label: "مؤكد", count: stats.saved, color: "text-green-700" },
                 { id: "review", label: "مراجعة ★", count: stats.review, color: "text-amber-700" },
                 { id: "pending", label: "منتظر", count: stats.pending, color: "text-blue-700" },
                 { id: "manual", label: "يدوي", count: stats.manual, color: "text-orange-700" },
                 { id: "visa", label: "فيزا", count: stats.visa, color: "text-cyan-700" },
                 { id: "jawwalPay", label: "جوال بي", count: stats.jawwalPay, color: "text-cyan-700" },
                 { id: "mahmoudWallet", label: "محفظة شخصية", count: stats.mahmoudWallet, color: "text-violet-700" },
                 { id: "held", label: "معلّقة", count: stats.held, color: "text-orange-700" },
                 { id: "nonReconciliationBank", label: "معلقات البنك غير التابعة للمطابقة", count: stats.nonReconciliationBank, color: "text-amber-700" },
                 { id: "uCashier", label: "فواتير بلا حوالات مطابقة", count: stats.uCashier, color: "text-red-700" },
                 { id: "uBank", label: "حوالات بلا فواتير مسجلة", count: stats.uBank, color: "text-red-700" },
              ] as const).filter(t => t.count > 0).map(t => (
                <button key={t.id} onClick={() => setTab(t.id as TabId)}
                  className={`min-h-[62px] min-w-[110px] shrink-0 rounded-2xl border px-3 py-2 text-right transition-all ${tab === t.id ? "border-blue-500 bg-blue-50/70" : "bg-card hover:bg-muted/30"}`}>
                  <span className="block truncate text-[10px] leading-4 text-muted-foreground">{t.label}</span>
                  <span className={`mt-0.5 block font-mono text-base font-bold leading-5 ${t.color}`}>{t.count}</span>
                </button>
              ))}
            </div>

            <div className="border rounded-xl bg-card overflow-hidden">
              <FilterBar search={resultSearch} onSearch={setResultSearch} type={resultType} onType={setResultType} minAmount={minAmount} onMinAmount={setMinAmount} maxAmount={maxAmount} onMaxAmount={setMaxAmount} />

              {tab === "saved" && (
                <>
                <div className="flex flex-wrap items-center gap-2 border-b bg-card px-3 py-2.5">
                  <button
                    type="button"
                    onClick={() => setSelectedSavedIds(new Set(visibleSavedRows.map(row => (row.savedMatch as SavedMatch).id)))}
                    className="rounded-lg border border-blue-200 bg-blue-50 px-3 py-1.5 text-xs text-blue-700 hover:bg-blue-100"
                  >
                    تحديد الكل
                  </button>
                  <button
                    type="button"
                    onClick={handleUnsaveSelected}
                    disabled={!selectedSavedIds.size}
                    className="flex items-center gap-1.5 rounded-lg border border-amber-200 bg-amber-50 px-3 py-1.5 text-xs text-amber-700 hover:bg-amber-100 disabled:cursor-not-allowed disabled:opacity-40"
                  >
                    <Link2Off className="h-3.5 w-3.5" /> إلغاء تأكيد المحدد ({selectedSavedIds.size})
                  </button>
                  <button
                    type="button"
                    onClick={handleUndoMatch}
                    disabled={!canUndoMatch}
                    title="تراجع عن آخر إجراء (Ctrl+Z)"
                    className="flex items-center gap-1.5 rounded-lg border border-blue-200 bg-blue-50 px-3 py-1.5 text-xs text-blue-700 hover:bg-blue-100 disabled:cursor-not-allowed disabled:opacity-40"
                  >
                    <RotateCcw className="h-3.5 w-3.5" /> تراجع عن آخر إجراء{undoMatchCount > 1 ? ` (${undoMatchCount})` : ""}
                  </button>
                </div>
                <div className="overflow-x-auto">
                <table className="w-full min-w-[1400px] text-sm">
                  <thead><tr className="bg-muted text-xs">
                    <th className="w-10 px-2 py-2.5 text-center">
                      <input
                        type="checkbox"
                        checked={visibleSavedRows.length > 0 && visibleSavedRows.every(row => selectedSavedIds.has((row.savedMatch as SavedMatch).id))}
                        onChange={event => setSelectedSavedIds(event.target.checked
                          ? new Set(visibleSavedRows.map(row => (row.savedMatch as SavedMatch).id))
                          : new Set())}
                        aria-label="تحديد المطابقات المؤكدة"
                      />
                    </th>
                    <th className="px-3 py-2.5 text-right">البيان (الكاشير)</th>
                    <th className="px-3 py-2.5 text-right">بيان البنك</th>
                    <th className="px-3 py-2.5 text-right">النوع</th>
                    <th className="px-3 py-2.5 text-right">المبلغ</th>
                    <th className="px-3 py-2.5 text-right">نوع المطابقة</th>
                    <th className="px-3 py-2.5 text-right">الحالة</th>
                    <th className="px-3 py-2.5 text-right">الثقة</th>
                    <th className="px-3 py-2.5 w-10"></th>
                  </tr></thead>
                  <tbody>
                    {visibleSavedRows.map((r: any, i: number) => {
                      const sm = r.savedMatch as SavedMatch;
                      const key = sm.id;
                      const isExpanded = expandedMatchKey === key;
                      const typeColor = sm.type === "مدفوع" ? "text-red-600" : "text-green-600";
                      const matchTypeLabel = sm.isManual ? "🖐️ يدوي" : "🤖 تلقائي";
                      const isFromHeld = !!r.cashier._fromHeld || !!r.bank._fromHeld;
                      const visaNum = extractVisaNumber(r.cashier.rawName);
                      const dateDiff = transactionDatesDiffer(r.bank.date, r.cashier.date);
                      const accountTypeDiff = accountTypesDiffer(r.bank.accountType, r.cashier.accountType);
                      return (
                        <React.Fragment key={i}>
                          <tr data-default-reference={`${defaultReferenceForRow(r.cashier)}|${defaultReferenceForRow(r.bank)}`} className={`border-t transition-colors ${isExpanded ? "bg-green-50" : dateDiff || accountTypeDiff ? "bg-amber-50/50 hover:bg-amber-50" : isFromHeld ? "bg-purple-50/50 hover:bg-purple-50" : "hover:bg-muted/20"}`}>
                            <td className="px-2 py-2.5 text-center">
                              <input
                                type="checkbox"
                                checked={selectedSavedIds.has(sm.id)}
                                onChange={event => setSelectedSavedIds(prev => {
                                  const next = new Set(prev);
                                  if (event.target.checked) next.add(sm.id); else next.delete(sm.id);
                                  return next;
                                })}
                                aria-label={`تحديد مطابقة ${r.cashier.name}`}
                              />
                            </td>
                            <td className="px-3 py-2.5 font-medium flex items-center gap-1.5">
                              <Shield className="w-3.5 h-3.5 text-green-600" />
                              {r.cashier.name}
                              <DefaultReferenceButton row={r.cashier} onOpen={handleOpenReference} />
                              {isFromHeld && <span className="text-[9px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 border border-purple-300">📦 من ملف سابق</span>}
                              {visaNum && <span className="text-[9px] px-1.5 py-0.5 rounded bg-teal-100 text-teal-700 border border-teal-300">فيزا {visaNum}</span>}
                            </td>
                            <td className="px-3 py-2.5 text-muted-foreground text-xs">
                              {r.bank.description}
                              <div className="mt-1"><DefaultReferenceButton row={r.bank} onOpen={handleOpenReference} /></div>
                              {sm.bankAccountType && <div className="text-[10px] text-muted-foreground">{sm.bankAccountType}</div>}
                            </td>
                            <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{sm.type}</td>
                            <td className="px-3 py-2.5 font-mono font-bold text-green-700">
                              {fmtNum(r.bank.rawAmount)}
                              {sm.isAmountDiff && <div className="text-[10px] text-orange-600 font-normal">كاشير: {fmtNum(r.cashier.matchAmount)}</div>}
                            </td>
                            <td className="px-3 py-2.5 text-xs">{matchTypeLabel}</td>
                            <td className="px-3 py-2.5"><ConfidenceBadge score={sm.matchScore} /></td>
                            <td className="px-3 py-2.5">
                              <div className="space-y-0.5">
                                <span className="text-green-600 bg-green-50 px-1.5 py-0.5 rounded text-xs border border-green-200 flex items-center gap-0.5 w-fit">
                                  <Check className="w-2.5 h-2.5"/>✓ محفوظ
                                </span>
                                {sm.isAmountDiff && <div className="text-[10px] text-amber-600">⚠️ اختلاف مبلغ: البنك هو المرجع</div>}
                                {sm.isPhoneDiff && <div className="text-[10px] text-amber-700">⚠️ الجوال في الفاتورة مختلف؛ تعديل الرقم إلى: {movementPhones(r.cashier, r.bank).bankPhone || "غير موجود"}</div>}
                                {sm.isNameDiff && <div className="text-[10px] text-amber-600">⚠️ اختلاف اسم</div>}
                                {accountTypeDiff && <div className="text-[10px] text-orange-700 font-medium">🏦 {r.bank.accountType || "؟"} ↔ {r.cashier.accountType || "؟"}</div>}
                                {dateDiff && <div className="text-[10px] text-orange-700 font-medium">📅 اختلاف التاريخ: فاتورة {r.cashier.date} ↔ حوالة {r.bank.date}</div>}
                                {sm.note && <div className="text-[10px] text-gray-500">📝 {sm.note}</div>}
                              </div>
                            </td>
                            <td className="px-3 py-2.5 text-center">
                              <div className="flex items-center justify-center gap-1">
                                <button
                                  type="button"
                                  onClick={() => reviewItems.some(item => item.id === sm.id)
                                    ? handleCompleteReviewItem(sm.id)
                                    : handleAddReviewItem(sm, r.cashier, r.bank)}
                                  title={reviewItems.some(item => item.id === sm.id) ? "إزالة نسخة المراجعة" : "إضافة نسخة إلى قائمة المراجعة"}
                                  aria-label={reviewItems.some(item => item.id === sm.id) ? "إزالة من قائمة المراجعة" : "إضافة إلى قائمة المراجعة"}
                                  aria-pressed={reviewItems.some(item => item.id === sm.id)}
                                  className={`rounded p-1 transition-colors ${reviewItems.some(item => item.id === sm.id) ? "text-amber-600 hover:bg-amber-50" : "text-muted-foreground hover:bg-amber-50 hover:text-amber-600"}`}>
                                  <Star className={`h-4 w-4 ${reviewItems.some(item => item.id === sm.id) ? "fill-current" : ""}`} />
                                </button>
                                <button onClick={() => setExpandedMatchKey(isExpanded ? null : key)}
                                  className={`p-1 rounded transition-colors ${isExpanded ? "text-blue-600 bg-blue-100" : "text-muted-foreground hover:text-foreground hover:bg-muted"}`}>
                                  <span className="text-base leading-none font-bold">⋮</span>
                                </button>
                              </div>
                            </td>
                          </tr>
                          {isExpanded && (
                            <tr className="bg-green-50 border-t border-green-200">
                              <td colSpan={8} className="px-4 py-3">
                                <div className="flex items-center gap-3 flex-wrap">
                                  <span className="text-xs text-green-700 font-medium">{r.cashier.name} ↔ {r.bank.description.slice(0, 30)}...</span>
                                  <RowNoteButton note={rowNotes[transactionKey(r.cashier)]} onSave={note => handleSaveRowNote(r.cashier, note)} />
                                  <RowNoteButton note={rowNotes[transactionKey(r.bank)]} onSave={note => handleSaveRowNote(r.bank, note)} />
                                  <button onClick={() => handleUnsaveMatch(sm)}
                                    className="flex items-center gap-1.5 px-3 py-1.5 bg-amber-50 text-amber-600 border border-amber-200 rounded-lg text-xs hover:bg-amber-100 transition-colors">
                                    <Link2Off className="w-3.5 h-3.5"/>إلغاء الحفظ
                                  </button>
                                  <button onClick={() => handleRejectMatch(r.cashier, r.bank)}
                                    className="flex items-center gap-1.5 px-3 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs hover:bg-red-100 transition-colors">
                                    <X className="w-3.5 h-3.5"/>رفض
                                  </button>
                                  {visaNum && (
                                    <button onClick={() => handleMoveToVisa(r.cashier, "saved")}
                                      className="flex items-center gap-1.5 px-3 py-1.5 bg-teal-600 text-white rounded-lg text-xs hover:bg-teal-700 transition-colors">
                                      <CreditCard className="w-3.5 h-3.5"/>نقل للفيزا
                                    </button>
                                  )}
                                  <span className="text-xs text-muted-foreground">{sm.type}: {fmtNum(r.bank.rawAmount)}</span>
                                  <button onClick={() => setExpandedMatchKey(null)} className="mr-auto p-1 text-muted-foreground hover:text-foreground"><X className="w-3.5 h-3.5"/></button>
                                </div>
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                    {!visibleSavedRows.length && <tr><td colSpan={8} className="py-10 text-center text-muted-foreground text-xs">لا توجد مطابقات مؤكدة تطابق الفلاتر</td></tr>}
                  </tbody>
                </table>
                </div>
                </>
              )}

              {tab === "review" && (
                <div className="p-4">
                  <div className="mb-3 flex flex-wrap items-center justify-between gap-2 rounded-xl border border-amber-200 bg-amber-50/70 p-3">
                    <div>
                      <h2 className="text-sm font-bold text-amber-900">نسخ المطابقات للمراجعة</h2>
                      <p className="mt-1 text-xs text-amber-800">هذه نسخ مستقلة للمراجعة فقط؛ حذفها أو إنهاء مراجعتها لا يغيّر المطابقات المؤكدة أو الحركات الأصلية.</p>
                    </div>
                    <span className="text-xs font-semibold text-amber-900">{reviewItems.length} عنصر</span>
                  </div>
                  <div className="overflow-x-auto rounded-xl border">
                    <table className="w-full min-w-[900px] text-sm">
                      <thead><tr className="bg-muted text-xs">
                        <th className="px-3 py-2.5 text-right">الفاتورة</th>
                        <th className="px-3 py-2.5 text-right">الحوالة</th>
                        <th className="px-3 py-2.5 text-right">المبلغ</th>
                        <th className="px-3 py-2.5 text-right">الحساب</th>
                        <th className="px-3 py-2.5 text-right">تاريخ الإضافة للمراجعة</th>
                        <th className="px-3 py-2.5 text-center">إجراء</th>
                      </tr></thead>
                      <tbody>
                        {reviewItems.filter(item => {
                          const search = resultSearch.trim().toLowerCase();
                          const amount = item.bank?.rawAmount ?? item.cashier?.matchAmount ?? 0;
                          const text = `${item.title ?? ""} ${item.detail ?? ""} ${item.cashier?.name ?? ""} ${item.bank?.description ?? ""} ${item.cashier?.ref ?? ""} ${item.bank?.ref ?? ""} ${item.cashier ? defaultReferenceForRow(item.cashier) : ""} ${item.bank ? defaultReferenceForRow(item.bank) : ""}`.toLowerCase();
                          return (!search || text.includes(search))
                            && (resultType === "all" || item.cashier?.type === resultType || item.bank?.type === resultType)
                            && (!minAmount || amount >= Number(minAmount))
                            && (!maxAmount || amount <= Number(maxAmount));
                        }).map(item => (
                          <tr key={item.id} data-default-reference={`${item.cashier ? defaultReferenceForRow(item.cashier) : ""}|${item.bank ? defaultReferenceForRow(item.bank) : ""}`} className="border-t">
                            <td className="px-3 py-2.5 font-medium">{item.cashier?.name ?? item.title ?? "—"}{item.cashier && <><DefaultReferenceButton row={item.cashier} onOpen={handleOpenReference} /><div className="text-[10px] text-muted-foreground">{item.cashier.date || "بلا تاريخ"}</div></>}</td>
                            <td className="px-3 py-2.5">{item.bank?.description ?? item.detail ?? "—"}{item.bank && <><DefaultReferenceButton row={item.bank} onOpen={handleOpenReference} /><div className="text-[10px] text-muted-foreground">{item.bank.date || "بلا تاريخ"}</div></>}</td>
                            <td className="px-3 py-2.5 font-mono font-semibold">{fmtNum(item.bank?.rawAmount ?? item.cashier?.matchAmount ?? 0)}</td>
                            <td className="px-3 py-2.5 text-xs">{item.bank?.accountType || item.cashier?.accountType || "—"}</td>
                            <td className="px-3 py-2.5 text-xs text-muted-foreground">{item.addedAt}</td>
                            <td className="px-3 py-2.5 text-center">
                              <button type="button" onClick={() => handleCompleteReviewItem(item.id)} className="inline-flex items-center gap-1 rounded-lg border border-emerald-200 bg-emerald-50 px-2.5 py-1.5 text-xs text-emerald-800 hover:bg-emerald-100">
                                <Check className="h-3.5 w-3.5" />تمت المراجعة / حذف
                              </button>
                            </td>
                          </tr>
                        ))}
                        {!reviewItems.length && <tr><td colSpan={6} className="py-10 text-center text-xs text-muted-foreground">لا توجد عناصر في قائمة المراجعة. اضغط النجمة بجانب أي حركة أو مطابقة لإضافة نسخة مستقلة هنا.</td></tr>}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {tab === "pending" && (
                <div className="flex gap-4">
                <div className="flex-1 min-w-0 overflow-hidden">
                  {pendingRows.length > 0 && (
                    <div className="flex items-center gap-3 px-4 py-2.5 bg-indigo-50/60 border-b border-indigo-100 flex-wrap">
                      <span className="text-xs text-indigo-800">إعادة المطابقة بتفحص المنتظرة، تفك ربط الضعيفة وتعيد فرزها</span>
                      <button onClick={handleRetryMatch}
                        className="mr-auto flex items-center gap-1.5 px-3 py-1.5 bg-indigo-600 text-white rounded-lg text-xs font-medium hover:bg-indigo-700 transition-colors">
                        🔄 إعادة المطابقة
                      </button>
                    </div>
                  )}
                  {pendingRows.length > 0 && (
                    <>
                    <div className="flex items-center gap-3 px-4 py-2.5 bg-blue-50/60 border-b border-blue-100 flex-wrap">
                      <label className="flex items-center gap-1.5 text-xs font-medium text-blue-800 cursor-pointer">
                        <input type="checkbox" checked={allPendingSelected} onChange={togglePendingSelectAll} className="rounded"/>
                        تحديد الكل
                      </label>
                      <button type="button" onClick={toggleCurrentPagePendingSelection} disabled={!currentPagePendingKeys.length}
                        aria-pressed={allCurrentPagePendingSelected}
                        className="rounded-md border border-blue-200 bg-white px-2.5 py-1 text-xs font-medium text-blue-800 hover:bg-blue-100 disabled:cursor-not-allowed disabled:opacity-40">
                        {allCurrentPagePendingSelected ? `إلغاء تحديد الـ ${currentPagePendingKeys.length} في الصفحة` : `تحديد الـ ${currentPagePendingKeys.length} في الصفحة`}
                      </button>
                      <span className="text-xs text-muted-foreground">{selectedPendingKeys.size} محدد من {pendingRows.length}</span>
                      <div className="mr-auto flex items-center gap-2 flex-wrap">
                        <button onClick={handleMoveSelectedPendingToManual} disabled={!selectedPendingKeys.size}
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-amber-50 text-amber-700 border border-amber-200 rounded-lg text-xs font-medium hover:bg-amber-100 disabled:opacity-40 transition-colors">
                          نقل المحدد لليدوي ({selectedPendingKeys.size})
                        </button>
                        <button onClick={handleMoveAllPendingToManual}
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-amber-50 text-amber-700 border border-amber-200 rounded-lg text-xs font-medium hover:bg-amber-100 transition-colors">
                          نقل الكل لليدوي ({pendingRows.length})
                        </button>
                        <button onClick={handleSaveSelectedPending} disabled={!selectedPendingKeys.size}
                          className="flex items-center gap-1.5 px-3 py-1.5 bg-green-600 text-white rounded-lg text-xs font-medium hover:bg-green-700 disabled:opacity-40 transition-colors">
                          <Save className="w-3.5 h-3.5"/>حفظ المحدد ({selectedPendingKeys.size})
                        </button>
                      </div>
                    </div>
                    <div className="px-4 py-1.5 bg-indigo-50/50 border-b border-indigo-100 text-[11px] text-indigo-700">
                      💡 من لوحة المفاتيح: ↑↓ أو N للتنقل وتخطي الحركة، B أو Enter لحفظ المطابقة المحددة، وR لرفضها. الحركة المحددة بإطار أزرق. لاستبدال عنصر اسحب من اللوحة وأفلته على خانته.
                    </div>
                    </>
                  )}
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 w-8"></th>
                      <th className="px-3 py-2.5 text-right">البيان (الكاشير)</th>
                      <th className="px-3 py-2.5 text-right">بيان البنك</th>
                      <th className="px-3 py-2.5 text-right">النوع</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                      <th className="px-3 py-2.5 text-right">نوع المطابقة</th>ا
                      <th className="px-3 py-2.5 text-right">الحالة</th>
                      <th className="px-3 py-2.5 w-10"></th>
                    </tr></thead>
                    <tbody>
                      {pagedPendingRows.map((r: any, i: number) => {
                        const key = transactionPairKey(r.cashier, r.bank);
                        const isExpanded = expandedMatchKey === key;
                        const isSelected = selectedPendingKeys.has(key);
                        const isDropTarget = dropTargetKey === key;
                        const isKeyboardFocused = i === keyboardPendingIndex;
                        const typeColor = r.cashier.type === "مدفوع" ? "text-red-600" : "text-green-600";
                        const matchTypeLabel = r.matchType === "phone"
                                              ? (r.isExactAmount
                                                ? "الجوال والمبلغ متطابقان"
                                                : `الجوال متطابق · فرق مبلغ ${fmtNum(r.amountDiff ?? 0)}`)
                                              : r.matchType === "firstSecond" ? "👤 الأول+الثاني" :
                                              r.matchType === "fourthName" ? "👤 الرابع" :
                                              r.matchType === "exact" ? "تطابق تام" :
                                              r.matchType === "bundle" ? "تجميع" :
                                              r.matchType === "typo" ? "تقريبي" : "غير محدد";
                        const isFromHeld = !!r.cashier._fromHeld || !!r.bank._fromHeld;
                        const visaNum = extractVisaNumber(r.cashier.rawName);
                        return (
                          <React.Fragment key={i}>
                            <tr ref={row => { pendingRowRefs.current[key] = row; }} data-default-reference={`${defaultReferenceForRow(r.cashier)}|${defaultReferenceForRow(r.bank)}`} className={`border-t transition-colors ${isKeyboardFocused ? "outline outline-2 outline-blue-500 outline-offset-[-2px]" : ""} ${isDropTarget ? (drawerMode === "bank" ? "bg-blue-50 ring-2 ring-blue-400 ring-inset" : "bg-emerald-50 ring-2 ring-emerald-400 ring-inset") : isSelected ? "bg-blue-50/70" : isExpanded ? "bg-blue-50" : isFromHeld ? "bg-purple-50/50 hover:bg-purple-50" : r.isApprox ? "bg-amber-50/30 hover:bg-amber-50/60" : "hover:bg-muted/20"}`}>
                              <td className="px-3 py-2.5">
                                <div
                                  ref={element => { pendingRowCheckboxRefs.current[key] = element; }}
                                  role="checkbox"
                                  aria-checked={isSelected}
                                  aria-label={`تحديد مطابقة ${r.cashier.name}`}
                                  tabIndex={0}
                                  data-pending-row-checkbox={key}
                                  onFocus={() => setKeyboardPendingIndex(i)}
                                  onClick={() => { setKeyboardPendingIndex(i); togglePendingSelect(key); }}
                                  className={`w-4 h-4 rounded border-2 flex items-center justify-center cursor-pointer transition-colors ${isSelected?"bg-blue-600 border-blue-600":"border-border"}`}>
                                  {isSelected && <Check className="w-2.5 h-2.5 text-white"/>}
                                </div>
                              </td>
                              <td
                                className={`px-3 py-2.5 font-medium transition-colors ${isDropTarget && drawerMode === "cashier" ? "bg-emerald-100" : ""}`}
                                onDragOver={(e) => { if (draggedItemId !== null && drawerMode === "cashier") { e.preventDefault(); setDropTargetKey(key); } }}
                                onDragLeave={() => { if (dropTargetKey === key) setDropTargetKey(null); }}
                                onDrop={(e) => { e.preventDefault(); if (draggedItemId !== null && drawerMode === "cashier") { handleDropCashier(r.bank, r.cashier, draggedItemId); } setDraggedItemId(null); }}
                              >
                                <div className="flex items-center gap-1.5">
                                  <span className="flex-1">
                                    {r.cashier.name}
                                    <div className="mt-1"><DefaultReferenceButton row={r.cashier} onOpen={handleOpenReference} /></div>
                                    {r.cashier.accountType && <div className="text-[10px] text-muted-foreground font-normal">{r.cashier.accountType}</div>}
                                  </span>
                                  {isFromHeld && <span className="text-[9px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 border border-purple-300 shrink-0">📦 من ملف سابق</span>}
                                  {visaNum && <span className="text-[9px] px-1.5 py-0.5 rounded bg-teal-100 text-teal-700 border border-teal-300 shrink-0">فيزا {visaNum}</span>}
                                  <button
                                    onClick={() => openDrawer(r.cashier, r.bank, "cashier")}
                                    title="افتح قائمة فواتير الكاشير للبحث والسحب"
                                    className="px-1.5 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded text-[10px] font-medium hover:bg-emerald-100 transition-colors flex items-center gap-0.5 shrink-0"
                                  >
                                    <Search className="w-2.5 h-2.5"/>بحث وسحب
                                  </button>
                                </div>
                              </td>
                              <td
                                className={`px-3 py-2.5 text-muted-foreground text-xs transition-colors ${isDropTarget && drawerMode === "bank" ? "bg-blue-100" : ""}`}
                                onDragOver={(e) => { if (draggedItemId !== null && drawerMode === "bank") { e.preventDefault(); setDropTargetKey(key); } }}
                                onDragLeave={() => { if (dropTargetKey === key) setDropTargetKey(null); }}
                                onDrop={(e) => { e.preventDefault(); if (draggedItemId !== null && drawerMode === "bank") { handleDropBank(r.cashier, r.bank, draggedItemId); } setDraggedItemId(null); }}
                              >
                                <div className="flex items-center gap-1.5">
                                  <span className="flex-1">{r.bank.description}<div className="mt-1"><DefaultReferenceButton row={r.bank} onOpen={handleOpenReference} /></div></span>
                                  {r.bank.accountType && <span className="text-[10px] text-muted-foreground">{r.bank.accountType}</span>}
                                  <button
                                    onClick={() => openDrawer(r.cashier, r.bank, "bank")}
                                    title="افتح قائمة الحوالات للبحث والسحب"
                                    className="px-1.5 py-0.5 bg-indigo-50 text-indigo-700 border border-indigo-200 rounded text-[10px] font-medium hover:bg-indigo-100 transition-colors flex items-center gap-0.5 shrink-0"
                                  >
                                    <Search className="w-2.5 h-2.5"/>بحث وسحب
                                  </button>
                                </div>
                              </td>
                              <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{r.cashier.type}</td>
                              <td className="px-3 py-2.5 font-mono font-bold text-blue-700">{fmtNum(r.bank.rawAmount)}</td>
                              <td className="px-3 py-2.5 text-xs">{matchTypeLabel}</td>
                              <td className="px-3 py-2.5">
                                <div className="flex flex-col gap-1">
                                  {r.isBundled ? (
                                    <span className="text-blue-600 bg-blue-50 px-1.5 py-0.5 rounded text-xs border border-blue-200">تجميع</span>
                                  ) : r.isApprox ? (
                                    <span className="text-amber-700 bg-amber-50 border border-amber-200 px-1.5 py-0.5 rounded text-xs flex items-center gap-0.5 w-fit">
                                      <AlertTriangle className="w-2.5 h-2.5"/>اسم تقريبي
                                    </span>
                                  ) : (
                                    <span className="text-blue-600 bg-blue-50 px-1.5 py-0.5 rounded text-xs">آلي ✓</span>
                                  )}
                                  <ConfidenceBadge score={Math.round(Math.min(1, nameSim(r.cashier.name, r.bank.description) + (r.isApprox ? 0.25 : 0.45)) * 100)} />
                                  {r.accountTypeDiff && (
                                    <span className="text-orange-700 bg-orange-50 border border-orange-200 px-1.5 py-0.5 rounded text-xs flex items-center gap-0.5 w-fit">
                                      <AlertTriangle className="w-2.5 h-2.5"/>{r.bank.accountType || "؟"} ↔ {r.cashier.accountType || "؟"}
                                    </span>
                                  )}
                                </div>
                              </td>
                              <td className="px-3 py-2.5 text-center">
                                <button onClick={() => setExpandedMatchKey(isExpanded ? null : key)}
                                  className={`p-1 rounded transition-colors ${isExpanded ? "text-blue-600 bg-blue-100" : "text-muted-foreground hover:text-foreground hover:bg-muted"}`}>
                                  <span className="text-base leading-none font-bold">⋮</span>
                                </button>
                              </td>
                            </tr>
                            {isExpanded && (
                              <tr className="bg-blue-50 border-t border-blue-200">
                                <td colSpan={8} className="px-4 py-3">
                                  <div className="flex items-center gap-3 flex-wrap">
                                    <span className="text-xs text-blue-700 font-medium">{r.cashier.name} ↔ {r.bank.description.slice(0, 30)}...</span>
                                    <RowNoteButton note={rowNotes[transactionKey(r.cashier)]} onSave={note => handleSaveRowNote(r.cashier, note)} />
                                    <RowNoteButton note={rowNotes[transactionKey(r.bank)]} onSave={note => handleSaveRowNote(r.bank, note)} />
                                    <button type="button" onClick={() => handleToggleReviewSnapshot(r.cashier, r.bank, "مطابقة منتظرة")}
                                      aria-pressed={reviewItems.some(item => item.id === `review-pair:${key}`)}
                                      className={`flex items-center gap-1.5 rounded-lg border px-3 py-1.5 text-xs transition-colors ${reviewItems.some(item => item.id === `review-pair:${key}`) ? "border-amber-300 bg-amber-100 text-amber-800" : "border-amber-200 bg-amber-50 text-amber-700 hover:bg-amber-100"}`}>
                                      <Star className={`h-3.5 w-3.5 ${reviewItems.some(item => item.id === `review-pair:${key}`) ? "fill-current" : ""}`}/>مراجعة
                                    </button>
                                    {isMahmoudWalletCashier(r.cashier) && <button type="button" onClick={() => handleMoveToMahmoudWallet(r.cashier)}
                                      className="rounded-lg border border-violet-200 bg-violet-50 px-3 py-1.5 text-xs text-violet-700 hover:bg-violet-100">ترحيل لمحفظة شخصية</button>}
                                    <button onClick={() => handleSaveMatch(r.cashier, r.bank)}
                                      className="flex items-center gap-1.5 px-3 py-1.5 bg-green-600 text-white rounded-lg text-xs hover:bg-green-700 transition-colors">
                                      <Save className="w-3.5 h-3.5"/>حفظ المطابقة
                                    </button>
                                    <button onClick={() => handleRejectMatch(r.cashier, r.bank)}
                                      className="flex items-center gap-1.5 px-3 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs hover:bg-red-100 transition-colors">
                                      <X className="w-3.5 h-3.5"/>رفض
                                    </button>
                                    {visaNum && (
                                      <button onClick={() => handleMoveToVisa(r.cashier, "pending")}
                                        className="flex items-center gap-1.5 px-3 py-1.5 bg-teal-600 text-white rounded-lg text-xs hover:bg-teal-700 transition-colors">
                                        <CreditCard className="w-3.5 h-3.5"/>نقل للفيزا
                                      </button>
                                    )}
                                    <button onClick={() => handleHoldCashier(r.cashier)}
                                      title="علّق فاتورة الكاشير لمطابقتها لاحقاً"
                                      className="flex items-center gap-1.5 rounded-lg border border-purple-200 bg-purple-50 px-3 py-1.5 text-xs text-purple-700 hover:bg-purple-100">
                                      <CreditCard className="h-3.5 w-3.5"/>علّق كاشير
                                    </button>
                                    <button onClick={() => handleHoldBank(r.bank)}
                                      title="علّق حوالة البنك لمطابقتها لاحقاً"
                                      className="flex items-center gap-1.5 rounded-lg border border-purple-200 bg-purple-50 px-3 py-1.5 text-xs text-purple-700 hover:bg-purple-100">
                                      <CreditCard className="h-3.5 w-3.5"/>علّق بنك
                                    </button>
                                    <span className="text-xs text-muted-foreground">{r.cashier.type}: {fmtNum(r.bank.rawAmount)}</span>
                                    <button onClick={() => setExpandedMatchKey(null)} className="mr-auto p-1 text-muted-foreground hover:text-foreground"><X className="w-3.5 h-3.5"/></button>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      })}
                      {!pagedPendingRows.length && <tr><td colSpan={8} className="py-10 text-center text-muted-foreground text-xs">لا توجد مطابقات منتظرة تطابق الفلاتر</td></tr>}
                    </tbody>
                  </table>
                  {visiblePendingRows.length > PENDING_PAGE_SIZE && (
                    <div className="flex items-center justify-center gap-3 px-4 py-3 border-t border-border bg-muted/30 text-xs">
                      <button onClick={() => setPendingPage(p => Math.max(1, p - 1))} disabled={pendingPage <= 1}
                        className="px-3 py-1.5 rounded-lg border border-border bg-card disabled:opacity-40 hover:bg-muted/50 transition-colors flex items-center gap-1">
                        <ChevronRight className="w-3.5 h-3.5"/>السابق
                      </button>
                      <span className="text-muted-foreground">
                        صفحة {pendingPage} من {Math.max(1, Math.ceil(visiblePendingRows.length / PENDING_PAGE_SIZE))} · عرض {pagedPendingRows.length} من {visiblePendingRows.length}
                      </span>
                      <button onClick={() => setPendingPage(p => Math.min(Math.max(1, Math.ceil(visiblePendingRows.length / PENDING_PAGE_SIZE)), p + 1))}
                        disabled={pendingPage >= Math.ceil(visiblePendingRows.length / PENDING_PAGE_SIZE)}
                        className="px-3 py-1.5 rounded-lg border border-border bg-card disabled:opacity-40 hover:bg-muted/50 transition-colors flex items-center gap-1">
                        التالي<ChevronLeft className="w-3.5 h-3.5"/>
                      </button>
                    </div>
                  )}
                </div>
                {drawerOpen && drawerCashier && drawerOldBank && (
                  <MatchDrawer
                    mode={drawerMode}
                    cashier={drawerCashier}
                    oldBank={drawerOldBank}
                    nameSearch={drawerNameSearch}
                    amountFrom={drawerAmountFrom}
                    amountTo={drawerAmountTo}
                    banks={drawerBanks}
                    cashiers={drawerCashiers}
                    draggedItemId={draggedItemId}
                    onNameSearch={setDrawerNameSearch}
                    onAmountFrom={setDrawerAmountFrom}
                    onAmountTo={setDrawerAmountTo}
                    onDragStart={setDraggedItemId}
                    onDragEnd={() => { setDraggedItemId(null); setDropTargetKey(null); }}
                    onReplaceBank={(b) => handleReplaceBank(drawerCashier, drawerOldBank, b)}
                    onReplaceCashier={(c) => handleReplaceCashier(drawerOldBank, drawerCashier, c)}
                    onClose={closeDrawer}
                  />
                )}
              </div>
              )}

              {tab === "manual" && (
                <div className="space-y-3 p-4">
                  <p className="text-center text-xs text-muted-foreground">المطابقات اليدوية تذهب مباشرة إلى "المؤكدة". الحركات في المجموعات اليدوية التالية محفوظة هنا.</p>
                  {manualGroups.map(group => (
                    <article key={group.id} className="rounded-xl border p-3">
                      <h3 className="mb-2 text-xs font-semibold">مجموعة يدوية · {group.id}</h3>
                      <div className="grid gap-2 md:grid-cols-2">
                        {group.cashiers.map(row => <div key={transactionKey(row)} data-default-reference={defaultReferenceForRow(row)} className="flex items-center justify-between gap-2 rounded bg-muted/40 p-2 text-xs">
                          <span>{row.name} · {fmtNum(row.matchAmount)} ₪</span><DefaultReferenceButton row={row} onOpen={handleOpenReference} />
                        </div>)}
                        {group.banks.map(row => <div key={transactionKey(row)} data-default-reference={defaultReferenceForRow(row)} className="flex items-center justify-between gap-2 rounded bg-muted/40 p-2 text-xs">
                          <span>{row.description} · {fmtNum(row.rawAmount)} ₪</span><DefaultReferenceButton row={row} onOpen={handleOpenReference} />
                        </div>)}
                      </div>
                    </article>
                  ))}
                </div>
              )}

              {tab === "visa" && (
                <div>
                  <div className="px-4 py-3 bg-teal-50/60 border-b border-teal-100">
                    <p className="text-xs text-teal-800">
                      هذه الخانة خاصة بالفواتير التي تحتوي على رقم فيزا (4 أرقام بعد / أو -).
                      تقدر تنقل أي فاتورة من "المنتظر" أو "غير المتطابق" أو "المؤكد" لهونا بضغطة زر "نقل للفيزا".
                    </p>
                  </div>
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 text-right">الاسم</th>
                      <th className="px-3 py-2.5 text-right">رقم الفيزا</th>
                      <th className="px-3 py-2.5 text-right">النوع</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                      <th className="px-3 py-2.5 text-right">التاريخ</th>
                      <th className="px-3 py-2.5 text-right">ملاحظة</th>
                      <th className="px-3 py-2.5 text-right">إجراء</th>
                    </tr></thead>
                    <tbody>
                      {visaItems.map(v => {
                        const typeColor = v.type === "مدفوع" ? "text-red-600" : "text-green-600";
                        return (
                          <tr key={v.id} data-default-reference={v.defaultRef} className="border-t hover:bg-muted/20 transition-colors">
                            <td className="px-3 py-2.5 font-medium">
                              {v.name}
                              {(() => {
                                const row = cashierSourceRows.find(candidate => candidate.id === v.cashierId && (candidate._fileSessionId ?? 0) === (v.cashierSessionId ?? 0));
                                return row ? <div className="mt-1"><DefaultReferenceButton row={row} onOpen={handleOpenReference} /></div> : null;
                              })()}
                              <div className="text-[10px] text-muted-foreground font-normal">{v.rawName}</div>
                            </td>
                            <td className="px-3 py-2.5">
                              <span className="font-mono font-bold text-teal-700 bg-teal-50 px-2 py-0.5 rounded border border-teal-200">{v.visaNumber}</span>
                            </td>
                            <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{v.type}</td>
                            <td className="px-3 py-2.5 font-mono font-bold text-teal-700">{fmtNum(v.amount)}</td>
                            <td className="px-3 py-2.5 text-xs text-muted-foreground">{v.date}</td>
                            <td className="px-3 py-2.5 text-xs text-gray-500">
                              {v.note || "—"}
                              <div className="text-[10px] text-muted-foreground">{v.movedAt}</div>
                            </td>
                            <td className="px-3 py-2.5">
                              <button onClick={() => handleRemoveVisaItem(v.id)}
                                className="flex items-center gap-1 px-2.5 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs hover:bg-red-100 transition-colors">
                                <X className="w-3.5 h-3.5"/>إزالة من الفيزا
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                      {!visaItems.length && <tr><td colSpan={7} className="py-10 text-center text-muted-foreground text-xs">
                        لا توجد عناصر فيزا حالياً. نقل أي فاتورة تحتوي على رقم (4 أرقام بعد / أو -) من تبويب "منتظر" أو "كاشير غير متطابق" بضغط زر "نقل للفيزا".
                      </td></tr>}
                    </tbody>
                  </table>
                </div>
              )}

              {tab === "jawwalPay" && (
                <div>
                  <div className="border-b border-cyan-100 bg-cyan-50/60 px-4 py-3">
                    <p className="text-xs text-cyan-800">هذه الخانة خاصة بكل الفواتير التي تحتوي على «جوال» أو «جوال بي». لا تدخل هذه الفواتير في المطابقة.</p>
                  </div>
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 text-right">الاسم</th>
                      <th className="px-3 py-2.5 text-right">البيان الأصلي</th>
                      <th className="px-3 py-2.5 text-right">النوع</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                      <th className="px-3 py-2.5 text-right">نوع الحساب</th>
                      <th className="px-3 py-2.5 text-right">التاريخ</th>
                      <th className="px-3 py-2.5 text-right">إجراء</th>
                    </tr></thead>
                    <tbody>
                      {results.filter(r => r.type === "jawwalPay").map((r: any) => (
                        <tr key={r.cashier.id} data-default-reference={defaultReferenceForRow(r.cashier)} className="border-t">
                          <td className="px-3 py-2.5 font-medium">{r.cashier.name} <DefaultReferenceButton row={r.cashier} onOpen={handleOpenReference} /></td>
                          <td className="px-3 py-2.5">{r.cashier.rawName}</td>
                          <td className="px-3 py-2.5">{r.cashier.type}</td>
                          <td className="px-3 py-2.5 font-mono font-bold text-cyan-700">{fmtNum(r.cashier.amount)}</td>
                          <td className="px-3 py-2.5">{r.cashier.accountType || "—"}</td>
                          <td className="px-3 py-2.5">{r.cashier.date || "—"}</td>
                          <td className="px-3 py-2.5">
                            <div className="flex flex-wrap gap-1.5">
                              <button type="button" onClick={() => handleToggleReviewSnapshot(r.cashier, undefined, "فاتورة جوال بي")}
                                aria-label={`إضافة فاتورة ${r.cashier.name} إلى المراجعة`} title="إضافة نسخة إلى المراجعة"
                                className="rounded-lg border border-amber-200 p-1.5 text-amber-700 hover:bg-amber-50"><Star className="h-4 w-4" /></button>
                              {isMahmoudWalletCashier(r.cashier) && (
                                <button type="button" onClick={() => handleMoveToMahmoudWallet(r.cashier)}
                                  title="ترحيل الفاتورة إلى محفظة شخصية"
                                  className="flex items-center gap-1 px-2.5 py-1.5 bg-violet-50 text-violet-700 border border-violet-200 rounded-lg text-xs hover:bg-violet-100 transition-colors">
                                  ترحيل لمحفظة شخصية
                                </button>
                              )}
                              <button onClick={() => handleRejectSpecialCashier(r.cashier)}
                                className="flex items-center gap-1 px-2.5 py-1.5 bg-amber-50 text-amber-700 border border-amber-200 rounded-lg text-xs hover:bg-amber-100 transition-colors">
                                <X className="w-3.5 h-3.5"/>رفض التصنيف وإرجاع للمنتظر
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                      {!results.some(r => r.type === "jawwalPay") && <tr><td colSpan={7} className="py-10 text-center text-muted-foreground text-xs">لا توجد فواتير جوال بي.</td></tr>}
                    </tbody>
                  </table>
                </div>
              )}

              {tab === "mahmoudWallet" && (
                <div>
                  <div className="border-b border-indigo-100 bg-indigo-50/60 px-4 py-3">
                    <p className="text-xs text-indigo-800">تعرض هذه الخانة الفواتير التي رحّلتها يدويًا إلى محفظة شخصية. لا تدخل هذه الفواتير في المطابقة، ويمكن إعادتها إليها بزر رفض التصنيف.</p>
                  </div>
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 text-right">الاسم</th>
                      <th className="px-3 py-2.5 text-right">البيان الأصلي</th>
                      <th className="px-3 py-2.5 text-right">النوع</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                      <th className="px-3 py-2.5 text-right">نوع الحساب</th>
                      <th className="px-3 py-2.5 text-right">التاريخ</th>
                      <th className="px-3 py-2.5 text-right">إجراء</th>
                    </tr></thead>
                    <tbody>
                      {results.filter(r => r.type === "mahmoudWallet").map((r: any) => (
                        <tr key={r.cashier.id} data-default-reference={defaultReferenceForRow(r.cashier)} className="border-t">
                          <td className="px-3 py-2.5 font-medium">{r.cashier.name} <DefaultReferenceButton row={r.cashier} onOpen={handleOpenReference} /></td>
                          <td className="px-3 py-2.5">{r.cashier.rawName}</td>
                          <td className="px-3 py-2.5">{r.cashier.type}</td>
                          <td className="px-3 py-2.5 font-mono font-bold text-indigo-700">{fmtNum(r.cashier.amount)}</td>
                          <td className="px-3 py-2.5">{r.cashier.accountType || "—"}</td>
                          <td className="px-3 py-2.5">{r.cashier.date || "—"}</td>
                          <td className="px-3 py-2.5">
                            <div className="flex flex-wrap gap-1.5">
                              <button type="button" onClick={() => handleToggleReviewSnapshot(r.cashier, undefined, "فاتورة محفظة شخصية")}
                                aria-label={`إضافة فاتورة ${r.cashier.name} إلى المراجعة`} title="إضافة نسخة إلى المراجعة"
                                className="rounded-lg border border-amber-200 p-1.5 text-amber-700 hover:bg-amber-50"><Star className="h-4 w-4" /></button>
                              <button onClick={() => handleRejectSpecialCashier(r.cashier)}
                                className="flex items-center gap-1 px-2.5 py-1.5 bg-amber-50 text-amber-700 border border-amber-200 rounded-lg text-xs hover:bg-amber-100 transition-colors">
                                <X className="w-3.5 h-3.5"/>رفض التصنيف وإرجاع للمنتظر
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                      {!results.some(r => r.type === "mahmoudWallet") && <tr><td colSpan={7} className="py-10 text-center text-muted-foreground text-xs">لا توجد فواتير محفظة شخصية.</td></tr>}
                    </tbody>
                  </table>
                </div>
              )}

              {tab === "held" && (
                <>
                {!!merchantHeldBankItems.length && <div className="flex flex-wrap items-center justify-between gap-2 border-b bg-amber-50/70 px-4 py-3">
                  <p className="text-xs text-amber-900">{merchantHeldBankItems.length} حركة مشتريات عمولة تجار جاهزة للترحيل إلى C ثم D.</p>
                  <button type="button" onClick={() => handleTransferHeldMerchantToC(merchantHeldBankItems)}
                    className="rounded-lg bg-indigo-600 px-3 py-2 text-xs font-medium text-white hover:bg-indigo-700">
                    ترحيل مشتريات التجار إلى C وD ({merchantHeldBankItems.length})
                  </button>
                </div>}
                <table className="w-full text-sm">
                  <thead><tr className="bg-muted text-xs">
                    <th className="px-3 py-2.5 text-right">النوع</th>
                    <th className="px-3 py-2.5 text-right">البيان</th>
                    <th className="px-3 py-2.5 text-right">النوع (مدفوع/مستلم)</th>
                    <th className="px-3 py-2.5 text-right">المبلغ</th>
                    <th className="px-3 py-2.5 text-right">تاريخ التعليق</th>
                    <th className="px-3 py-2.5 text-right">ملاحظة</th>
                    <th className="px-3 py-2.5 text-right">إجراء</th>
                  </tr></thead>
                  <tbody>
                    {heldItems.map(h => {
                      const isBank = h.kind === "bank";
                      const label = isBank ? (h.data as BankRow).description : (h.data as CashierRow).name;
                      const amt = isBank ? (h.data as BankRow).rawAmount : (h.data as CashierRow).amount;
                      const type = h.data.type;
                      const typeColor = type === "مدفوع" ? "text-red-600" : "text-green-600";
                      return (
                        <tr key={h.id} data-default-reference={defaultReferenceForRow(h.data)} className="border-t hover:bg-muted/20 transition-colors">
                          <td className="px-3 py-2.5">
                            <span className={`px-1.5 py-0.5 rounded text-xs font-medium ${isBank ? "bg-blue-100 text-blue-800" : "bg-green-100 text-green-800"}`}>
                              {isBank ? "بنك" : "كاشير"}
                            </span>
                          </td>
                          <td className="px-3 py-2.5 font-medium">{label} <DefaultReferenceButton row={h.data} onOpen={handleOpenReference} /></td>
                          <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{type}</td>
                          <td className="px-3 py-2.5 font-mono font-semibold text-purple-700">{fmtNum(amt)}</td>
                          <td className="px-3 py-2.5 text-xs text-muted-foreground">{h.heldAt}</td>
                          <td className="px-3 py-2.5 text-xs text-gray-500">{h.note || "—"}</td>
                          <td className="px-3 py-2.5">
                            <div className="flex gap-1">
                              <button onClick={() => handleUnhold(h)}
                                className="flex items-center gap-1 px-2.5 py-1.5 bg-green-600 text-white rounded-lg text-xs hover:bg-green-700 transition-colors">
                                <Link2 className="w-3.5 h-3.5"/>استرجاع للمطابقة
                              </button>
                              <button type="button"
                                onClick={() => isBank
                                  ? handleToggleReviewSnapshot(undefined, h.data as BankRow, "حركة معلّقة")
                                  : handleToggleReviewSnapshot(h.data as CashierRow, undefined, "حركة معلّقة")}
                                title="إضافة نسخة إلى المراجعة"
                                aria-label={`إضافة ${label} إلى المراجعة`}
                                className="rounded-lg border border-amber-200 p-1.5 text-amber-700 hover:bg-amber-50">
                                <Star className="h-4 w-4" />
                              </button>
                              {isBank && isMerchantPurchaseDescription(label) && <button type="button"
                                onClick={() => handleTransferHeldMerchantToC([h])}
                                className="rounded-lg border border-indigo-200 bg-indigo-50 px-2 py-1.5 text-xs text-indigo-700 hover:bg-indigo-100">
                                مطابقة C وD
                              </button>}
                              <button onClick={() => handleDeleteHeld(h.id)}
                                className="p-1.5 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors">
                                <Trash2 className="w-3.5 h-3.5"/>
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                    {!heldItems.length && <tr><td colSpan={7} className="py-10 text-center text-muted-foreground text-xs">لا توجد عناصر معلّقة حالياً. علّق أي حوالة أو فاتورة من تبويب "منتظر" أو "كاشير/بنك غير متطابق" لمطابقتها لاحقاً.</td></tr>}
                  </tbody>
                </table>
                </>
              )}

              {tab === "nonReconciliationBank" && (
                <div>
                  <p className="border-b border-amber-100 bg-amber-50/70 px-4 py-3 text-xs text-amber-900">
                    هذه الحركات محفوظة للرجوع إليها، لكنها مستبعدة من المطابقة ومجاميع التقارير.
                  </p>
                  <div className="flex flex-wrap items-center gap-2 border-b bg-white px-4 py-3">
                    <select value={nonReconBankTypeFilter} onChange={event => {
                      setNonReconBankTypeFilter(event.target.value);
                      setSelectedNonReconBankKeys(new Set());
                    }} className="rounded-lg border bg-white px-3 py-2 text-xs">
                      <option value="all">كل التصنيفات</option>
                      <option value="merchant">مشتريات عمولة تجار</option>
                      <option value="other">باقي الحركات المستبعدة</option>
                    </select>
                    <button type="button" onClick={() => setSelectedNonReconBankKeys(new Set(visibleFilteredNonReconciliationBankRows.map(({ bank }) => transactionKey(bank))))}
                      disabled={!visibleFilteredNonReconciliationBankRows.length}
                      className="rounded-lg border border-slate-200 px-3 py-2 text-xs disabled:opacity-40">
                      تحديد الظاهر ({visibleFilteredNonReconciliationBankRows.length})
                    </button>
                    <button type="button" onClick={handleMoveSelectedNonReconBankToHeld}
                      disabled={!selectedNonReconBankKeys.size}
                      className="rounded-lg bg-amber-600 px-3 py-2 text-xs font-medium text-white disabled:opacity-40">
                      نقل المحدد إلى المعلّقات ({selectedNonReconBankKeys.size})
                    </button>
                  </div>
                  <table className="w-full text-sm">
                    <thead><tr className="bg-muted text-xs">
                      <th className="px-3 py-2.5 text-right">تحديد</th>
                      <th className="px-3 py-2.5 text-right">التاريخ</th>
                      <th className="px-3 py-2.5 text-right">البيان</th>
                      <th className="px-3 py-2.5 text-right">التصنيف</th>
                      <th className="px-3 py-2.5 text-right">نوع الحساب</th>
                      <th className="px-3 py-2.5 text-right">النوع</th>
                      <th className="px-3 py-2.5 text-right">مدين</th>
                      <th className="px-3 py-2.5 text-right">دائن</th>
                      <th className="px-3 py-2.5 text-right">المبلغ</th>
                    </tr></thead>
                    <tbody>
                      {visibleFilteredNonReconciliationBankRows.map(({ bank, reason }) => (
                        <tr key={`${bank._fileSessionId ?? bankFileSessionId}:${bank.id}`} data-default-reference={defaultReferenceForRow(bank)} className="border-t hover:bg-muted/20">
                          <td className="px-3 py-2.5"><input type="checkbox" checked={selectedNonReconBankKeys.has(transactionKey(bank))}
                            onChange={() => setSelectedNonReconBankKeys(previous => {
                              const next = new Set(previous);
                              const key = transactionKey(bank);
                              if (next.has(key)) next.delete(key); else next.add(key);
                              return next;
                            })} aria-label={`تحديد حركة ${bank.description}`} /></td>
                          <td className="px-3 py-2.5">{bank.date || "—"}</td>
                          <td className="px-3 py-2.5">{bank.description} <DefaultReferenceButton row={bank} onOpen={handleOpenReference} /></td>
                          <td className="px-3 py-2.5 text-amber-800">{reason}</td>
                          <td className="px-3 py-2.5">{bank.accountType || "—"}</td>
                          <td className={`px-3 py-2.5 text-xs font-semibold ${bank.type === "مدفوع" ? "text-red-600" : "text-green-600"}`}>{bank.type}</td>
                          <td className="px-3 py-2.5 font-mono">{fmtNum(bank.debit)}</td>
                          <td className="px-3 py-2.5 font-mono">{fmtNum(bank.credit)}</td>
                          <td className="px-3 py-2.5 font-mono font-semibold">{fmtNum(bank.rawAmount)}</td>
                        </tr>
                      ))}
                      {!visibleFilteredNonReconciliationBankRows.length && <tr><td colSpan={9} className="py-10 text-center text-muted-foreground text-xs">لا توجد حركات مصنّفة هنا.</td></tr>}
                    </tbody>
                  </table>
                </div>
              )}

              {tab === "uCashier" && (
                <table className="w-full text-sm">
                  <thead><tr className="bg-muted text-xs">
                    <th className="px-3 py-2.5 text-right">البيان</th>
                    <th className="px-3 py-2.5 text-right">النوع</th>
                    <th className="px-3 py-2.5 text-right">المبلغ</th>
                    <th className="px-3 py-2.5 text-right">نوع الحساب المسجل</th>
                    <th className="px-3 py-2.5 text-right">رقم المستخدم</th>
                    <th className="px-3 py-2.5 text-right">السبب</th>
                    <th className="px-3 py-2.5 text-right">اقتراح</th>
                    <th className="px-3 py-2.5 text-right">إجراء</th>
                  </tr></thead>
                  <tbody>
                    {visibleUCashierRows.map((r: any, i: number) => {
                      const isExpanded = expandedUnmatchedCashier === r.cashier.id;
                      const typeColor = r.cashier.type === "مدفوع" ? "text-red-600" : "text-green-600";
                      const availableBanks = unmatchedBankForSuggestions.filter(
                        b => !wasPairRejected(rejectedPairs, r.cashier, b) && !savedBankIds.has(transactionKey(b)) && b.type === r.cashier.type
                      );
                      const suggestions = isExpanded ? getSuggestions(r.cashier, availableBanks, rejectedPairs, savedKeys, claimedNames, suggestionOwnerMap, nameAliases) : [];
                      const isFromHeld = !!r.cashier._fromHeld;
                      const visaNum = extractVisaNumber(r.cashier.rawName);
                      return (
                        <React.Fragment key={i}>
                          <tr data-default-reference={defaultReferenceForRow(r.cashier)} className={`border-t transition-colors ${isExpanded ? "bg-blue-50" : isFromHeld ? "bg-purple-50/50 hover:bg-purple-50" : r.reason === "اختلاف في الاسم" ? "bg-amber-50/30 hover:bg-amber-50/60" : "hover:bg-muted/20"}`}>
                            <td className="px-3 py-2.5 font-medium">
                              {r.cashier.name}
                              <span className="mr-2"><DefaultReferenceButton row={r.cashier} onOpen={handleOpenReference} /></span>
                              {isFromHeld && <span className="mr-1.5 text-[9px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 border border-purple-300">📦 من ملف سابق</span>}
                              {visaNum && <span className="mr-1.5 text-[9px] px-1.5 py-0.5 rounded bg-teal-100 text-teal-700 border border-teal-300">فيزا {visaNum}</span>}
                            </td>
                            <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{r.cashier.type}</td>
                            <td className="px-3 py-2.5 font-mono font-semibold text-red-700">{fmtNum(r.cashier.amount)}</td>
                            <td className="px-3 py-2.5">{r.cashier.accountType || "—"}</td>
                            <td className="px-3 py-2.5 font-mono">{r.cashier.userId || "—"}</td>
                            <td className="px-3 py-2.5">
                              <span className={`px-1.5 py-0.5 rounded text-xs font-medium ${r.reason === "اختلاف في الاسم" ? "bg-amber-100 text-amber-800" : "bg-red-100 text-red-800"}`}>{r.reason}</span>
                            </td>
                            <td className="px-3 py-2.5">
                              <button onClick={() => setExpandedUnmatchedCashier(isExpanded ? null : r.cashier.id)}
                                className={`flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium transition-colors ${isExpanded ? "bg-blue-100 text-blue-700" : "bg-muted hover:bg-muted/80 text-muted-foreground"}`}>
                                <Sparkles className="w-3.5 h-3.5"/>{isExpanded ? "إخفاء" : "اقتراح بديل"}
                              </button>
                            </td>
                            <td className="px-3 py-2.5">
                              <div className="flex flex-col gap-1">
                                <button onClick={() => handleHoldCashier(r.cashier)}
                                  title="علّق هذه الفاتورة لمطابقتها لاحقاً"
                                  className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-purple-50 text-purple-700 border border-purple-200 hover:bg-purple-100 transition-colors w-fit">
                                  <CreditCard className="w-3.5 h-3.5"/>علّق
                                </button>
                                  <button type="button" onClick={() => handleToggleReviewSnapshot(r.cashier, undefined, "فاتورة غير مطابقة")}
                                    title="نسخ الفاتورة إلى قائمة المراجعة" aria-label={`إضافة فاتورة ${r.cashier.name} للمراجعة`}
                                    aria-pressed={reviewItems.some(item => item.id === `review-cashier:${transactionKey(r.cashier)}`)}
                                    className={`flex w-fit items-center gap-1 rounded-lg border px-2.5 py-1.5 text-xs transition-colors ${reviewItems.some(item => item.id === `review-cashier:${transactionKey(r.cashier)}`) ? "border-amber-300 bg-amber-100 text-amber-800" : "border-amber-200 bg-amber-50 text-amber-700 hover:bg-amber-100"}`}>
                                    <Star className={`h-3.5 w-3.5 ${reviewItems.some(item => item.id === `review-cashier:${transactionKey(r.cashier)}`) ? "fill-current" : ""}`} />مراجعة
                                  </button>
                                  {isMahmoudWalletCashier(r.cashier) && (
                                    <button onClick={() => handleMoveToMahmoudWallet(r.cashier)}
                                      title="ترحيل هذه الفاتورة إلى محفظة شخصية"
                                      className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-violet-50 text-violet-700 border border-violet-200 hover:bg-violet-100 transition-colors w-fit">
                                      <CreditCard className="w-3.5 h-3.5"/>ترحيل لمحفظة شخصية
                                    </button>
                                  )}
                                  {visaNum && (
                                  <button onClick={() => handleMoveToVisa(r.cashier, "unmatched")}
                                    title="نقل هذه الفاتورة للفيزا"
                                    className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-teal-600 text-white border border-teal-600 hover:bg-teal-700 transition-colors w-fit">
                                    <CreditCard className="w-3.5 h-3.5"/>نقل للفيزا
                                  </button>
                                )}
                              </div>
                            </td>
                          </tr>
                          {isExpanded && (
                            <tr className="bg-blue-50 border-t border-blue-100">
                              <td colSpan={6} className="px-4 py-3">
                                {suggestions.length === 0 ? (
                                  <p className="text-xs text-muted-foreground">
                                    {availableBanks.length === 0 ? "جميع الحوالات البنكية المتاحة تم رفضها أو حفظها لهذا الكاشير" : "لم يتم إيجاد اقتراحات مشابهة من البنك (لا الاسم ولا المبلغ متطابقان)."}
                                  </p>
                                ) : (
                                  <div className="space-y-2">
                                    <p className="text-xs font-medium text-blue-700 mb-2">أقرب الحوالات البنكية المشابهة (مرتبة حسب قوة الاقتراح):</p>
                                    {suggestions.map(({ bank, score, matchType, amountDiff, reason }) => {
                                      const matchTypeLabel = matchType === "firstSecond" ? "👤 الأول+الثاني" :
                                                            matchType === "fourthName" ? "👤 الرابع" :
                                                            matchType === "exact" ? "تطابق تام" :
                                                            matchType === "typo" ? "✏️ اسم تقريبي" :
                                                            matchType === "amount_only" ? "مبلغ فقط (ضعيف)" : "غير محدد";
                                      const weak = matchType === "amount_only";
                                      return (
                                        <div key={bank.id} className={`flex items-center gap-3 p-2.5 bg-white rounded-lg border ${weak ? "border-gray-200" : "border-blue-200"}`}>
                                          <div className="flex-1 text-xs">
                                            <span className="font-medium">{bank.description}</span>
                                            <span className={`text-xs font-semibold mr-2 ${bank.type === "مدفوع" ? "text-red-600" : "text-green-600"}`}>{bank.type}</span>
                                            <span className="text-muted-foreground mr-2">{fmtNum(bank.rawAmount)}</span>
                                            <span className={`mr-1 ${weak ? "text-gray-400" : "text-blue-500"}`}>({(score * 100).toFixed(0)}% تطابق)</span>
                                            <span className="text-gray-400 mr-1">| {matchTypeLabel}</span>
                                            {amountDiff > 0.01 && <span className="text-orange-500 mr-1">| فرق: {fmtNum(amountDiff)}</span>}
                                            <div className="text-[10px] text-gray-400 mt-0.5">{reason}</div>
                                          </div>
                                          <button onClick={() => handleAcceptSuggestion(r.cashier, bank)}
                                            className="flex items-center gap-1 px-2.5 py-1.5 bg-green-600 text-white rounded-lg text-xs hover:bg-green-700 transition-colors">
                                            <Check className="w-3.5 h-3.5"/>قبول
                                          </button>
                                          <button onClick={() => handleRejectSuggestion(r.cashier, bank)}
                                            className="flex items-center gap-1 px-2.5 py-1.5 bg-red-50 text-red-600 border border-red-200 rounded-lg text-xs hover:bg-red-100 transition-colors">
                                            <X className="w-3.5 h-3.5"/>رفض
                                          </button>
                                          <button onClick={() => handleHoldBank(bank)}
                                            title="علّق حوالة البنك هذه لمطابقتها لاحقاً"
                                            className="flex items-center gap-1 px-2.5 py-1.5 bg-purple-50 text-purple-700 border border-purple-200 rounded-lg text-xs hover:bg-purple-100 transition-colors">
                                            <CreditCard className="w-3.5 h-3.5"/>علّق
                                          </button>
                                        </div>
                                      );
                                    })}
                                    <p className="text-xs text-muted-foreground mt-1">
                                      أو <button onClick={() => { setPage("manual"); }} className="text-blue-600 hover:underline">اذهب للمطابقة اليدوية</button> لربط يدوي كامل
                                    </p>
                                  </div>
                                )}
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                    {!visibleUCashierRows.length && <tr><td colSpan={8} className="py-10 text-center text-muted-foreground text-xs">لا توجد سجلات تطابق الفلاتر</td></tr>}
                  </tbody>
                </table>
              )}

              {tab === "uBank" && (
                <table className="w-full text-sm">
                  <thead><tr className="bg-muted text-xs">
                    <th className="px-3 py-2.5 text-right">البيان</th>
                    <th className="px-3 py-2.5 text-right">النوع</th>
                    <th className="px-3 py-2.5 text-right">المبلغ</th>
                    <th className="px-3 py-2.5 text-right">نوع الحساب المسجل</th>
                    <th className="px-3 py-2.5 text-right">السبب</th>
                    <th className="px-3 py-2.5 text-right">تعليق</th>
                  </tr></thead>
                  <tbody>
                    {visibleUBankRows.map((r: any, i: number) => {
                      const typeColor = r.bank.type === "مدفوع" ? "text-red-600" : "text-green-600";
                      const isFromHeld = !!r.bank._fromHeld;
                      return (
                        <tr key={i} data-default-reference={defaultReferenceForRow(r.bank)} className={`border-t transition-colors ${isFromHeld ? "bg-purple-50/50 hover:bg-purple-50" : "hover:bg-muted/20"}`}>
                          <td className="px-3 py-2.5 text-muted-foreground">
                            {r.bank.description}
                            <span className="mr-2"><DefaultReferenceButton row={r.bank} onOpen={handleOpenReference} /></span>
                            {isFromHeld && <span className="mr-1.5 text-[9px] px-1.5 py-0.5 rounded bg-purple-100 text-purple-700 border border-purple-300">📦 من ملف سابق</span>}
                          </td>
                          <td className={`px-3 py-2.5 text-xs font-semibold ${typeColor}`}>{r.bank.type}</td>
                          <td className="px-3 py-2.5 font-mono font-semibold text-orange-700">{fmtNum(r.bank.rawAmount)}</td>
                          <td className="px-3 py-2.5">{r.bank.accountType || "—"}</td>
                          <td className="px-3 py-2.5">
                            <span className={`px-1.5 py-0.5 rounded text-xs font-medium ${r.reason === "اختلاف في الاسم" ? "bg-amber-100 text-amber-800" : "bg-gray-100 text-gray-700"}`}>{r.reason}</span>
                          </td>
                          <td className="px-3 py-2.5">
                            <button onClick={() => handleHoldBank(r.bank)}
                              title="علّق هذه الحوالة لمطابقتها لاحقاً"
                              className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-purple-50 text-purple-700 border border-purple-200 hover:bg-purple-100 transition-colors">
                              <CreditCard className="w-3.5 h-3.5"/>علّق
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                    {!visibleUBankRows.length && <tr><td colSpan={6} className="py-10 text-center text-muted-foreground text-xs">لا توجد سجلات تطابق الفلاتر</td></tr>}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
