import * as XLSX from "xlsx";
import {
  cleanStatementText,
  extractStatementRows,
  parseTabularRows,
  removeStatementBalances,
  type StatementCell,
} from "../lib/account-statement-processor";

type MergeRole = "bank" | "cash" | "jawwalBank" | "invoice";
type ColumnRole = "date" | "bankDate" | "actualDate" | "entryDescription" | "description" | "debit" | "credit" | "accountType" | "ref" | "totalAmount" | "movementAmount" | "direction";

interface ParseFile {
  fileName: string;
  buffer: ArrayBuffer;
}

interface ParseRequest {
  role: MergeRole;
  files: ParseFile[];
  baseHeaders: string[];
}

const HINTS = {
  date: ["date", "تاريخ", "value", "posting"],
  desc: ["description", "narrative", "detail", "إيضاح", "بيان", "وصف", "narr", "payee", "particular"],
  debit: ["debit", "مدين", "مستلم", "مبالغ مستلمة", "received", "deposit", "money in", "دخول", "إيداع"],
  credit: ["credit", "دائن", "مدفوع", "مبالغ مدفوعة", "paid", "withdrawal", "money out", "خروج", "سحب"],
  name: ["name", "اسم", "customer", "client", "employee", "موظف", "عميل", "الزبون", "بيان", "وصف", "narrative", "detail", "إيضاح"],
  accountType: ["account type", "account_type", "نوع الحساب", "نوع حساب", "نوع", "حساب", "account"],
  ref: ["reference", "ref", "transaction", "trx", "id", "reference no", "reference number", "مرجع", "رقم المرجع", "رقم الحركة", "رقم الحوالة", "transaction id", "txn", "voucher", "سند", "رقم"],
  totalAmount: ["المبلغ الكلي", "total amount", "gross amount"],
  movementAmount: ["مبلغ الحركة", "transaction amount", "amount"],
  direction: ["مدين/دائن", "مدين دائن", "debit credit", "transaction side", "نوع القيد"],
};

function normalizeColumnName(value: string): string {
  return value.toLowerCase().normalize("NFKC").replace(/[\u064b-\u065f\u0670]/g, "").replace(/[_-]+/g, " ").replace(/[^\p{L}\p{N}]+/gu, " ").replace(/\s+/g, " ").trim();
}

function columnRole(header: string, role: MergeRole): ColumnRole | null {
  const normalized = normalizeColumnName(header);
  if (role === "bank") {
    if (/^(?:actual date|التاريخ الفعلي|تاريخ فعلي|التاريخ الحقيقي|تاريخ حقيقي)$/.test(normalized)) return "actualDate";
    if (/^(?:bank date|التاريخ البنكي|تاريخ بنكي)$/.test(normalized)) return "bankDate";
  }
  if (role === "cash" && /^(?:بيان القيد|قيد بيان)$/.test(normalized)) return "entryDescription";
  if (role === "invoice" && /^(?:بيان القيد|قيد بيان)$/.test(normalized)) return "entryDescription";
  const roles: Array<[ColumnRole, string[]]> = [
    ["date", HINTS.date],
    ["description", role === "cash" || role === "invoice" ? HINTS.name : HINTS.desc],
    ["debit", HINTS.debit],
    ["credit", HINTS.credit],
    ["accountType", HINTS.accountType],
    ["ref", HINTS.ref],
  ];
  if (role === "jawwalBank") roles.push(["totalAmount", HINTS.totalAmount], ["movementAmount", HINTS.movementAmount], ["direction", HINTS.direction]);
  let best: { role: ColumnRole; score: number } | null = null;
  for (const [candidateRole, hints] of roles) {
    for (const hint of hints) {
      const candidate = normalizeColumnName(hint);
      const score = normalized === candidate ? 4 : candidate.length > 2 && normalized.includes(candidate) ? 1 : 0;
      if (score > (best?.score ?? 0)) best = { role: candidateRole, score };
    }
  }
  return best?.role ?? null;
}

function toNum(value: unknown): number {
  if (value === "" || value == null) return 0;
  if (typeof value === "number") return value;
  return parseFloat(String(value).replace(/,/g, "")) || 0;
}

function cleanJawwalPayLedger(headers: string[], rows: Record<string, unknown>[]) {
  const entryHeader = headers.find(header => /رقم\s*القيد/.test(normalizeColumnName(header)));
  const debitHeader = headers.find(header => /^(?:ال)?مدين$|^debit$/.test(normalizeColumnName(header)));
  const creditHeader = headers.find(header => /^(?:ال)?دائن$|^credit$/.test(normalizeColumnName(header)));
  const isLedger = !!entryHeader && !!debitHeader && !!creditHeader && headers.some(header => /رصيد|balance/i.test(header));
  const outputHeaders = isLedger ? headers.filter(header => !/رصيد|balance/i.test(header)) : headers;
  if (!isLedger) return { headers: outputHeaders, rows };
  const outputRows = rows.filter(row => {
    const entryNumber = String(row[entryHeader] ?? "").trim().replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)));
    return /^\d+$/.test(entryNumber) && (Math.abs(toNum(row[debitHeader])) > 0 || Math.abs(toNum(row[creditHeader])) > 0);
  }).map(row => Object.fromEntries(outputHeaders.map(header => [header, cleanStatementText((row[header] ?? "") as StatementCell)])));
  return { headers: outputHeaders, rows: outputRows };
}

function rowsToObjects(headers: string[], rows: StatementCell[][]): Record<string, unknown>[] {
  return rows.map(values => {
    const row: Record<string, unknown> = {};
    headers.forEach((header, index) => { row[header] = values[index] ?? ""; });
    return row;
  });
}

function cleanRows(headers: string[], rows: Record<string, unknown>[]) {
  return removeStatementBalances(headers, rows);
}

function fillDownInvoiceContext(headers: string[], rows: Record<string, unknown>[]) {
  const contextHeaders = headers.filter(header => {
    const normalized = normalizeColumnName(header).replace(/\s*\(\d+\)$/, "");
    const dateColumn = /^(?:(?:transaction|invoice|movement|value|posting) )?date$/.test(normalized)
      || /^(?:تاريخ|التاريخ|تاريخ الفاتورة|تاريخ الحركة|تاريخ المعاملة|تاريخ العملية)$/.test(normalized);
    const movementColumn = /^(?:transaction|transaction type|transaction id|transaction number|transaction reference|movement|movement type|operation type|operation number|نوع الحركة|نوع المعاملة|الحركة|بيان القيد|رقم الحركة|رقم المعاملة|رقم العملية|رقم القيد)$/.test(normalized);
    return dateColumn || movementColumn;
  });
  if (!contextHeaders.length) return rows;

  const lastValues = new Map<string, unknown>();
  return rows.map(row => {
    const next = { ...row };
    for (const header of contextHeaders) {
      const value = String(next[header] ?? "").trim();
      if (value) lastValues.set(header, next[header]);
      else if (lastValues.has(header)) next[header] = lastValues.get(header);
    }
    return next;
  });
}

function parseSheet(buffer: ArrayBuffer) {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheets = workbook.SheetNames.map(name => ({
    rows: XLSX.utils.sheet_to_json<StatementCell[]>(workbook.Sheets[name], { header: 1, defval: "", raw: false }),
  }));
  const statement = extractStatementRows(sheets);
  if (statement) return cleanRows(statement.headers, rowsToObjects(statement.headers, statement.rows));
  for (const sheet of sheets) {
    const parsed = parseTabularRows(sheet.rows);
    if (parsed) return cleanRows(parsed.headers, rowsToObjects(parsed.headers, parsed.rows));
  }
  throw new Error("لم أتمكن من العثور على صف عناوين وبيانات في الملف.");
}

function parseJawwalPaySheet(buffer: ArrayBuffer) {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const statement = extractStatementRows(workbook.SheetNames.map(name => ({
    rows: XLSX.utils.sheet_to_json<StatementCell[]>(workbook.Sheets[name], { header: 1, defval: "", raw: false }),
  })));
  if (statement?.format !== "ledger") return parseSheet(buffer);
  return cleanJawwalPayLedger(statement.headers, rowsToObjects(statement.headers, statement.rows));
}

function mergeParsedFiles(parsedFiles: Array<{ fileName: string; headers: string[]; rows: Record<string, unknown>[] }>, role: MergeRole, baseHeaders: string[]) {
  const headers = [...baseHeaders];
  const headerByRole = new Map<ColumnRole, string>();
  const headerByNormalizedName = new Map<string, string>();
  headers.forEach(header => {
    const normalized = normalizeColumnName(header);
    if (!headerByNormalizedName.has(normalized)) headerByNormalizedName.set(normalized, header);
    const detectedRole = columnRole(header, role);
    if (detectedRole && !headerByRole.has(detectedRole)) headerByRole.set(detectedRole, header);
  });

  for (const parsed of parsedFiles) {
    for (const header of parsed.headers) {
      const detectedRole = columnRole(header, role);
      const canonicalHeader = detectedRole ? headerByRole.get(detectedRole) : undefined;
      const normalizedHeader = normalizeColumnName(header);
      if (!canonicalHeader && !headerByNormalizedName.has(normalizedHeader)) {
        headers.push(header);
        headerByNormalizedName.set(normalizedHeader, header);
        if (detectedRole) headerByRole.set(detectedRole, header);
      }
    }
  }

  const rows = parsedFiles.flatMap(parsed => {
    const sourceToTarget = new Map<string, string>();
    parsed.headers.forEach(sourceHeader => {
      const detectedRole = columnRole(sourceHeader, role);
      const targetHeader = detectedRole ? headerByRole.get(detectedRole) : undefined;
      sourceToTarget.set(sourceHeader, headerByNormalizedName.get(normalizeColumnName(sourceHeader)) ?? targetHeader ?? sourceHeader);
    });
    return parsed.rows.map(sourceRow => {
      const row: Record<string, unknown> = {};
      headers.forEach(header => { row[header] = ""; });
      for (const [sourceHeader, value] of Object.entries(sourceRow)) {
        const targetHeader = sourceToTarget.get(sourceHeader);
        if (targetHeader) row[targetHeader] = value;
      }
      return row;
    });
  });
  return { headers, rows, fileName: parsedFiles.map(parsed => parsed.fileName).join(" + ") };
}

self.onmessage = (event: MessageEvent<ParseRequest>) => {
  try {
    const { role, files, baseHeaders } = event.data;
    const parsedFiles = files.map(file => {
      const parsed = role === "jawwalBank" ? parseJawwalPaySheet(file.buffer) : parseSheet(file.buffer);
      const cleaned = cleanRows(parsed.headers, parsed.rows);
      return {
        fileName: file.fileName,
        ...cleaned,
        rows: role === "invoice" ? fillDownInvoiceContext(cleaned.headers, cleaned.rows) : cleaned.rows,
      };
    });
    self.postMessage({ ok: true, ...mergeParsedFiles(parsedFiles, role, baseHeaders) });
  } catch (error) {
    self.postMessage({ ok: false, message: error instanceof Error ? error.message : "تعذّر قراءة الملف." });
  }
};
