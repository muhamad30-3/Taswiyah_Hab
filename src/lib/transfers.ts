import * as XLSX from "xlsx"
import { getMonthKey, parseTableDate } from "./table-view"

export type TransferSource = "bank" | "wallet" | "platform"
export type TransferDirection = "incoming" | "outgoing"

export type TransferRecord = {
  id: string
  importId?: string
  accountName?: string
  source: TransferSource
  date: string
  reference: string
  name: string
  details: string
  currency: string
  amount: number
  direction: TransferDirection
  dedupeKey: string
}

export type TransferImport = {
  id: string
  source: TransferSource
  accountName?: string
  currency?: string
  fileName: string
  importedAt: string
  importedCount: number
  duplicateCount: number
}

export type TransferData = {
  records: TransferRecord[]
  imports: TransferImport[]
}

export type TransferColumns = {
  date: number | null
  reference: number | null
  name: number | null
  amount: number | null
  debit: number | null
  credit: number | null
  direction: number | null
  details: number | null
  currency: number | null
}

export type TransferSheet = {
  headers: string[]
  rows: unknown[][]
  columns: TransferColumns
}

export type ParsedTransferFile = {
  records: TransferRecord[]
  recordRows: Array<{ record: TransferRecord; rowNumber: number }>
  duplicates: Array<{
    rowNumber: number
    record: TransferRecord
    matchedRecord: TransferRecord
    matchedRowNumber: number | null
  }>
  mismatches: Array<{
    rowNumber: number
    record: TransferRecord
    matchedRecord: TransferRecord
    matchedRowNumber: number | null
  }>
  invalidRows: Array<{
    rowNumber: number
    values: unknown[]
    reason: string
  }>
  invalidCount: number
  duplicateCount: number
}

export const EMPTY_TRANSFER_DATA: TransferData = { records: [], imports: [] }

const HEADER_TERMS: Record<keyof TransferColumns, RegExp> = {
  date: /(?:date|تاريخ)/i,
  reference: /(?:reference|transaction\s*(?:id|no|number)|ref(?:erence)?\s*(?:no|number)?|رقم\s*(?:مرجعي|الحوالة|العملية)|معرف\s*العملية)/i,
  name: /(?:name|customer|recipient|beneficiary|sender|اسم|زبون|عميل|مستفيد|مرسل|مستلم)/i,
  amount: /(?:amount|value|المبلغ|القيمة)/i,
  debit: /(?:debit|paid|withdrawal|مبالغ\s*مدفوعة|مدين|مدفوع|سحب|صادر)/i,
  credit: /(?:credit|received|deposit|مبالغ\s*مستلمة|دائن|مستلم|إيداع|ايداع|وارد)/i,
  direction: /(?:direction|transaction\s*type|type|نوع\s*(?:العملية|الحركة)|اتجاه)/i,
  details: /(?:description|details|memo|narrative|الايضاحات|التفاصيل|البيان|الوصف|ملاحظات)/i,
  currency: /(?:currency|العملة|نوع\s*العملة)/i,
}

function normalizeHeader(value: unknown) {
  return String(value ?? "")
    .normalize("NFKC")
    .toLocaleLowerCase("ar")
    .replace(/[\u064b-\u065f\u0670\u0640]/g, "")
    .replace(/[أإآ]/g, "ا")
    .replace(/[^\p{L}\p{N}]+/gu, "")
}

function normalizeDigits(value: string) {
  return value
    .replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, digit => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)))
}

export function normalizeTransferSearch(value: unknown) {
  return normalizeDigits(String(value ?? ""))
    .normalize("NFKC")
    .toLocaleLowerCase("ar")
    .replace(/[\u064b-\u065f\u0670\u0640]/g, "")
    .replace(/[أإآ]/g, "ا")
    .replace(/[ىي]/g, "ي")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim()
}

function nameSkeleton(value: unknown) {
  const arabicToLatin: Record<string, string> = {
    ا: "a", ب: "b", ت: "t", ث: "th", ج: "j", ح: "h", خ: "kh",
    د: "d", ذ: "th", ر: "r", ز: "z", س: "s", ش: "sh", ص: "s",
    ض: "d", ط: "t", ظ: "z", ع: "", غ: "gh", ف: "f", ق: "q",
    ك: "k", ل: "l", م: "m", ن: "n", ه: "h", و: "w", ي: "y",
    ة: "h", ء: "", ئ: "y", ؤ: "w", ى: "y",
  }
  const normalized = normalizeTransferSearch(value).replace(/\s+/g, "")
  const phonetic = Array.from(normalized, character => arabicToLatin[character] ?? character).join("")
  return phonetic
    .replace(/[aeiou]/g, "")
    .replace(/(.)\1+/g, "$1")
}

function parseAmount(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null
  if (typeof value !== "string" || !value.trim()) return null
  let text = normalizeDigits(value).replace(/[\u200e\u200f\u202a-\u202e\u2066-\u2069]/g, "").trim()
  const negative = /^\(.*\)$/.test(text) || /^\s*-/.test(text) || /-\s*$/.test(text)
  text = text
    .replace(/[₪$€£٬]/g, "")
    .replace(/(?:شيكل|nis|ils|usd|eur|jod|دولار|يورو|دينار)/gi, "")
    .replace(/\s/g, "")
    .replace(/[^\d,.'+-]/g, "")
  if (!/\d/.test(text)) return null
  if (text.includes(",") && text.includes(".")) {
    text = text.lastIndexOf(",") > text.lastIndexOf(".")
      ? text.replace(/\./g, "").replace(",", ".")
      : text.replace(/,/g, "")
  } else if (text.includes(",")) {
    text = /,\d{1,2}$/.test(text) ? text.replace(",", ".") : text.replace(/,/g, "")
  }
  const amount = Number(text.replace(/'/g, ""))
  return Number.isFinite(amount) ? (negative ? -Math.abs(amount) : amount) : null
}

function createId() {
  return typeof globalThis.crypto?.randomUUID === "function"
    ? globalThis.crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`
}

function stableFingerprint(value: string) {
  let first = 0x811c9dc5
  let second = 0x9e3779b9
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index)
    first = Math.imul(first ^ code, 0x01000193)
    second = Math.imul(second ^ code, 0x85ebca6b)
  }
  return `${(first >>> 0).toString(16)}${(second >>> 0).toString(16)}`
}

function parseDirection(value: unknown): TransferDirection | null {
  const normalized = normalizeTransferSearch(value)
  if (/(incoming|received|credit|دائن|وارد|استلام|مستلم|قبض|ايداع)/.test(normalized)) return "incoming"
  if (/(outgoing|sent|debit|مدين|صادر|ارسال|مرسل|دفع|سحب)/.test(normalized)) return "outgoing"
  return null
}

function toIsoDate(value: unknown) {
  const date = parseTableDate(value)
  if (!date) return null
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`
}

function extractReference(details: string) {
  return details.match(/\/\s*(\d{8,})\s*\//)?.[1] ?? ""
}

function createRecord(
  source: TransferSource,
  date: string,
  reference: string,
  name: string,
  details: string,
  currency: string,
  amount: number,
  direction: TransferDirection,
  occurrence: number,
  accountName: string,
  dedupeCurrency: string,
): TransferRecord {
  const cleanReference = normalizeTransferSearch(reference).replace(/\s/g, "")
  const identity = cleanReference
    ? `${source}:${normalizeTransferSearch(accountName)}:${normalizeTransferSearch(dedupeCurrency)}:ref:${cleanReference}:amount:${Math.abs(amount)}`
    : `${source}:${normalizeTransferSearch(accountName)}:${normalizeTransferSearch(dedupeCurrency)}:row:${stableFingerprint(JSON.stringify([date, name, details, currency, amount, direction]))}:${occurrence}`
  return {
    id: createId(),
    source,
    date,
    reference: reference.trim(),
    name: name.trim(),
    details: details.trim(),
    currency: currency.trim(),
    amount: Math.abs(amount),
    direction,
    dedupeKey: identity,
  }
}

export function canonicalTransferDedupeKey(record: Pick<TransferRecord, "dedupeKey" | "reference" | "amount">) {
  if (!record.reference) return record.dedupeKey
  const marker = record.dedupeKey.lastIndexOf(":ref:")
  if (marker < 0) return record.dedupeKey
  const prefix = record.dedupeKey.slice(0, marker)
  const reference = normalizeTransferSearch(record.reference).replace(/\s/g, "")
  return `${prefix}:ref:${reference}:amount:${Math.abs(record.amount)}`
}

export function canonicalTransferReferenceKey(record: Pick<TransferRecord, "dedupeKey" | "reference">) {
  if (!record.reference) return ""
  const marker = record.dedupeKey.lastIndexOf(":ref:")
  if (marker < 0) return ""
  const prefix = record.dedupeKey.slice(0, marker)
  const reference = normalizeTransferSearch(record.reference).replace(/\s/g, "")
  return `${prefix}:ref:${reference}`
}

export function suggestTransferColumns(headers: string[]): TransferColumns {
  const columns = {} as TransferColumns
  for (const key of Object.keys(HEADER_TERMS) as Array<keyof TransferColumns>) {
    const index = headers.findIndex(header => HEADER_TERMS[key].test(normalizeHeader(header)))
    columns[key] = index >= 0 ? index : null
  }
  const actualDate = headers.findIndex(header => /(?:actual|فعلي)/i.test(normalizeHeader(header)) && HEADER_TERMS.date.test(normalizeHeader(header)))
  if (actualDate >= 0) columns.date = actualDate
  if ([columns.details, columns.amount, columns.debit, columns.credit].includes(columns.name)) columns.name = null
  if ([columns.amount, columns.debit, columns.credit].includes(columns.details)) columns.details = null
  if (columns.amount === columns.debit || columns.amount === columns.credit) columns.amount = null
  return columns
}

export async function readTransferFile(file: File): Promise<TransferSheet> {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true, dense: true })
  for (const sheetName of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], {
      header: 1,
      raw: false,
      defval: "",
      blankrows: false,
    })
    if (rows.length < 2) continue
    const headers = rows[0].map((value, index) => String(value || `عمود ${index + 1}`))
    return { headers, rows: rows.slice(1), columns: suggestTransferColumns(headers) }
  }
  throw new Error("لم يتم العثور على ورقة تحتوي على صف عناوين وبيانات.")
}

export function parseTransferRows(
  rows: unknown[][],
  columns: TransferColumns,
  source: TransferSource,
  accountName = "",
  importCurrency = "",
): ParsedTransferFile {
  if (columns.date === null) throw new Error("حدد عمود التاريخ قبل استيراد الملف.")
  if (columns.amount === null && columns.debit === null && columns.credit === null) {
    throw new Error("حدد عمود المبلغ أو عمودي المدفوع والمستلم.")
  }

  const records: TransferRecord[] = []
  const recordRows: ParsedTransferFile["recordRows"] = []
  const seenReferencesInFile = new Map<string, { record: TransferRecord; rowNumber: number }>()
  const rowOccurrences = new Map<string, number>()
  const duplicates: ParsedTransferFile["duplicates"] = []
  const mismatches: ParsedTransferFile["mismatches"] = []
  const invalidRows: ParsedTransferFile["invalidRows"] = []
  let invalidCount = 0

  for (const [rowIndex, row] of rows.entries()) {
    const rowNumber = rowIndex + 2
    const date = toIsoDate(row[columns.date])
    const details = columns.details === null ? "" : String(row[columns.details] ?? "").trim()
    const reference = extractReference(details)
      || (columns.reference === null ? "" : String(row[columns.reference] ?? "").trim())
    const name = columns.name === null ? "" : String(row[columns.name] ?? "").trim()
    const currency = columns.currency === null ? "" : String(row[columns.currency] ?? "").trim()
    const defaultName = name || details
    const candidates: Array<{ value: unknown; direction: TransferDirection }> = []

    if (columns.debit !== null || columns.credit !== null) {
      if (columns.debit !== null) candidates.push({ value: row[columns.debit], direction: "outgoing" })
      if (columns.credit !== null) candidates.push({ value: row[columns.credit], direction: "incoming" })
    } else if (columns.amount !== null) {
      const rawAmount = row[columns.amount]
      const amount = parseAmount(rawAmount)
      const direction = columns.direction === null ? null : parseDirection(row[columns.direction])
      if (amount !== null) {
        const inferredDirection = direction ?? (amount < 0 ? "outgoing" : "incoming")
        candidates.push({ value: rawAmount, direction: inferredDirection })
      }
    }

    let rowHadTransfer = false
    let rowHadDuplicate = false
    let rowHadMismatch = false
    let rowHadInvalidTransfer = false
    let invalidReason = "المبلغ مفقود أو غير صالح."
    for (const candidate of candidates) {
      const amount = parseAmount(candidate.value)
      if (amount === null || amount === 0) {
        if (String(candidate.value ?? "").trim()) {
          rowHadInvalidTransfer = true
          invalidReason = "المبلغ غير صالح."
        }
        continue
      }
      if (!date) {
        rowHadInvalidTransfer = true
        invalidReason = "التاريخ مفقود أو غير صالح."
        continue
      }
      const baseFingerprint = `${source}:row:${stableFingerprint(JSON.stringify([date, defaultName, details, currency, Math.abs(amount), candidate.direction]))}`
      const occurrence = reference ? 1 : (rowOccurrences.get(baseFingerprint) ?? 0) + 1
      if (!reference) rowOccurrences.set(baseFingerprint, occurrence)
      const record = createRecord(
        source,
        date,
        reference,
        defaultName,
        details,
        importCurrency || currency,
        amount,
        candidate.direction,
        occurrence,
        accountName,
        importCurrency,
      )
      if (reference) {
        const referenceKey = canonicalTransferReferenceKey(record)
        const matchedReference = seenReferencesInFile.get(referenceKey)
        if (matchedReference) {
          if (matchedReference.record.amount === record.amount) {
            duplicates.push({ rowNumber, record, matchedRecord: matchedReference.record, matchedRowNumber: matchedReference.rowNumber })
            rowHadDuplicate = true
          } else {
            mismatches.push({ rowNumber, record, matchedRecord: matchedReference.record, matchedRowNumber: matchedReference.rowNumber })
            rowHadMismatch = true
          }
          continue
        }
        seenReferencesInFile.set(referenceKey, { record, rowNumber })
      }
      records.push(record)
      recordRows.push({ record, rowNumber })
      rowHadTransfer = true
    }
    const hasData = row.some(value => String(value ?? "").trim() !== "")
    if (hasData && (rowHadInvalidTransfer || (!rowHadTransfer && !rowHadDuplicate && !rowHadMismatch))) {
      invalidCount++
      invalidRows.push({ rowNumber, values: row, reason: invalidReason })
    }
  }

  return { records, recordRows, duplicates, mismatches, invalidRows, invalidCount, duplicateCount: duplicates.length }
}

export function getTransferMonth(record: TransferRecord) {
  return getMonthKey(record.date)
}

export function transferMatchesSearch(record: TransferRecord, query: string) {
  return createTransferSearchMatcher(query)(record)
}

const transferSearchCache = new WeakMap<TransferRecord, {
  fields: string[]
  nameTokens: Array<{ value: string; skeleton: string }>
  nameSkeleton: string
}>()

export function createTransferSearchMatcher(query: string) {
  const normalizedQuery = normalizeTransferSearch(query)
  if (!normalizedQuery) return () => true
  const queryDate = toIsoDate(query)
  const queryTokens = normalizedQuery.split(/\s+/)
  const querySkeletons = queryTokens.map(nameSkeleton)
  const querySkeleton = nameSkeleton(query)
  return (record: TransferRecord) => {
    if (queryDate && record.date === queryDate) return true
    let searchable = transferSearchCache.get(record)
    if (!searchable) {
      searchable = {
        fields: [record.name, record.reference, record.details, record.currency, record.accountName, record.date, record.amount]
          .map(value => normalizeTransferSearch(value)),
        nameTokens: normalizeTransferSearch(record.name).split(/\s+/).filter(Boolean)
          .map(value => ({ value, skeleton: nameSkeleton(value) })),
        nameSkeleton: nameSkeleton(record.name),
      }
      transferSearchCache.set(record, searchable)
    }
    if (searchable.fields.some(value => value.includes(normalizedQuery))) return true
    const matchesAllNameTokens = queryTokens.every((token, index) => {
      const skeleton = querySkeletons[index]
      return searchable.nameTokens.some(nameToken =>
        nameToken.value.includes(token) ||
        (skeleton.length >= 3 && nameToken.skeleton.includes(skeleton)),
      )
    })
    return matchesAllNameTokens || (querySkeleton.length >= 3 && searchable.nameSkeleton.includes(querySkeleton))
  }
}
