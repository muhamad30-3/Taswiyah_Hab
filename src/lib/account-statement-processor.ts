export type StatementCell = string | number | boolean | Date | null | undefined

export type StatementSheet = { rows: StatementCell[][] }

export type StatementFormat = "ledger" | "bank" | "generic"

export type ExtractedStatement = {
  format: StatementFormat
  headers: string[]
  rows: StatementCell[][]
  headerBlocks: number
  sourceRows: number
}

type HeaderIndexes =
  | {
      format: "bank"
      indexes: number[]
      transactionIndexes: number[]
    }
  | {
      format: "ledger"
      indexes: number[]
      entryNumber: number
      transactionIndexes: number[]
    }

export const LEDGER_HEADERS = ["رقم القيد", "التاريخ", "المدين", "الدائن", "الرصيد", "البيان", "بيان القيد"]
export const BANK_HEADERS = ["التاريخ البنكي", "التاريخ الفعلي", "الإيضاحات", "تاريخ الحق", "مبالغ مدفوعة", "مبالغ مستلمة", "الرصيد (شيكل)"]

export function normalizeHeader(value: StatementCell): string {
  return String(value ?? "")
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[\u064B-\u065F\u0670\u0640]/g, "")
    .replace(/[أإآ]/g, "ا")
    .replace(/\s+/g, "")
}

export function cleanStatementText(value: StatementCell): StatementCell {
  if (typeof value !== "string") return value
  return value
    .replace(/(?:بطاقة\s+(?:ال)?ائتمان(?:ية)?|credit\s*card)\s*[/／]?\s*/gi, "")
    .replace(/^\s*\/\s*|\s*\/\s*$/g, "")
    .replace(/\s+/g, " ")
    .trim()
}

function headerCandidateScore(row: StatementCell[]): number {
  const indexes = row.map((cell, index) => hasValue(cell) ? index : -1).filter((index) => index >= 0)
  if (indexes.length < 2) return -1
  const normalized = indexes.map((index) => normalizeHeader(row[index]))
  if (new Set(normalized).size < 2) return -1
  const labels = indexes.filter((index) => typeof row[index] === "string" && !/^[\d.,-]+$/.test(String(row[index]).trim())).length
  if (labels / indexes.length < 0.75) return -1
  const hints = /date|description|details|narrative|name|debit|credit|money|amount|paid|received|balance|reference|transaction|رقم|تاريخ|بيان|وصف|تفاصيل|مدين|دائن|مبلغ|مدفوع|مستلم|رصيد|مرجع|حركة|ايضاح|ملاحظات/i
  return indexes.reduce((score, index) => score + (hints.test(String(row[index])) ? 1 : 0), 0)
}

export function parseTabularRows(rows: StatementCell[][]): { headers: string[]; rows: StatementCell[][] } | null {
  const nextNonEmptyRows = new Array<number>(rows.length).fill(-1)
  let nextNonEmpty = -1
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    nextNonEmptyRows[index] = nextNonEmpty
    if (rows[index].some(hasValue)) nextNonEmpty = index
  }
  let headerIndex = -1
  let bestScore = -1
  for (let index = 0; index < rows.length; index += 1) {
    const score = headerCandidateScore(rows[index])
    if (score < 0) continue
    const headerIndexes = rows[index].map((cell, column) => hasValue(cell) ? column : -1).filter((column) => column >= 0)
    const nextDataIndex = nextNonEmptyRows[index]
    const nextData = nextDataIndex >= 0 ? rows[nextDataIndex] : undefined
    if (!nextData || !headerIndexes.some((column) => hasValue(nextData[column]))) continue
    if (score > bestScore) {
      bestScore = score
      headerIndex = index
    }
  }
  if (headerIndex < 0) return null

  const sourceHeaders = rows[headerIndex]
  const columns = sourceHeaders.map((cell, index) => ({ index, header: hasValue(cell) ? String(cell).trim() : "" })).filter((column) => column.header)
  const headers: string[] = []
  const used = new Map<string, number>()
  for (const column of columns) {
    const count = (used.get(normalizeHeader(column.header)) ?? 0) + 1
    used.set(normalizeHeader(column.header), count)
    headers.push(count > 1 ? `${column.header} (${count})` : column.header)
  }
  const normalizedHeaders = columns.map((column) => normalizeHeader(column.header))
  const data = rows.slice(headerIndex + 1).filter((row) => row.some(hasValue) && !isMetadataRow(row)).flatMap((row) => {
    const normalizedRow = columns.map((column) => normalizeHeader(row[column.index]))
    if (normalizedRow.length === normalizedHeaders.length && normalizedRow.every((cell, index) => cell === normalizedHeaders[index])) return []
    if (!columns.some((column) => hasValue(row[column.index]))) return []
    return [columns.map((column) => row[column.index] ?? "")]
  })
  return { headers, rows: data }
}

function findHeaderIndexes(row: StatementCell[]): HeaderIndexes | null {
  const headers = row.map(normalizeHeader)
  const find = (predicate: (header: string) => boolean) => headers.findIndex(predicate)
  const bankIndexes = [
    find((header) => /تاريخ.*بنكي|بنكي.*تاريخ/.test(header)),
    find((header) => /تاريخ.*فعلي|فعلي.*تاريخ/.test(header)),
    find((header) => /ايضاحات|ملاحظات|clarification|description/.test(header)),
    find((header) => /تاريخ.*حق|حق.*تاريخ/.test(header)),
    find((header) => /مبالغ?مدفوعة|مدفوع|paid/.test(header)),
    find((header) => /مبالغ?مستلمة|مستلم|received/.test(header)),
    find((header) => /رصيد|balance/.test(header)),
  ]
  if (bankIndexes.slice(0, 6).every((index) => index >= 0)) {
    return { format: "bank", indexes: bankIndexes, transactionIndexes: bankIndexes }
  }

  const entryNumber = find((header) => /رقم/.test(header) && /قيد|حركة|معاملة|عملية|سند|مرجع/.test(header))
  const transactionIndexes = [
    entryNumber,
    find((header) => /تاريخ/.test(header) || header === "date"),
    find((header) => /مدين|debit/.test(header)),
    find((header) => /دائن|credit/.test(header)),
    find((header) => /رصيد|balance/.test(header)),
    find((header) => /^(ال)?بيان$|وصف|تفاصيل|description|details/.test(header)),
    find((header) => /بيان.*قيد|قيد.*بيان|نوعالحركة|نوعالعملية/.test(header)),
  ]
  if (
    entryNumber >= 0 &&
    transactionIndexes[1] >= 0 &&
    transactionIndexes[2] >= 0 &&
    transactionIndexes[3] >= 0 &&
    transactionIndexes[2] !== transactionIndexes[3] &&
    transactionIndexes.slice(5).some((index) => index >= 0)
  ) {
    return {
      format: "ledger",
      indexes: transactionIndexes,
      entryNumber,
      transactionIndexes: transactionIndexes.filter((index) => index >= 0),
    }
  }
  return null
}

function hasValue(value: StatementCell): boolean {
  return value !== null && value !== undefined && String(value).trim() !== ""
}

function isMetadataRow(row: StatementCell[]): boolean {
  return row.some((cell) => /issued\s+(?:on|by)|أصدرها|أصدرت\s+بتاريخ|نسخة\s+مرخصة|رقم\s+الصفحة/i.test(String(cell ?? "")))
}

export function isStatementBalanceSummaryRow(values: unknown[]): boolean {
  return values.some((value) => {
    const normalized = normalizeHeader(String(value ?? "")).replace(/\s+/g, "")
    return /^(?:الرصيد|رصيد)(?:(?:السابق|الافتتاحي|الافتتاح|الحالي|الختامي|النهائي|المدور|المرحل|اولالمدة|اخرالمدة|اولمدة|اخرمدة|بدايةالمدة|نهايةالمدة))?$/.test(normalized)
      || /^(?:(?:opening|closing|previous|current|ending|beginning)?balance|balance(?:broughtforward|broughtdown|carriedforward|carrieddown))$/.test(normalized)
  })
}

export function removeStatementBalances(
  headers: string[],
  rows: Array<Record<string, unknown>>,
): { headers: string[]; rows: Array<Record<string, unknown>> } {
  const keptHeaders = headers.filter((header) => {
    const normalized = normalizeHeader(header)
    return !normalized.includes("رصيد") && !normalized.includes("balance")
  })
  return {
    headers: keptHeaders,
    rows: rows
      .filter((row) => !isStatementBalanceSummaryRow(headers.map((header) => row[header])))
      .map((row) => Object.fromEntries(keptHeaders.map((header) => [header, cleanStatementText((row[header] ?? "") as StatementCell)]))),
  }
}

function isDateValue(value: StatementCell): boolean {
  if (value instanceof Date) return !Number.isNaN(value.getTime())
  const text = String(value ?? "").trim()
  if (/^\d{4,5}$/.test(text)) {
    const serial = Number(text)
    return serial >= 20000 && serial <= 80000
  }
  const match = text.match(/^(\d{1,4})[/-](\d{1,2})[/-](\d{1,4})$/)
  if (!match) return false
  const [, first, second, third] = match
  const [year, month, day] = first.length === 4
    ? [Number(first), Number(second), Number(third)]
    : [Number(third), Number(second), Number(first)]
  return year >= 1900 && year <= 2200 && month >= 1 && month <= 12 && day >= 1 && day <= 31
}

function extractPageDate(row: StatementCell[]): string {
  for (const cell of row) {
    const text = String(cell ?? "")
    const match = text.match(/(?:في|on)\s*[:：]?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{4}|\d{4}[/-]\d{1,2}[/-]\d{1,2})/i)
    if (match && isDateValue(match[1])) return match[1]
  }
  return ""
}

function isNumericAmount(value: StatementCell): boolean {
  if (typeof value === "number") return Number.isFinite(value)
  const text = String(value ?? "").trim().replace(/[,\s٬]/g, "").replace(/[^\d.+()-]/g, "")
  return /^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(text) || /^\((?:\d+\.?\d*|\.\d+)\)$/.test(text)
}

export function extractStatementRows(sheets: StatementSheet[]): ExtractedStatement | null {
  const special = extractRecognizedStatementRows(sheets)
  if (special) return special

  let headers: string[] | null = null
  const rows: StatementCell[][] = []
  let sourceRows = 0
  for (const sheet of sheets) {
    sourceRows += sheet.rows.length
    const parsed = parseTabularRows(sheet.rows)
    if (!parsed) continue
    if (!headers) headers = parsed.headers
    const sourceIndex = new Map(parsed.headers.map((header, index) => [normalizeHeader(header), index]))
    rows.push(...parsed.rows.map((row) => headers!.map((header) => {
      const index = sourceIndex.get(normalizeHeader(header))
      return index === undefined ? "" : row[index] ?? ""
    })))
  }
  return headers ? { format: "generic", headers, rows, headerBlocks: sheets.length, sourceRows } : null
}

function extractRecognizedStatementRows(sheets: StatementSheet[]): ExtractedStatement | null {
  const rows: StatementCell[][] = []
  let format: StatementFormat | null = null
  let headers: string[] = []
  let headerBlocks = 0
  let sourceRows = 0

  for (const sheet of sheets) {
    let active: HeaderIndexes | null = null
    let readingTransactions = false
    let pageDate = ""
    for (const sourceRow of sheet.rows) {
      sourceRows += 1
      const detectedPageDate = extractPageDate(sourceRow)
      if (detectedPageDate) pageDate = detectedPageDate
      const detected = findHeaderIndexes(sourceRow)
      if (detected) {
        active = detected
        readingTransactions = true
        headerBlocks += 1
        if (!format) {
          format = detected.format
          headers = detected.format === "bank" ? [...BANK_HEADERS] : [...LEDGER_HEADERS]
        } else if (format !== detected.format) {
          active = null
          readingTransactions = false
        }
        continue
      }
      if (!sourceRow.some(hasValue)) continue
      if (!active || !readingTransactions || active.format !== format) continue
      if (active.format === "bank") {
        const transaction = active.transactionIndexes.map((index) => sourceRow[index] ?? "")
        if (!isDateValue(transaction[0]) && !isDateValue(transaction[1])) continue
        if (!isNumericAmount(transaction[4]) && !isNumericAmount(transaction[5])) continue
        rows.push(transaction)
      } else {
        if (isMetadataRow(sourceRow)) continue
        const transaction = active.indexes.map((index) => index < 0 ? "" : sourceRow[index] ?? "")
        if (isDateValue(transaction[1])) {
          if (typeof transaction[1] === "string") transaction[1] = transaction[1].trim()
        } else if (pageDate) {
          transaction[1] = pageDate
        }
        if (!hasValue(transaction[0])) continue
        if (!transaction.some(hasValue)) continue
        rows.push(transaction)
      }
    }
  }
  return format && rows.length ? { format, headers, rows, headerBlocks, sourceRows } : null
}

export function mergeSelectedStatements(
  documents: Array<{ headers: string[]; rows: StatementCell[][]; selectedColumns: boolean[]; selectedRows: boolean[] }>,
): { headers: string[]; rows: StatementCell[][] } {
  const headers: string[] = []
  const indexes = new Map<string, number>()
  const normalize = (value: string) => normalizeHeader(value)
  documents.forEach((document) => document.headers.forEach((header, index) => {
    if (!document.selectedColumns[index]) return
    const key = normalize(header)
    if (!indexes.has(key)) {
      indexes.set(key, headers.length)
      headers.push(header)
    }
  }))
  const rows = documents.flatMap((document) => document.rows.flatMap((row, rowIndex) => {
    if (!document.selectedRows[rowIndex]) return []
    const merged = Array.from({ length: headers.length }, () => "" as StatementCell)
    document.headers.forEach((header, columnIndex) => {
      if (!document.selectedColumns[columnIndex]) return
      const target = indexes.get(normalize(header))
      if (target !== undefined) merged[target] = row[columnIndex] ?? ""
    })
    return [merged]
  }))
  return { headers, rows }
}
