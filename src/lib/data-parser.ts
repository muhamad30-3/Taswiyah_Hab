import * as XLSX from "xlsx"
import { parseTabularRows, type StatementCell } from "./account-statement-processor"
import { readUploadedDocument } from "./document-reader"

export type DataCategory =
  | "SALES_AND_CATEGORIES"
  | "ELECTRONIC_WALLETS"
  | "OPERATING_EXPENSES"
  | "DEBTS_AND_LEDGERS"

export type ExpenseSubCategory =
  | "packaging_supplies"
  | "staff_meals"
  | "fuel_and_energy"
  | "maintenance"
  | "general_expenses"

export type ParsedFileData = {
  summary: {
    totalAmount: number
    recordCount: number
    currency?: string
  }
  details: {
    totalSales?: number
    cashAmount?: number
    app1Amount?: number
    app2Amount?: number
    app3Amount?: number
    newDebtsAmount?: number
    totalExpenseAmount?: number
    actualStatementTotal?: number
    collectedDebtsAmount?: number
    vendorPaymentsMade?: number
    litersCount?: number
    costPerLiter?: number
    cogs?: number
    cogsDetected?: boolean
    profit?: number
  }
  categoriesBreakdown?: Array<{
    categoryName: string
    sales: number
    cogs: number
    profit: number
  }>
  rawParsedRows: Record<string, unknown>[]
  amountDetected: boolean
  amountColumn: string | null
  truncated: boolean
  extractionNote: string | null
}

export type ParsedData = ParsedFileData

type ParsedTable = { headers: string[]; rows: Record<string, unknown>[] }

const maximumParsedRows = 10_000
const maximumRawRowsBytes = 4 * 1024 * 1024
const currencyPatterns = [
  { code: "ILS", pattern: /(?:₪|شيكل|nis|ils)/i },
  { code: "USD", pattern: /(?:\$|usd|دولار)/i },
  { code: "EUR", pattern: /(?:€|eur|يورو)/i },
  { code: "JOD", pattern: /(?:jod|دينار)/i },
]

const amountHeaderHints = [
  "amount", "total", "sales", "revenue", "gross", "net", "paid", "received", "debit", "credit",
  "المبلغ", "مبلغ", "اجمالي", "المجموع", "المبيعات", "الايرادات", "مدين", "دائن", "مدفوع",
  "مستلم", "صرف", "قبض", "قيمة", "تكلفة", "التكلفة",
]
const cashHeaderHints = ["cash", "cash amount", "نقد", "نقدا", "كاش", "نقدي", "الصندوق"]
const appHeaderHints = [
  ["jawwal pay", "jawwalpay", "jawwal", "جوال باي", "جوال باى", "جوال"],
  ["bank app", "bank", "banking app", "تطبيق البنك", "البنك", "بنك"],
  ["app3", "app 3", "wallet 3", "application 3", "تطبيق 3", "تطبيق3", "محفظة 3", "المحفظة الثالثة"],
]
const debtHeaderHints = ["debt", "due", "remaining", "balance due", "آجل", "اجل", "المتبقي", "ذمم", "دين", "الديون"]
const litersHeaderHints = ["liters", "litres", "liter", "litre", "اللترات", "لترات", "لتر", "كمية الوقود"]
const debitHints = ["debit", "withdrawal", "paid", "مدين", "مدفوع", "صرف", "خروج", "سحب"]
const creditHints = ["credit", "deposit", "received", "inflow", "دائن", "مستلم", "قبض", "دخول", "ايداع", "إيداع"]
const descriptionHeaderPattern = /(?:description|detail|item|name|statement|narrative|البيان|بيان|وصف|تفاصيل|الصنف|البند|الجهة|الطرف)/
const summaryRowPattern = /(?:^|\s)(?:total|subtotal|grand total|opening balance|closing balance|balance|الاجمالي|اجمالى|المجموع|الرصيد|رصيد)(?:\s|$)/
const pdfMovementHints = /(?:transaction|transfer|payment|withdrawal|deposit|debit|credit|purchase|refund|حركة|تحويل|دفع|سحب|إيداع|ايداع|مدين|دائن|شراء|استرجاع|قيد|قبض|تحصيل|سداد|مورد)/i
const pdfCreditHints = /(?:credit|deposit|inflow|received|دائن|مستلم|قبض|تحصيل|ايداع|إيداع|إيداع)/i
const pdfDebitHints = /(?:debit|withdrawal|outflow|paid|مدين|مدفوع|صرف|سحب|سداد)/i
const currencyHints = /(?:₪|شيكل|nis|ils|\$|usd|دولار|€|eur|يورو|jod|دينار)/i

function normalizeText(value: unknown) {
  return String(value ?? "")
    .normalize("NFKC")
    .toLocaleLowerCase("ar")
    .replace(/[\u064b-\u065f\u0670\u0640]/g, "")
    .replace(/[أإآ]/g, "ا")
    .replace(/[_-]+/g, " ")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim()
}

export function classifyExpenseByLabel(customLabel: string): ExpenseSubCategory {
  const normalizedLabel = normalizeText(customLabel)
  if (/(نيلون|اكياس|تغليف|طباعة)/.test(normalizedLabel)) return "packaging_supplies"
  if (/(طعام|غذاء|اكل|عمال|وجبات|ضيافة)/.test(normalizedLabel)) return "staff_meals"
  if (/(سولار|بنزين|محروقات|مولد|ديزل|وقود)/.test(normalizedLabel)) return "fuel_and_energy"
  if (/(صيانة|تصليح|قطع غيار)/.test(normalizedLabel)) return "maintenance"
  return "general_expenses"
}

function normalizeDigits(value: string) {
  return value.replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
}

function toAmount(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null
  if (typeof value !== "string" || !value.trim()) return null
  let normalized = normalizeDigits(value).replace(/[\u200e\u200f\u202a-\u202e\u2066-\u2069]/g, "").trim()
  const negative = /^\(.*\)$/.test(normalized) || /^\s*-/.test(normalized) || /-\s*$/.test(normalized)
  normalized = normalized
    .replace(/[₪$€£]/g, "")
    .replace(/(?:شيكل|nis|ils|usd|eur|jod|دولار|يورو|دينار)/gi, "")
    .replace(/\s/g, "")
    .replace(/[^\d,.'+-]/g, "")
  if (!/\d/.test(normalized)) return null
  if (normalized.includes(",") && normalized.includes(".")) {
    if (normalized.lastIndexOf(",") > normalized.lastIndexOf(".")) {
      normalized = normalized.replace(/\./g, "").replace(",", ".")
    } else {
      normalized = normalized.replace(/,/g, "")
    }
  }
  else if (normalized.includes(",")) normalized = /,\d{1,2}$/.test(normalized) ? normalized.replace(",", ".") : normalized.replace(/,/g, "")
  const amount = Number(normalized.replace(/'/g, ""))
  if (!Number.isFinite(amount)) return null
  return negative ? -Math.abs(amount) : amount
}

function containsHint(value: string, hints: string[]) {
  const normalized = normalizeText(value)
  return hints.some(hint => normalized.includes(normalizeText(hint)))
}

function cellToJson(value: StatementCell): unknown {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? "" : value.toISOString()
  if (typeof value === "number" && !Number.isFinite(value)) return ""
  return value ?? ""
}

function sheetRowsToTable(sheetName: string, rows: StatementCell[][]): ParsedTable | null {
  const parsed = parseTabularRows(rows)
  if (!parsed) return null
  return {
    headers: parsed.headers,
    rows: parsed.rows.map(values => {
      const row: Record<string, unknown> = { "اسم الورقة": sheetName }
      parsed.headers.forEach((header, index) => { row[header] = cellToJson(values[index]) })
      return row
    }),
  }
}

function amountColumn(headers: string[]) {
  const candidates = headers.flatMap(header => {
    const normalized = normalizeText(header)
    if (!amountHeaderHints.some(hint => normalized.includes(normalizeText(hint)))) return []
    let score = 20
    if (/(?:total|اجمالي|المجموع|المبيعات|الايرادات)/.test(normalized)) score += 50
    if (/(?:amount|مبلغ|قيمة|تكلفة)/.test(normalized)) score += 40
    if (/(?:debit|credit|مدين|دائن|صرف|قبض)/.test(normalized)) score -= 30
    return [{ header, score }]
  })
  return candidates.sort((left, right) => right.score - left.score)[0] ?? null
}

function findHeader(headers: string[], pattern: RegExp) {
  return headers.find(header => pattern.test(normalizeText(header))) ?? null
}

function findHintHeader(headers: string[], hints: string[]) {
  return headers.find(header => containsHint(header, hints)) ?? null
}

function findIndexedAppHeader(headers: string[], index: number) {
  const pattern = new RegExp(`(?:app(?:lication)?|تطبيق|محفظة)\\s*${index}`, "i")
  return headers.find(header => pattern.test(normalizeText(header).replace(/\s/g, ""))) ?? null
}

function getRowAmount(row: Record<string, unknown>, header: string | null) {
  return header ? toAmount(row[header]) : null
}

function meaningfulRows(table: ParsedTable) {
  const descriptionHeaders = table.headers.filter(header => descriptionHeaderPattern.test(normalizeText(header)))
  if (!descriptionHeaders.length) return table.rows
  return table.rows.filter(row => !descriptionHeaders.some(header => summaryRowPattern.test(normalizeText(row[header]))))
}

function getCurrency(rows: Record<string, unknown>[], headers: string[]) {
  const sample = [...headers, ...rows.slice(0, 20).flatMap(row => Object.values(row).map(String))].join(" ")
  return currencyPatterns.find(currency => currency.pattern.test(sample))?.code ?? "ILS"
}

function limitRawRows(rows: Record<string, unknown>[]) {
  const limited: Record<string, unknown>[] = []
  let bytes = 0
  for (const row of rows) {
    if (limited.length >= maximumParsedRows) break
    const rowBytes = JSON.stringify(row).length * 2
    if (bytes + rowBytes > maximumRawRowsBytes) break
    limited.push(row)
    bytes += rowBytes
  }
  return { rows: limited, truncated: limited.length < rows.length }
}

function result(
  table: ParsedTable,
  summaryRows: Record<string, unknown>[],
  totalAmount: number,
  amountDetected: boolean,
  amountHeader: string | null,
  details: ParsedFileData["details"],
  categoriesBreakdown?: ParsedFileData["categoriesBreakdown"],
  note?: string,
): ParsedFileData {
  const rows = limitRawRows(summaryRows)
  return {
    summary: {
      totalAmount: amountDetected ? totalAmount : 0,
      recordCount: summaryRows.length,
      currency: getCurrency(summaryRows, table.headers),
    },
    details,
    ...(categoriesBreakdown?.length ? { categoriesBreakdown } : {}),
    rawParsedRows: rows.rows,
    amountDetected,
    amountColumn: amountHeader,
    truncated: rows.truncated,
    extractionNote: note ?? (amountDetected ? null : "تم استخراج الصفوف، لكن تعذر تحديد عمود مبلغ بثقة. راجع الملف يدوياً."),
  }
}

function parseSalesTable(table: ParsedTable): ParsedFileData {
  const rows = meaningfulRows(table)
  const headers = table.headers
  const cashHeader = findHintHeader(headers, cashHeaderHints)
  const appHeaders = appHeaderHints.map((hints, index) =>
    findHintHeader(headers, hints) ?? findIndexedAppHeader(headers, index + 1))
  const debtHeader = findHintHeader(headers, debtHeaderHints)
  const paymentHeaders = new Set([cashHeader, ...appHeaders, debtHeader].filter((header): header is string => !!header))
  const totalHeader = headers
    .map(header => {
      const normalized = normalizeText(header)
      if (paymentHeaders.has(header) || /(?:cash|نقد|كاش|jawwal|bank|wallet|app|تطبيق|محفظة|آجل|اجل|debt|remaining)/.test(normalized)) return null
      const score = /(?:invoice total|total invoice|grand total|total sales|sales total|اجمالي الفاتورة|قيمة الفاتورة|المبلغ الكلي)/.test(normalized)
        ? 100
        : /(?:total|اجمالي|المجموع|المبيعات|amount|مبلغ)/.test(normalized) ? 50 : 0
      return score ? { header, score } : null
    })
    .filter((candidate): candidate is { header: string; score: number } => candidate !== null)
    .sort((left, right) => right.score - left.score)[0]?.header
    ?? (() => {
      const genericAmount = amountColumn(headers)?.header
      return genericAmount && !paymentHeaders.has(genericAmount) ? genericAmount : null
    })()
  let totalSales = 0
  let cashAmount = 0
  let app1Amount = 0
  let app2Amount = 0
  let app3Amount = 0
  let newDebtsAmount = 0
  let salesDetected = false

  for (const row of rows) {
    const total = getRowAmount(row, totalHeader)
    if (total !== null) {
      totalSales += total
      salesDetected = true
    }
    cashAmount += Math.abs(getRowAmount(row, cashHeader) ?? 0)
    app1Amount += Math.abs(getRowAmount(row, appHeaders[0]) ?? 0)
    app2Amount += Math.abs(getRowAmount(row, appHeaders[1]) ?? 0)
    app3Amount += Math.abs(getRowAmount(row, appHeaders[2]) ?? 0)
    newDebtsAmount += Math.max(0, getRowAmount(row, debtHeader) ?? 0)
  }

  const hasPaymentColumns = !!cashHeader || appHeaders.some(Boolean) || !!debtHeader
  const details: ParsedFileData["details"] = {
    ...(salesDetected ? { totalSales } : {}),
    ...(cashHeader ? { cashAmount } : {}),
    ...(appHeaders[0] ? { app1Amount } : {}),
    ...(appHeaders[1] ? { app2Amount } : {}),
    ...(appHeaders[2] ? { app3Amount } : {}),
    ...(debtHeader ? { newDebtsAmount } : {}),
  }

  const categoryHeader = findHeader(headers, /(?:category|department|اسم القسم|القسم|الفئة)/)
  const cogsHeader = findHeader(headers, /(?:cogs|cost of goods|cost|التكلفة|تكلفة البضاعة|كلفة)/)
  const profitHeader = findHeader(headers, /(?:profit|الربح|ربح)/)
  const salesHeader = totalHeader
  const categories = new Map<string, { categoryName: string; sales: number; cogs: number; profit: number }>()
  if (categoryHeader && (cogsHeader || profitHeader)) {
    for (const row of rows) {
      const categoryName = String(row[categoryHeader] ?? "").trim()
      if (!categoryName) continue
      const entry = categories.get(categoryName) ?? { categoryName, sales: 0, cogs: 0, profit: 0 }
      const sales = getRowAmount(row, salesHeader) ?? 0
      const cogs = getRowAmount(row, cogsHeader)
      const profit = getRowAmount(row, profitHeader)
      entry.sales += sales
      entry.cogs += cogsHeader ? Math.abs(cogs ?? 0) : salesHeader && profitHeader ? sales - (profit ?? 0) : 0
      entry.profit += profitHeader ? profit ?? 0 : salesHeader && cogsHeader ? sales - Math.abs(cogs ?? 0) : 0
      categories.set(categoryName, entry)
    }
  }
  const categoriesBreakdown = Array.from(categories.values())
  if (categoriesBreakdown.length && !salesDetected) salesDetected = true
  const totalAmount = totalHeader && salesDetected
    ? totalSales
    : categoriesBreakdown.length
      ? categoriesBreakdown.reduce((sum, item) => sum + item.sales, 0)
      : cashAmount + app1Amount + app2Amount + app3Amount + newDebtsAmount
  if (categoriesBreakdown.length) {
    details.cogsDetected = !!cogsHeader || (!!salesHeader && !!profitHeader)
    if (details.cogsDetected) details.cogs = categoriesBreakdown.reduce((sum, item) => sum + item.cogs, 0)
    if (profitHeader || (salesHeader && cogsHeader)) {
      details.profit = categoriesBreakdown.reduce((sum, item) => sum + item.profit, 0)
    }
  }
  return result(
    table,
    rows,
    totalAmount,
    salesDetected || hasPaymentColumns,
    totalHeader ?? cashHeader ?? appHeaders.find(Boolean) ?? debtHeader,
    details,
    categoriesBreakdown,
    salesDetected ? undefined : "تم استخراج الصفوف، لكن لم يتم التعرف على إجمالي المبيعات أو أعمدة المقبوضات.",
  )
}

function parseExpenseTable(table: ParsedTable, label: string): ParsedFileData {
  const rows = meaningfulRows(table)
  const headers = table.headers
  const amountHeaderValue = amountColumn(headers)?.header
  const litersHeader = findHintHeader(headers, litersHeaderHints)
  const explicitTotalHeader = findHeader(headers, /(?:total expense|expense total|grand total|amount|المبلغ|القيمة|التكلفة|اجمالي المصروف|المجموع)/)
  const selectedAmountHeader = explicitTotalHeader ?? amountHeaderValue
  let totalExpenseAmount = 0
  let amountDetected = false
  let litersCount = 0
  let litersDetected = false
  for (const row of rows) {
    const amount = getRowAmount(row, selectedAmountHeader)
    if (amount !== null) {
      totalExpenseAmount += Math.abs(amount)
      amountDetected = true
    }
    const liters = getRowAmount(row, litersHeader)
    if (liters !== null) {
      litersCount += Math.abs(liters)
      litersDetected = true
    }
  }
  const details: ParsedFileData["details"] = { ...(amountDetected ? { totalExpenseAmount } : {}) }
  if (litersDetected) {
    details.litersCount = litersCount
    if (litersCount > 0 && amountDetected) details.costPerLiter = totalExpenseAmount / litersCount
  }
  const isFuel = classifyExpenseByLabel(label) === "fuel_and_energy"
  return result(
    table,
    rows,
    totalExpenseAmount,
    amountDetected,
    selectedAmountHeader,
    details,
    undefined,
    isFuel && !litersDetected
      ? "تم تحليل المبلغ، لكن لم يتم العثور على عمود اللترات."
      : amountDetected ? undefined : "لم يتم التعرف على عمود المبلغ أو التكلفة.",
  )
}

function parseWalletTable(table: ParsedTable): ParsedFileData {
  const rows = meaningfulRows(table)
  const headers = table.headers
  const creditHeader = findHintHeader(headers, creditHints)
  const debitHeader = findHintHeader(headers, debitHints)
  const genericAmountHeader = amountColumn(headers)?.header ?? findHeader(headers, /(?:^amount$|^المبلغ$|^قيمة الحركة$)/)
  const directionHeader = findHeader(headers, /(?:direction|transaction type|type|مدين دائن|نوع الحركة|نوع العملية|نوع القيد)/)
  const descriptionHeader = findHeader(headers, descriptionHeaderPattern)
  let actualStatementTotal = 0
  let amountDetected = false
  let usedCreditColumn = false

  for (const row of rows) {
    if (creditHeader) {
      const credit = getRowAmount(row, creditHeader)
      if (credit !== null) {
        actualStatementTotal += Math.max(0, credit)
        amountDetected = true
        usedCreditColumn = true
        continue
      }
    }
    if (genericAmountHeader) {
      const amount = getRowAmount(row, genericAmountHeader)
      const direction = `${row[directionHeader ?? ""] ?? ""} ${row[descriptionHeader ?? ""] ?? ""}`
      if (amount !== null && !pdfDebitHints.test(direction) && (pdfCreditHints.test(direction) || (!debitHeader && !directionHeader))) {
        actualStatementTotal += Math.max(0, amount)
        amountDetected = true
      }
    }
  }
  return result(
    table,
    rows,
    actualStatementTotal,
    amountDetected,
    creditHeader ?? genericAmountHeader,
    amountDetected ? { actualStatementTotal } : {},
    undefined,
    amountDetected
      ? usedCreditColumn || directionHeader
        ? undefined
        : "تم جمع عمود المبلغ لعدم وجود مؤشر يحدد اتجاه الحركة؛ راجع ما إذا كانت القيم مقبوضة فقط."
      : "لم يتم تحديد عمود المقبوضات/الإيداعات. لا يتم جمع المدين تلقائياً حتى لا يختلط مع المقبوضات.",
  )
}

function parseDebtTable(table: ParsedTable): ParsedFileData {
  const rows = meaningfulRows(table)
  const amountHeaderValue = amountColumn(table.headers)?.header
    ?? findHeader(table.headers, /(?:amount|المبلغ|القيمة|مدين|دائن|مدفوع|مستلم)/)
  const descriptionHeader = findHeader(table.headers, descriptionHeaderPattern)
  const customerHeader = findHeader(table.headers, /(?:customer|client|الزبون|العميل|الزبائن)/)
  let collectedDebtsAmount = 0
  let vendorPaymentsMade = 0
  let totalAmount = 0
  let amountDetected = false
  let collectionDetected = false
  let paymentDetected = false
  for (const row of rows) {
    const amount = getRowAmount(row, amountHeaderValue)
    if (amount === null) continue
    amountDetected = true
    const text = `${row[descriptionHeader ?? ""] ?? ""} ${row[customerHeader ?? ""] ?? ""}`
    if (/(?:تحصيل|سداد.*زبون|قبض.*دين|customer.*collect|debt.*collect|collected.*debt)/i.test(text)) {
      collectedDebtsAmount += Math.abs(amount)
      collectionDetected = true
    } else if (/(?:سداد.*(?:مورد|supplier|vendor)|(?:مورد|supplier|vendor).*سداد|دفع.*(?:مورد|supplier|vendor)|vendor.*payment|supplier.*payment|(?:مدفوع|payment|paid))/i.test(text)) {
      vendorPaymentsMade += Math.abs(amount)
      paymentDetected = true
    } else {
      totalAmount += Math.abs(amount)
    }
  }
  totalAmount += collectedDebtsAmount + vendorPaymentsMade
  const details: ParsedFileData["details"] = {
    ...(collectionDetected ? { collectedDebtsAmount } : {}),
    ...(paymentDetected ? { vendorPaymentsMade } : {}),
  }
  return result(
    table,
    rows,
    totalAmount,
    amountDetected,
    amountHeaderValue,
    details,
    undefined,
    amountDetected
      ? collectionDetected || paymentDetected
        ? undefined
        : "تم جمع مبالغ الدفتر، لكن لم يتم التعرف على وصف يميز تحصيل ديون الزبائن من مدفوعات الموردين."
      : "لم يتم تحديد عمود للمبالغ في دفتر الديون والتحصيلات.",
  )
}

function parseTabularTable(table: ParsedTable, category: DataCategory, label: string) {
  switch (category) {
    case "SALES_AND_CATEGORIES": return parseSalesTable(table)
    case "OPERATING_EXPENSES": return parseExpenseTable(table, label)
    case "ELECTRONIC_WALLETS": return parseWalletTable(table)
    case "DEBTS_AND_LEDGERS": return parseDebtTable(table)
  }
}

function parseSpreadsheet(file: File, category: DataCategory, label: string): Promise<ParsedFileData> {
  return file.arrayBuffer().then(buffer => {
    const workbook = XLSX.read(buffer, {
      type: "array",
      cellDates: true,
      raw: file.name.toLocaleLowerCase("en").endsWith(".csv"),
      codepage: 65001,
    })
    const tables = workbook.SheetNames.flatMap(sheetName => {
      const sheet = workbook.Sheets[sheetName]
      if (!sheet) return []
      const rows = XLSX.utils.sheet_to_json<StatementCell[]>(sheet, {
        header: 1,
        defval: "",
        raw: true,
        blankrows: false,
      })
      const table = sheetRowsToTable(sheetName, rows)
      return table ? [table] : []
    })
    if (!tables.length) throw new Error("لم يتم العثور على جدول بصف عناوين وبيانات في الملف.")

    const combined: ParsedFileData[] = tables.map(table => parseTabularTable(table, category, label))
    const amountDetected = combined.some(parsed => parsed.amountDetected)
    const details: ParsedFileData["details"] = {}
    for (const key of [
      "totalSales", "cashAmount", "app1Amount", "app2Amount", "app3Amount", "newDebtsAmount",
      "totalExpenseAmount", "actualStatementTotal", "collectedDebtsAmount", "vendorPaymentsMade",
      "litersCount", "cogs", "profit",
    ] as const) {
      const values = combined.flatMap(parsed => parsed.details[key] === undefined ? [] : [parsed.details[key]!])
      if (values.length) details[key] = values.reduce((sum, value) => sum + value, 0)
    }
    if (combined.some(parsed => parsed.details.cogsDetected !== undefined)) {
      details.cogsDetected = combined.some(parsed => parsed.details.cogsDetected === true)
    }
    if ((details.litersCount ?? 0) > 0 && (details.totalExpenseAmount ?? 0) > 0) {
      details.costPerLiter = details.totalExpenseAmount! / details.litersCount!
    }
    const rawRows = tables.flatMap(table => table.rows)
    const limitedRows = limitRawRows(rawRows)
    const categoryTotals = new Map<string, { categoryName: string; sales: number; cogs: number; profit: number }>()
    for (const category of combined.flatMap(parsed => parsed.categoriesBreakdown ?? [])) {
      const total = categoryTotals.get(category.categoryName) ?? { ...category, sales: 0, cogs: 0, profit: 0 }
      total.sales += category.sales
      total.cogs += category.cogs
      total.profit += category.profit
      categoryTotals.set(category.categoryName, total)
    }
    const categoriesBreakdown = Array.from(categoryTotals.values())
    const summary = {
      totalAmount: combined.reduce((sum, parsed) => sum + parsed.summary.totalAmount, 0),
      recordCount: combined.reduce((sum, parsed) => sum + parsed.summary.recordCount, 0),
      currency: combined.find(parsed => parsed.summary.currency)?.summary.currency ?? "ILS",
    }
    const notes = Array.from(new Set(combined.flatMap(parsed => parsed.extractionNote ? [parsed.extractionNote] : [])))
    return {
      summary: { ...summary, totalAmount: amountDetected ? summary.totalAmount : 0 },
      details,
      ...(categoriesBreakdown.length ? { categoriesBreakdown } : {}),
      rawParsedRows: limitedRows.rows,
      amountDetected,
      amountColumn: Array.from(new Set(combined.flatMap(parsed => parsed.amountColumn ? [parsed.amountColumn] : []))).join("، ") || null,
      truncated: limitedRows.truncated,
      extractionNote: notes.length ? notes.join(" ") : null,
    }
  })
}

function parsePdfAmount(line: string): number | null {
  if (summaryRowPattern.test(normalizeText(line))) return null
  const normalizedLine = normalizeDigits(line)
    .replace(/\d{4}[/-]\d{1,2}[/-]\d{1,2}/g, " ")
    .replace(/\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/g, " ")
  const matches = [...normalizedLine.matchAll(/[-(]?\s*\d[\d,]*(?:\.\d{1,2})?\s*(?:₪|شيكل|nis|ils|\$|usd|دولار|€|eur|يورو)?/gi)]
  if (!matches.length) return null
  const candidate = matches[matches.length - 1][0].trim()
  if (!currencyHints.test(candidate) && !/\.\d{1,2}\b/.test(candidate) && !pdfMovementHints.test(line)) return null
  return toAmount(candidate)
}

function parsePdfTable(text: string, category: DataCategory): ParsedFileData {
  const lines = text.split(/\r?\n/).map(line => line.replace(/\s+/g, " ").trim()).filter(Boolean)
  const rows: Record<string, unknown>[] = []
  let actualStatementTotal = 0
  let totalExpenseAmount = 0
  let collectedDebtsAmount = 0
  let vendorPaymentsMade = 0
  for (const line of lines) {
    const amount = parsePdfAmount(line)
    if (amount === null) continue
    const description = line.replace(String(Math.abs(amount)), "").trim()
    rows.push({
      التاريخ: line.match(/\d{4}[/-]\d{1,2}[/-]\d{1,2}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/)?.[0] ?? "",
      البيان: description,
      المبلغ: amount,
    })
    if (category === "ELECTRONIC_WALLETS") {
      if (pdfCreditHints.test(line) && !pdfDebitHints.test(line)) actualStatementTotal += Math.max(0, amount)
    } else if (category === "OPERATING_EXPENSES") {
      totalExpenseAmount += Math.abs(amount)
    } else if (category === "DEBTS_AND_LEDGERS") {
      if (/(?:تحصيل|قبض.*دين|debt.*collect)/i.test(line)) collectedDebtsAmount += Math.abs(amount)
      else if (/(?:مورد|supplier|vendor|سداد|دفع)/i.test(line)) vendorPaymentsMade += Math.abs(amount)
    }
  }
  const details: ParsedFileData["details"] = {}
  if (category === "ELECTRONIC_WALLETS" && actualStatementTotal > 0) details.actualStatementTotal = actualStatementTotal
  if (category === "OPERATING_EXPENSES" && rows.length) details.totalExpenseAmount = totalExpenseAmount
  if (category === "DEBTS_AND_LEDGERS") {
    if (collectedDebtsAmount) details.collectedDebtsAmount = collectedDebtsAmount
    if (vendorPaymentsMade) details.vendorPaymentsMade = vendorPaymentsMade
  }
  const amountDetected = category === "ELECTRONIC_WALLETS" ? actualStatementTotal > 0
    : category === "OPERATING_EXPENSES" || category === "DEBTS_AND_LEDGERS" ? rows.length > 0
      : false
  const totalAmount = category === "ELECTRONIC_WALLETS" ? actualStatementTotal
    : category === "OPERATING_EXPENSES" ? totalExpenseAmount
      : category === "DEBTS_AND_LEDGERS" ? rows.reduce((sum, row) => sum + Math.abs(toAmount(row.المبلغ) ?? 0), 0)
        : 0
  const limitedRows = limitRawRows(rows)
  return {
    summary: { totalAmount: amountDetected ? totalAmount : 0, recordCount: rows.length, currency: getCurrency(rows, ["المبلغ"]) },
    details,
    rawParsedRows: limitedRows.rows,
    amountDetected,
    amountColumn: amountDetected ? "المبلغ" : null,
    truncated: limitedRows.truncated,
    extractionNote: category === "SALES_AND_CATEGORIES"
      ? "تم استخراج نص PDF، لكن لا يمكن تصنيف مبيعات POS أو بنود الأقسام بثقة من هذا الكشف."
      : category === "ELECTRONIC_WALLETS" && !amountDetected && rows.length
        ? "تم استخراج مبالغ من PDF، لكن لم يتم التأكد من اتجاه المقبوضات. راجع الكشف يدوياً."
        : category === "DEBTS_AND_LEDGERS" && rows.length && !collectedDebtsAmount && !vendorPaymentsMade
          ? "تم استخراج مبالغ، لكن لم يتم التعرف على تحصيلات الزبائن أو مدفوعات الموردين من أوصاف PDF."
          : amountDetected ? null : "تم استخراج نص PDF دون تحديد حركات مالية ومبالغ بثقة. راجع الكشف يدوياً.",
  }
}

export async function parseIngestionFile(
  file: File,
  category: DataCategory,
  label: string,
  onProgress?: (status: string) => void,
): Promise<ParsedFileData> {
  const extension = file.name.split(".").pop()?.toLocaleLowerCase("en")
  if (extension === "xlsx" || extension === "csv") {
    onProgress?.("جارٍ قراءة الجداول وتصنيف الأعمدة...")
    return parseSpreadsheet(file, category, label)
  }
  if (extension === "pdf") {
    onProgress?.("جارٍ استخراج نص PDF...")
    const text = await readUploadedDocument(file, status => onProgress?.(status))
    return parsePdfTable(text, category)
  }
  throw new Error("الصيغ المدعومة هي XLSX وCSV وPDF فقط.")
}

export function upgradeParsedData(value: unknown): ParsedFileData | null {
  if (!value || typeof value !== "object") return null
  const candidate = value as Record<string, unknown>
  if (candidate.summary && typeof candidate.summary === "object") return value as ParsedFileData
  if (typeof candidate.totalAmount !== "number" || typeof candidate.itemsCount !== "number") return null
  return {
    summary: { totalAmount: candidate.totalAmount, recordCount: candidate.itemsCount, currency: "ILS" },
    details: {},
    rawParsedRows: Array.isArray(candidate.rawRows) ? candidate.rawRows as Record<string, unknown>[] : [],
    amountDetected: candidate.amountDetected === true,
    amountColumn: typeof candidate.amountColumn === "string" ? candidate.amountColumn : null,
    truncated: candidate.truncated === true,
    extractionNote: typeof candidate.extractionNote === "string" ? candidate.extractionNote : null,
  }
}
