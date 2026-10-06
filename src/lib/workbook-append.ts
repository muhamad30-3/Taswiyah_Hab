import * as XLSX from "xlsx-js-style"
import type { ExtractedContact, ExtractedDocument } from "./document-extractor"

export const CONTACT_WORKSHEET_NAME = "جهات الاتصال"
export const FINANCIAL_WORKSHEET_NAME = "الملخص المالي"

const contactColumns = ["اسم الشخص", "البريد الإلكتروني", "رقم الهاتف", "اسم المرسل"]
const contactColumnsWithTime = [...contactColumns, "تاريخ الرسالة", "وقت الرسالة"]

const headerCellStyle = {
  fill: { patternType: "solid", fgColor: { rgb: "17365D" } },
  font: { bold: true, color: { rgb: "FFFFFF" } },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
  border: {
    top: { style: "thin", color: { rgb: "D6DFEA" } },
    bottom: { style: "thin", color: { rgb: "D6DFEA" } },
    left: { style: "thin", color: { rgb: "D6DFEA" } },
    right: { style: "thin", color: { rgb: "D6DFEA" } },
  },
}
const bodyCellStyle = (alternate: boolean) => ({
  fill: { patternType: "solid", fgColor: { rgb: alternate ? "EAF2F8" : "FFFFFF" } },
  alignment: { vertical: "center", wrapText: true },
  border: {
    top: { style: "thin", color: { rgb: "D6DFEA" } },
    bottom: { style: "thin", color: { rgb: "D6DFEA" } },
    left: { style: "thin", color: { rgb: "D6DFEA" } },
    right: { style: "thin", color: { rgb: "D6DFEA" } },
  },
})

export function styleWorksheetTable(sheet: XLSX.WorkSheet, headerRow: number, firstDataRow: number, rowCount: number, columns: number[]) {
  styleTable(sheet, headerRow, firstDataRow, rowCount, columns)
}

function styleTable(
  sheet: XLSX.WorkSheet,
  headerRow: number,
  firstDataRow: number,
  rowCount: number,
  columns: number[],
  headerColumns = columns,
) {
  for (const column of headerColumns) {
    const headerAddress = XLSX.utils.encode_cell({ r: headerRow, c: column })
    if (sheet[headerAddress]) sheet[headerAddress].s = headerCellStyle
  }
  for (const column of columns) {
    for (let offset = 0; offset < rowCount; offset++) {
      const address = XLSX.utils.encode_cell({ r: firstDataRow + offset, c: column })
      if (sheet[address]) sheet[address].s = bodyCellStyle(offset % 2 === 1)
    }
  }
}

const headerGroups: string[][] = [
  ["date", "transactiondate", "التاريخ", "تاريخ الحركة", "تاريخ المعاملة", "التاريخ مقروء من الصورة"],
  ["time", "transactiontime", "الوقت", "وقت الحركة", "وقت المعاملة"],
  ["transactiontype", "type", "نوع الحركة", "نوع المعاملة", "البيان", "تفاصيل الحركة"],
  ["cashiername", "collectorname", "اسم المحصل", "اسم الكاشير", "اسم أمين الصندوق", "الكاشير"],
  ["cashiermobile", "cashierphone", "collectormobile", "رقم المحصل", "هاتف المحصل", "هاتف الكاشير"],
  ["clientname", "customername", "اسم العميل", "اسم الزبون", "اسم الزبون العميل", "اسم الشخص"],
  ["amount", "المبلغ", "مبلغ الحركة", "القيمة", "المبلغ المقروء"],
  ["email", "البريد الالكتروني", "الايميل", "البريد"],
  ["phone", "mobile", "telephone", "رقم الهاتف", "الهاتف", "الجوال", "رقم الجوال"],
  ["sender", "sendername", "اسم المرسل"],
  ["label", "البند"],
  ["value", "القيمة المالية"],
]

function normalizeHeader(value: unknown) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f\u064b-\u065f\u0670\u0640]/g, "")
    .toLowerCase()
    .replace(/[\u200e\u200f\u202a-\u202e\u2066-\u2069]/g, "")
    .replace(/[^a-z0-9\u0600-\u06ff]/g, "")
}

function canonicalHeader(value: unknown) {
  const normalized = normalizeHeader(value)
  return headerGroups.find(group => group.some(alias => normalizeHeader(alias) === normalized))?.[0] ?? normalized
}

function sourceTable(result: ExtractedDocument, keepMessageTimestamps: boolean) {
  const tables: Array<{ name: string; columns: string[]; rows: string[][]; preferredSheet: string }> = []
  if (result.contacts.length) {
    const columns = keepMessageTimestamps ? contactColumnsWithTime : contactColumns
    const rows = result.contacts.map((contact: ExtractedContact) => [
      contact.name,
      contact.email,
      contact.phone,
      contact.sender,
      ...(keepMessageTimestamps ? [contact.date, contact.time] : []),
    ])
    tables.push({ name: CONTACT_WORKSHEET_NAME, columns, rows, preferredSheet: "contact" })
  }
  if (result.rows.length) {
    tables.push({ name: "", columns: result.columns, rows: result.rows, preferredSheet: "" })
  }
  if (result.facts.length) {
    tables.push({
      name: FINANCIAL_WORKSHEET_NAME,
      columns: ["البند", "القيمة"],
      rows: result.facts.map(fact => [fact.label, fact.value]),
      preferredSheet: "financial",
    })
  }
  return tables
}

function preferredSheetNames(kind: string) {
  if (kind === "contact") return /جهات?\s*الاتصال|contacts?/i
  if (kind === "financial") return /الملخص\s*المالي|financial|summary/i
  return null
}

function getUsedCells(sheet: XLSX.WorkSheet) {
  return Object.entries(sheet).flatMap(([address, cell]) => {
    if (address.startsWith("!")) return []
    let coordinate: XLSX.CellAddress
    try {
      coordinate = XLSX.utils.decode_cell(address)
    } catch {
      return []
    }
    const value = cell as XLSX.CellObject | undefined
    return value && (value.v !== undefined || value.f !== undefined || value.w !== undefined)
      ? [{ ...coordinate, value: value.v }]
      : []
  })
}

function appendTableToSheet(sheet: XLSX.WorkSheet, columns: string[], rows: string[][]) {
  const usedCells = getUsedCells(sheet)
  if (!usedCells.length) {
    XLSX.utils.sheet_add_aoa(sheet, [columns, ...rows], { origin: 0 })
    styleTable(sheet, 0, 1, rows.length, columns.map((_, index) => index))
    return
  }

  const firstUsedRow = Math.min(...usedCells.map(cell => cell.r))
  const lastUsedRow = Math.max(...usedCells.map(cell => cell.r))
  const lastUsedColumn = Math.max(...usedCells.map(cell => cell.c))
  const scanEnd = Math.min(lastUsedRow, firstUsedRow + 39)
  let headerRow = -1
  let bestHeaderScore = 0

  for (let row = firstUsedRow; row <= scanEnd; row++) {
    let score = 0
    for (const cell of usedCells) {
      if (cell.r === row && columns.some(column => canonicalHeader(column) === canonicalHeader(cell.value))) {
        score++
      }
    }
    if (score > bestHeaderScore) {
      headerRow = row
      bestHeaderScore = score
    }
  }

  if (headerRow < 0) {
    const newHeaderRow = lastUsedRow + 1
    XLSX.utils.sheet_add_aoa(sheet, [columns, ...rows], { origin: { r: newHeaderRow, c: 0 } })
    styleTable(sheet, newHeaderRow, newHeaderRow + 1, rows.length, columns.map((_, index) => index))
    return
  }

  const targetHeaders = new Map<string, number>()
  const newHeaderColumns: number[] = []
  const headerCells = usedCells.filter(cell => cell.r === headerRow)
  for (const cell of headerCells) {
    const normalized = canonicalHeader(cell.value)
    if (normalized && !targetHeaders.has(normalized)) targetHeaders.set(normalized, cell.c)
  }

  let nextColumn = Math.max(lastUsedColumn, ...headerCells.map(cell => cell.c)) + 1
  const targetColumns = columns.map(column => {
    const normalized = canonicalHeader(column)
    const existing = targetHeaders.get(normalized)
    if (existing !== undefined) return existing
    const allocated = nextColumn++
    targetHeaders.set(normalized, allocated)
    newHeaderColumns.push(allocated)
    XLSX.utils.sheet_add_aoa(sheet, [[column]], { origin: { r: headerRow, c: allocated } })
    return allocated
  })

  const appendedRows = rows.map(row => {
    const target = Array.from({ length: nextColumn }, () => "")
    row.forEach((value, index) => { target[targetColumns[index]] = value ?? "" })
    return target
  })
  const startColumn = Math.min(...targetColumns)
  const alignedRows = appendedRows.map(row => row.slice(startColumn))
  const firstDataRow = lastUsedRow + 1
  XLSX.utils.sheet_add_aoa(sheet, alignedRows, { origin: { r: firstDataRow, c: startColumn } })
  styleTable(sheet, headerRow, firstDataRow, rows.length, targetColumns, newHeaderColumns)
}

export function appendExtractedDataToWorkbook(
  workbook: XLSX.WorkBook,
  selectedSheetName: string,
  result: ExtractedDocument,
  keepMessageTimestamps: boolean,
) {
  const tables = sourceTable(result, keepMessageTimestamps)
  if (!tables.length) throw new Error("لم يتم استخراج بيانات مناسبة للإضافة إلى Excel.")

  const appendedTo: Array<{ sheetName: string; rowCount: number }> = []
  for (const table of tables) {
    let sheetName = table.preferredSheet ? "" : selectedSheetName
    if (table.preferredSheet) {
      const namePattern = preferredSheetNames(table.preferredSheet)
      const existingSheet = workbook.SheetNames.find(name => namePattern?.test(name))
      if (existingSheet) sheetName = existingSheet
      else if (selectedSheetName) {
        const sheet = workbook.Sheets[selectedSheetName]
        const headers = sheet ? getUsedCells(sheet).map(cell => cell.value) : []
        const hasMatchingHeaders = table.columns.some(column => headers.some(header => canonicalHeader(header) === canonicalHeader(column)))
        if (hasMatchingHeaders) sheetName = selectedSheetName
      }
    }

    if (!sheetName) {
      const desired = table.name || "البيانات"
      sheetName = desired
      let suffix = 2
      while (workbook.SheetNames.includes(sheetName)) sheetName = `${desired} ${suffix++}`
      workbook.SheetNames.push(sheetName)
      workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet([])
    } else if (!workbook.SheetNames.includes(sheetName) || !workbook.Sheets[sheetName]) {
      throw new Error("ورقة العمل المحددة غير موجودة في ملف Excel.")
    }

    const target = workbook.Sheets[sheetName]
    if (!target) throw new Error("تعذر فتح ورقة العمل المحددة.")
    appendTableToSheet(target, table.columns, table.rows)
    appendedTo.push({ sheetName, rowCount: table.rows.length })
  }
  return {
    sheetName: appendedTo.map(item => item.sheetName).join("، "),
    rowCount: appendedTo.reduce((total, item) => total + item.rowCount, 0),
  }
}
