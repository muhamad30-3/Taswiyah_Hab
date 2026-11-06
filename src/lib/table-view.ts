export function normalizeTableHeader(value: unknown) {
  return String(value ?? "")
    .normalize("NFKC")
    .toLocaleLowerCase("ar")
    .replace(/[\u064b-\u065f\u0670\u0640]/g, "")
    .replace(/[أإآ]/g, "ا")
    .replace(/\s+/g, "")
}

function normalizeDigits(value: string) {
  return value
    .replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, digit => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)))
}

export function parseTableDate(value: unknown): Date | null {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value
  if (typeof value === "number" && Number.isFinite(value) && value >= 20_000 && value <= 80_000) {
    const date = new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86_400_000)
    return Number.isNaN(date.getTime()) ? null : date
  }
  if (typeof value !== "string" && typeof value !== "number") return null
  const text = normalizeDigits(String(value)).trim()
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/)
  const regional = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})(?:[ T].*)?$/)
  if (!iso && !regional) return null
  let year: number
  let month: number
  let day: number
  if (iso) {
    year = Number(iso[1])
    month = Number(iso[2])
    day = Number(iso[3])
  } else {
    day = Number(regional![1])
    month = Number(regional![2])
    year = Number(regional![3])
    if (year < 100) year += year < 50 ? 2000 : 1900
  }
  const date = new Date(year, month - 1, day)
  if (year < 1900 || year > 2200 || date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null
  return date
}

export function getTableDateColumn(headers: string[], preferActualDate = true) {
  const dateIndexes = headers.flatMap((header, index) =>
    /date|تاريخ/.test(normalizeTableHeader(header)) ? [index] : [],
  )
  if (preferActualDate) {
    const actualDate = dateIndexes.find(index => /فعلي|actual/.test(normalizeTableHeader(headers[index])))
    if (actualDate !== undefined) return actualDate
  }
  return dateIndexes[0] ?? -1
}

export function getMonthKey(value: unknown) {
  const date = parseTableDate(value)
  return date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}` : null
}

export function formatMonthKey(monthKey: string) {
  const match = monthKey.match(/^(\d{4})-(\d{2})$/)
  if (!match) return monthKey
  return new Intl.DateTimeFormat("ar", { month: "long", year: "numeric", timeZone: "UTC" })
    .format(new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, 1)))
}

function numericValue(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null
  if (typeof value !== "string" || !value.trim()) return null
  let text = normalizeDigits(value).trim()
  const isNegative = /^\(.*\)$/.test(text) || /^-/.test(text) || /-$/.test(text)
  text = text
    .replace(/[₪$€£٬\s]/g, "")
    .replace(/^\((.*)\)$/, "$1")
    .replace(/[^\d.,']/g, "")
  if (!text || !/\d/.test(text)) return null
  if (text.includes(",") && !text.includes(".")) {
    text = /,\d{1,2}$/.test(text) ? text.replace(",", ".") : text.replace(/,/g, "")
  } else {
    text = text.replace(/,/g, "")
  }
  const number = Number(text.replace(/'/g, ""))
  if (!Number.isFinite(number)) return null
  return isNegative ? -Math.abs(number) : number
}

export function compareTableValues(left: unknown, right: unknown, dateColumn = false) {
  if (dateColumn) {
    const leftDate = parseTableDate(left)?.getTime()
    const rightDate = parseTableDate(right)?.getTime()
    if (leftDate !== undefined || rightDate !== undefined) {
      if (leftDate === undefined) return 1
      if (rightDate === undefined) return -1
      return leftDate - rightDate
    }
  }
  const leftNumber = numericValue(left)
  const rightNumber = numericValue(right)
  if (leftNumber !== null && rightNumber !== null) return leftNumber - rightNumber
  return String(left ?? "").localeCompare(String(right ?? ""), "ar", { numeric: true, sensitivity: "base" })
}
