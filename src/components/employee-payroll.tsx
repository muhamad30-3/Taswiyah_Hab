"use client"

import { useEffect, useRef, useState, type FormEvent, type KeyboardEvent } from "react"
import { Banknote, CalendarDays, Check, ChevronDown, FilePlus2, FileSpreadsheet, Lock, MessageCircle, Pencil, Plus, Printer, RotateCw, Trash2, Unlock, Users, X } from "lucide-react"
import * as XLSX from "xlsx"
import { centralStorageEnabled, centralStorageGet, centralStorageSet } from "../lib/central-storage"
import { REPORT_LOGO_PRINT_STYLES, reportLogoMarkup } from "../lib/report-branding"
import { useGlobalLoading } from "./global-loading"

type Employee = {
  id: string
  employeeNumber: string
  name: string
  dailyRate: number
  customerNumber: string
  phone: string
}

type PayrollLine = {
  employeeId: string
  employeeNumber: string
  name: string
  dailyRate: number
  phone?: string
  attendanceDays?: number
  gross: number
  bonus?: number
  advance: number
  customerNumber: string
  debtBalance: number | null
  debtCollectionEnabled?: boolean
  debtDeduction: number
  percentageDeduction: number
  percentageBase: number
}

type PayrollSheet = {
  id: string
  name: string
  createdAt: string
  paymentDate: string
  periodFrom: string
  periodTo: string
  days: number
  percentageEnabled: boolean
  percentage: number
  percentageBase: "salary" | "debt" | "manual"
  locked?: boolean
  lines: PayrollLine[]
}

type PayrollData = { employees: Employee[]; sheets: PayrollSheet[] }
type EmployeeDraft = Omit<Employee, "id" | "dailyRate"> & { id?: string; dailyRate: string }
type DebtDataset = {
  records?: Array<{
    customerNumber?: unknown
    balance?: unknown
    balanceDirection?: unknown
  }>
}

const EMPTY_DATA: PayrollData = { employees: [], sheets: [] }
const PAYROLL_STORAGE_KEY = "employee_payroll"
const CUSTOMER_DATASET_KEY = "customer_balances_dataset"
const INPUT_CLASS = "w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm outline-none focus:border-cyan-600 focus:ring-2 focus:ring-cyan-100"

function createId() {
  return typeof globalThis.crypto?.randomUUID === "function"
    ? globalThis.crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`
}

function normalizeUsername(username: string) {
  return username.normalize("NFKC").trim().toLocaleLowerCase("en-US")
}

function normalizeNumber(value: string) {
  return value
    .replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, digit => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)))
}

function amount(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100
}

function evaluateArithmeticExpression(value: string) {
  const expression = normalizeNumber(value).trim().replace(/[×]/g, "*").replace(/[÷]/g, "/").replace(/,/g, ".")
  const tokens: Array<number | string> = []
  let offset = 0
  const tokenPattern = /\s*(?:(\d+(?:\.\d*)?|\.\d+)|([()+\-*/]))/y
  while (offset < expression.length) {
    tokenPattern.lastIndex = offset
    const match = tokenPattern.exec(expression)
    if (!match) return null
    tokens.push(match[1] === undefined ? match[2] : Number(match[1]))
    offset = tokenPattern.lastIndex
  }
  if (!tokens.length) return null

  let index = 0
  function parsePrimary(): number {
    const token = tokens[index++]
    if (token === "+" || token === "-") {
      const operand = parsePrimary()
      return token === "-" ? -operand : operand
    }
    if (token === "(") {
      const result = parseExpression()
      if (tokens[index++] !== ")") throw new Error("parenthesis")
      return result
    }
    if (typeof token === "number") return token
    throw new Error("number")
  }
  function parseTerm(): number {
    let result = parsePrimary()
    while (tokens[index] === "*" || tokens[index] === "/") {
      const operator = tokens[index++]
      const operand = parsePrimary()
      if (operator === "/" && operand === 0) throw new Error("division")
      result = operator === "*" ? result * operand : result / operand
    }
    return result
  }
  function parseExpression(): number {
    let result = parseTerm()
    while (tokens[index] === "+" || tokens[index] === "-") {
      const operator = tokens[index++]
      const operand = parseTerm()
      result = operator === "+" ? result + operand : result - operand
    }
    return result
  }

  try {
    const result = parseExpression()
    return index === tokens.length && Number.isFinite(result) ? result : null
  } catch {
    return null
  }
}

function formatAmount(value: number) {
  return new Intl.NumberFormat("ar", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value)
}

function formatDailyRate(value: number) {
  return new Intl.NumberFormat("ar", { maximumFractionDigits: 10 }).format(value)
}

function lineBonus(line: PayrollLine) {
  return Number.isFinite(line.bonus) ? line.bonus! : 0
}

function lineTotal(line: PayrollLine) {
  return amount(line.gross + lineBonus(line))
}

function handlePayrollGridKeyDown(event: KeyboardEvent<HTMLInputElement>, column: string, rowIndex: number) {
  if (event.key !== "ArrowDown" && event.key !== "ArrowUp") return
  event.preventDefault()
  const direction = event.key === "ArrowDown" ? 1 : -1
  event.currentTarget
    .closest("tbody")
    ?.querySelector<HTMLInputElement>(`[data-payroll-column="${column}"][data-payroll-row="${rowIndex + direction}"]`)
    ?.focus()
}

function formatDate(value: string) {
  if (!value) return "—"
  return new Date(`${value}T12:00:00`).toLocaleDateString("ar")
}

function escapeHtml(value: unknown) {
  return String(value ?? "").replace(/[&<>"']/g, character => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  })[character]!)
}

function whatsappPhone(phone: string) {
  let digits = normalizeNumber(phone).replace(/\D/g, "")
  if (digits.startsWith("00")) digits = digits.slice(2)
  if (digits.startsWith("970") || digits.startsWith("972")) return digits
  if (digits.startsWith("0")) return `970${digits.slice(1)}`
  if (digits.length === 9 && digits.startsWith("5")) return `970${digits}`
  return digits
}

function openEmployeeWhatsApp(phone: string, message: string) {
  const digits = whatsappPhone(phone)
  if (digits.length < 8 || digits.length > 15) return false
  window.open(`https://wa.me/${digits}?text=${encodeURIComponent(message)}`, "_blank", "noopener,noreferrer")
  return true
}

function emptyEmployee(): EmployeeDraft {
  return { employeeNumber: "", name: "", dailyRate: "", customerNumber: "", phone: "" }
}

function emptySheetForm() {
  const today = new Date().toISOString().slice(0, 10)
  return {
    name: "",
    paymentDate: today,
    periodFrom: today,
    periodTo: today,
    days: 7,
  }
}

function readDebtBalances(value: unknown) {
  const dataset = value as DebtDataset | undefined
  const balances = new Map<string, number>()
  for (const record of dataset?.records ?? []) {
    if (typeof record.customerNumber !== "string" || typeof record.balance !== "number" || !Number.isFinite(record.balance)) continue
    if (record.balanceDirection !== "receivable" || record.balance <= 0) continue
    balances.set(record.customerNumber.trim(), amount(record.balance))
  }
  return balances
}

function isPayrollData(value: unknown): value is PayrollData {
  if (!value || typeof value !== "object") return false
  const candidate = value as PayrollData
  return Array.isArray(candidate.employees) && Array.isArray(candidate.sheets)
}

function roundLine(value: number) {
  return amount(Number.isFinite(value) ? Math.max(0, value) : 0)
}

function calculateDebtDeduction(gross: number, advance: number, debtBalance: number | null) {
  return roundLine(Math.min(Math.max(0, gross - advance), debtBalance ?? 0))
}

function buildPayrollLine(employee: Employee, days: number, debts: Map<string, number>): PayrollLine {
  const gross = amount(employee.dailyRate * days)
  const debtBalance = employee.customerNumber ? debts.get(employee.customerNumber) ?? null : null
  return {
    employeeId: employee.id,
    employeeNumber: employee.employeeNumber,
    name: employee.name,
    dailyRate: employee.dailyRate,
    phone: employee.phone ?? "",
    attendanceDays: days,
    gross,
    bonus: 0,
    advance: 0,
    customerNumber: employee.customerNumber,
    debtBalance,
    debtCollectionEnabled: true,
    debtDeduction: calculateDebtDeduction(gross, 0, debtBalance),
    percentageDeduction: 0,
    percentageBase: 0,
  }
}

function getAttendanceDays(line: PayrollLine, sheet: PayrollSheet) {
  return Number.isInteger(line.attendanceDays) ? line.attendanceDays! : sheet.days
}

function calculatePercentage(sheet: PayrollSheet, line: PayrollLine, base = line.percentageBase) {
  const limit = sheet.percentageBase === "debt"
    ? Math.min(lineTotal(line), line.debtBalance ?? 0)
    : lineTotal(line)
  return roundLine(Math.min(limit, base * sheet.percentage / 100))
}

export function EmployeePayroll({ username }: { username: string }) {
  const normalizedUsername = normalizeUsername(username)
  const payrollKey = `${PAYROLL_STORAGE_KEY}:${normalizedUsername}`
  const customerKey = `${CUSTOMER_DATASET_KEY}:${normalizedUsername}`
  const legacyCustomerKey = CUSTOMER_DATASET_KEY
  const dataRef = useRef<PayrollData>(EMPTY_DATA)
  const busyRef = useRef(false)
  const [data, setData] = useState<PayrollData>(EMPTY_DATA)
  const [loadedKey, setLoadedKey] = useState<string | null>(null)
  const [debts, setDebts] = useState<Map<string, number>>(new Map())
  const [debtDatasetReady, setDebtDatasetReady] = useState(false)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState("")
  const [notice, setNotice] = useState("")
  const [tab, setTab] = useState<"sheets" | "employees">("sheets")
  const [employeeDraft, setEmployeeDraft] = useState<EmployeeDraft | null>(null)
  const [sheetForm, setSheetForm] = useState(emptySheetForm)
  const [sheetDraft, setSheetDraft] = useState<PayrollSheet | null>(null)
  const [creatingSheet, setCreatingSheet] = useState(false)
  const [expandedSheetId, setExpandedSheetId] = useState<string | null>(null)
  const [lineDrafts, setLineDrafts] = useState<Record<string, { dailyRate: string; advance: string; bonus: string; percentageBase: string; attendanceDays: string }>>({})
  const [percentageDraft, setPercentageDraft] = useState<{ sheetId: string; percentage: string; base: PayrollSheet["percentageBase"] } | null>(null)
  const [overwritePrompt, setOverwritePrompt] = useState<(() => void) | null>(null)
  const [unlockDraft, setUnlockDraft] = useState<{ sheetId: string; password: string } | null>(null)
  const [unlockError, setUnlockError] = useState("")
  const [employeePayrollDialog, setEmployeePayrollDialog] = useState<{ sheetId: string; employeeId: string; dailyRate: string } | null>(null)

  useGlobalLoading(loadedKey !== payrollKey && !error, "جاري تحميل بيانات الرواتب...")
  useGlobalLoading(busy, "جاري معالجة بيانات الرواتب...")

  useEffect(() => {
    let active = true
    setLoadedKey(null)
    setData(EMPTY_DATA)
    dataRef.current = EMPTY_DATA
    setError("")
    setNotice("")
    setDebtDatasetReady(false)
    const load = async () => {
      try {
        const [storedPayroll, customerDataset, legacyDataset] = await Promise.all([
          centralStorageEnabled
            ? centralStorageGet<PayrollData>(payrollKey)
            : Promise.resolve(JSON.parse(localStorage.getItem(payrollKey) ?? "null") as PayrollData | null),
          centralStorageEnabled
            ? centralStorageGet<DebtDataset>(customerKey)
            : Promise.resolve(JSON.parse(localStorage.getItem(customerKey) ?? "null") as DebtDataset | null),
          centralStorageEnabled
            ? centralStorageGet<DebtDataset>(legacyCustomerKey)
            : Promise.resolve(JSON.parse(localStorage.getItem(legacyCustomerKey) ?? "null") as DebtDataset | null),
        ])
        if (storedPayroll !== undefined && storedPayroll !== null && !isPayrollData(storedPayroll)) {
          throw new Error("بيانات الرواتب المحفوظة غير صالحة.")
        }
        if (active) {
          let debtDeductionsMigrated = false
          const next = storedPayroll
            ? {
              ...storedPayroll,
              sheets: storedPayroll.sheets.map(sheet => {
                if (sheet.locked || sheet.percentageEnabled) return sheet
                let sheetChanged = false
                const lines = sheet.lines.map(line => {
                  const debtDeduction = calculateDebtDeduction(lineTotal(line), line.advance, line.debtBalance)
                  if (debtDeduction === line.debtDeduction) return line
                  debtDeductionsMigrated = true
                  sheetChanged = true
                  return { ...line, debtDeduction }
                })
                return sheetChanged ? { ...sheet, lines } : sheet
              }),
            }
            : EMPTY_DATA
          if (debtDeductionsMigrated) {
            if (centralStorageEnabled) await centralStorageSet(payrollKey, next)
            else localStorage.setItem(payrollKey, JSON.stringify(next))
          }
          dataRef.current = next
          setData(next)
          setDebts(readDebtBalances(customerDataset ?? legacyDataset))
          setDebtDatasetReady(true)
          setLoadedKey(payrollKey)
        }
      } catch (cause) {
        if (active) setError(cause instanceof Error ? `تعذر تحميل منصة الرواتب: ${cause.message}` : "تعذر تحميل منصة الرواتب.")
      }
    }
    void load()
    return () => { active = false }
  }, [customerKey, legacyCustomerKey, payrollKey])

  const saveData = async (next: PayrollData, successMessage: string) => {
    try {
      if (centralStorageEnabled) await centralStorageSet(payrollKey, next)
      else localStorage.setItem(payrollKey, JSON.stringify(next))
      dataRef.current = next
      setData(next)
      setNotice(successMessage)
      setError("")
      return true
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر الحفظ ولم يتم اعتماد التعديل: ${cause.message}` : "تعذر الحفظ ولم يتم اعتماد التعديل.")
      return false
    }
  }

  const startBusy = () => {
    if (busyRef.current || loadedKey !== payrollKey) return false
    busyRef.current = true
    setBusy(true)
    setError("")
    setNotice("")
    return true
  }

  const endBusy = () => {
    busyRef.current = false
    setBusy(false)
  }

  const refreshDebts = async () => {
    setError("")
    try {
      const current = centralStorageEnabled
        ? await centralStorageGet<DebtDataset>(customerKey)
        : JSON.parse(localStorage.getItem(customerKey) ?? "null") as DebtDataset | null
      let dataset = current
      if (!current) {
        dataset = centralStorageEnabled
          ? await centralStorageGet<DebtDataset>(legacyCustomerKey)
          : JSON.parse(localStorage.getItem(legacyCustomerKey) ?? "null") as DebtDataset | null
      }
      setDebts(readDebtBalances(dataset))
      setDebtDatasetReady(true)
      setNotice("تم تحديث أرصدة ديون الزبائن المرتبطة.")
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر تحديث كشف الديون: ${cause.message}` : "تعذر تحديث كشف الديون.")
    }
  }

  const saveEmployee = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!employeeDraft || !startBusy()) return
    const employeeNumber = employeeDraft.employeeNumber.trim()
    const name = employeeDraft.name.trim()
    const dailyRate = evaluateArithmeticExpression(employeeDraft.dailyRate)
    const customerNumber = employeeDraft.customerNumber.trim()
    if (!employeeNumber || !name || dailyRate === null || dailyRate < 0) {
      setError("أدخل رقم الموظف واسمه، وتأكد من صحة عملية الأجر اليومي وأن نتيجتها غير سالبة.")
      endBusy()
      return
    }
    if (dataRef.current.employees.some(item => item.employeeNumber === employeeNumber && item.id !== employeeDraft.id)) {
      setError("رقم الموظف مسجل مسبقاً.")
      endBusy()
      return
    }
    const employee: Employee = {
      id: employeeDraft.id ?? createId(),
      employeeNumber,
      name,
      dailyRate,
      customerNumber,
      phone: employeeDraft.phone.trim(),
    }
    const employees = employeeDraft.id
      ? dataRef.current.employees.map(item => item.id === employee.id ? employee : item)
      : [...dataRef.current.employees, employee]
    try {
      if (await saveData({ ...dataRef.current, employees }, employeeDraft.id ? "تم تحديث بيانات الموظف." : "تمت إضافة الموظف.")) {
        setEmployeeDraft(null)
      }
    } finally {
      endBusy()
    }
  }

  const deleteEmployee = async (employee: Employee) => {
    if (!window.confirm(`حذف الموظف «${employee.name}» من قائمة الموظفين؟ ستبقى كشوفه السابقة كما هي.`)) return
    if (!startBusy()) return
    try {
      const next = { ...dataRef.current, employees: dataRef.current.employees.filter(item => item.id !== employee.id) }
      await saveData(next, "تم حذف الموظف من القائمة. لم تتغير الكشوف السابقة.")
    } finally {
      endBusy()
    }
  }

  const createSheet = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!startBusy()) return
    const days = Number(normalizeNumber(String(sheetForm.days)))
    if (!sheetForm.name.trim() || !sheetForm.paymentDate || !sheetForm.periodFrom || !sheetForm.periodTo
      || sheetForm.periodTo < sheetForm.periodFrom || !Number.isInteger(days) || days < 1 || days > 366) {
      setError("أكمل اسم الكشف والتواريخ، وتأكد أن عدد الأيام صحيح وأن نهاية الفترة لا تسبق بدايتها.")
      endBusy()
      return
    }
    const sheet: PayrollSheet = {
      id: createId(),
      name: sheetForm.name.trim(),
      createdAt: new Date().toISOString(),
      paymentDate: sheetForm.paymentDate,
      periodFrom: sheetForm.periodFrom,
      periodTo: sheetForm.periodTo,
      days,
      percentageEnabled: false,
      percentage: 0,
      percentageBase: "debt",
      lines: dataRef.current.employees.map(employee => buildPayrollLine(employee, days, debts)),
    }
    try {
      if (await saveData({ ...dataRef.current, sheets: [sheet, ...dataRef.current.sheets] }, "تم إنشاء كشف الرواتب وحفظه.")) {
        setCreatingSheet(false)
        setExpandedSheetId(sheet.id)
        setSheetForm(emptySheetForm())
      }
    } finally {
      endBusy()
    }
  }

  const editSheetMetadata = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!sheetDraft || !startBusy()) return
    const days = Number(normalizeNumber(String(sheetDraft.days)))
    if (!sheetDraft.name.trim() || !sheetDraft.paymentDate || !sheetDraft.periodFrom || !sheetDraft.periodTo
      || sheetDraft.periodTo < sheetDraft.periodFrom || !Number.isInteger(days) || days < 1 || days > 366) {
      setError("أكمل اسم الكشف والتواريخ، وتأكد من صحة الفترة وعدد الأيام.")
      endBusy()
      return
    }
    const oldSheet = dataRef.current.sheets.find(sheet => sheet.id === sheetDraft.id)
    if (!oldSheet) {
      setError("كشف الرواتب المطلوب تعديله غير موجود.")
      endBusy()
      return
    }
    if (oldSheet.locked) {
      setError("هذا الكشف مثبت. فك التثبيت أولاً لتعديل بياناته.")
      endBusy()
      return
    }
    const updated: PayrollSheet = {
      ...sheetDraft,
      name: sheetDraft.name.trim(),
      days,
      lines: sheetDraft.lines.map(line => {
        const previousAttendanceDays = getAttendanceDays(line, oldSheet)
        const attendanceDays = previousAttendanceDays === oldSheet.days ? days : previousAttendanceDays
        const gross = amount(line.dailyRate * attendanceDays)
        const nextLine = { ...line, attendanceDays, gross }
        if (!sheetDraft.percentageEnabled) {
          nextLine.debtDeduction = calculateDebtDeduction(lineTotal(nextLine), line.advance, line.debtBalance)
        }
        const percentageBase = sheetDraft.percentageBase === "salary"
          ? lineTotal(nextLine)
          : sheetDraft.percentageBase === "debt"
            ? line.debtBalance ?? 0
            : line.percentageBase
        return sheetDraft.percentageEnabled
          ? { ...nextLine, percentageDeduction: calculatePercentage(sheetDraft, nextLine, percentageBase) }
          : nextLine
      }),
    }
    if (updated.lines.some(line => line.advance + line.debtDeduction + (updated.percentageEnabled ? line.percentageDeduction : 0) > lineTotal(line))) {
      setError("تعديل عدد أيام الكشف يجعل بعض المقبوضات أو الخصومات أكبر من إجمالي راتب الموظف؛ عدّل المبالغ أولاً.")
      endBusy()
      return
    }
    try {
      if (await saveData({
        ...dataRef.current,
        sheets: dataRef.current.sheets.map(sheet => sheet.id === updated.id ? updated : sheet),
      }, "تم تعديل بيانات كشف الرواتب.")) setSheetDraft(null)
    } finally {
      endBusy()
    }
  }

  const deleteSheet = async (sheet: PayrollSheet) => {
    if (sheet.locked) {
      setError("هذا الكشف مثبت. فك التثبيت أولاً لحذفه.")
      return
    }
    if (!window.confirm(`حذف كشف الرواتب «${sheet.name}» وجميع بياناته؟`)) return
    if (!startBusy()) return
    try {
      const next = { ...dataRef.current, sheets: dataRef.current.sheets.filter(item => item.id !== sheet.id) }
      if (await saveData(next, "تم حذف كشف الرواتب.")) {
        if (expandedSheetId === sheet.id) setExpandedSheetId(null)
      }
    } finally {
      endBusy()
    }
  }

  const toggleSheetLock = async (sheet: PayrollSheet) => {
    const currentSheet = dataRef.current.sheets.find(item => item.id === sheet.id)
    if (!currentSheet) {
      setError("كشف الرواتب المطلوب تثبيته غير موجود.")
      return
    }
    if (currentSheet.locked) {
      setUnlockDraft({ sheetId: currentSheet.id, password: "" })
      setUnlockError("")
      setError("")
      return
    }

    for (const line of currentSheet.lines) {
      const key = `${currentSheet.id}:${line.employeeId}`
      const draft = lineDrafts[key]
      if (!draft) continue
      const dailyRate = evaluateArithmeticExpression(draft.dailyRate)
      if (dailyRate === null || dailyRate < 0) {
        setError(`أدخل أجراً يومياً صحيحاً للموظف ${line.name} قبل تثبيت الكشف.`)
        return
      }
      const updates = currentSheet.percentageEnabled
        ? {
          dailyRate,
          attendanceDays: Number(normalizeNumber(draft.attendanceDays)),
          bonus: Number(normalizeNumber(draft.bonus || "0")),
          ...(currentSheet.percentageBase === "manual"
            ? { percentageBase: Number(normalizeNumber(draft.percentageBase || "0")) }
            : {}),
        }
        : {
          dailyRate,
          advance: Number(normalizeNumber(draft.advance || "0")),
          bonus: Number(normalizeNumber(draft.bonus || "0")),
          attendanceDays: Number(normalizeNumber(draft.attendanceDays)),
        }
      if (!await updateLine(currentSheet.id, line.employeeId, updates)) return
    }

    if (!startBusy()) return
    try {
      const latestSheet = dataRef.current.sheets.find(item => item.id === currentSheet.id)
      if (!latestSheet || latestSheet.locked) {
        setError("تعذر تثبيت الكشف لأنه لم يعد متاحاً أو أنه مثبت مسبقاً.")
        return
      }
      const updated = { ...latestSheet, locked: true }
      if (await saveData({
        ...dataRef.current,
        sheets: dataRef.current.sheets.map(item => item.id === updated.id ? updated : item),
      }, "تم تثبيت الكشف. لا يمكن تعديل بياناته قبل فك التثبيت.")) {
        setSheetDraft(null)
        setPercentageDraft(null)
      }
    } finally {
      endBusy()
    }
  }

  const unlockSheet = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!unlockDraft) return
    if (!normalizedUsername || normalizeUsername(unlockDraft.password) !== normalizedUsername) {
      setUnlockError(normalizedUsername ? "كلمة السر غير صحيحة. أدخل اسم المستخدم الحالي." : "تعذر التحقق: اسم المستخدم غير متوفر.")
      setUnlockDraft(current => current ? { ...current, password: "" } : null)
      return
    }
    if (!startBusy()) return
    try {
      const sheet = dataRef.current.sheets.find(item => item.id === unlockDraft.sheetId)
      if (!sheet || !sheet.locked) {
        setError("الكشف غير موجود أو أنه غير مثبت.")
        return
      }
      const updated = { ...sheet, locked: false }
      if (await saveData({
        ...dataRef.current,
        sheets: dataRef.current.sheets.map(item => item.id === updated.id ? updated : item),
      }, "تم فك تثبيت الكشف. يمكنك الآن تعديل بياناته.")) {
        setUnlockError("")
        setUnlockDraft(null)
      } else {
        setUnlockError("تعذر حفظ فك التثبيت. حاول مرة أخرى.")
      }
    } finally {
      endBusy()
    }
  }

  const updateLine = async (sheetId: string, employeeId: string, updates: Partial<Pick<PayrollLine, "dailyRate" | "advance" | "bonus" | "percentageBase" | "attendanceDays" | "debtCollectionEnabled">>): Promise<boolean> => {
    if (!startBusy()) return false
    const sheet = dataRef.current.sheets.find(item => item.id === sheetId)
    const line = sheet?.lines.find(item => item.employeeId === employeeId)
    if (!sheet || !line) {
      setError("تعذر العثور على سطر راتب الموظف.")
      endBusy()
      return false
    }
    if (sheet.locked) {
      setError("هذا الكشف مثبت. فك التثبيت أولاً لتعديل بياناته.")
      endBusy()
      return false
    }
    if (sheet.percentageEnabled && "advance" in updates) {
      setError("لا يمكن تعديل المقبوض أو الخصم يدويًا أثناء تفعيل النسبة. أوقف النسبة أولاً.")
      endBusy()
      return false
    }
    const nextLine = { ...line, ...updates }
    if ("debtCollectionEnabled" in updates) {
      if (nextLine.debtCollectionEnabled === false) {
        nextLine.debtDeduction = 0
        if (sheet.percentageEnabled && sheet.percentageBase === "debt") nextLine.percentageDeduction = 0
      } else {
        nextLine.debtDeduction = sheet.percentageEnabled
          ? 0
          : calculateDebtDeduction(lineTotal(nextLine), nextLine.advance, nextLine.debtBalance)
      }
    }
    if ("dailyRate" in updates || "attendanceDays" in updates) {
      const attendanceDays = Number(nextLine.attendanceDays ?? getAttendanceDays(line, sheet))
      const dailyRate = Number(nextLine.dailyRate)
      if (!Number.isFinite(dailyRate) || dailyRate < 0) {
        setError(`أدخل أجراً يومياً صالحاً وغير سالب للموظف ${line.name}.`)
        endBusy()
        return false
      }
      nextLine.dailyRate = dailyRate
      if (!Number.isFinite(attendanceDays) || attendanceDays < 0 || attendanceDays > sheet.days) {
        setError(`أيام حضور ${line.name} يجب أن تكون بين صفر و${sheet.days}.`)
        endBusy()
        return false
      }
      nextLine.attendanceDays = attendanceDays
      nextLine.gross = amount(dailyRate * attendanceDays)
    } else {
      nextLine.attendanceDays = getAttendanceDays(line, sheet)
    }
    nextLine.bonus = Number(nextLine.bonus ?? 0)
    for (const key of ["advance", "bonus", "percentageBase"] as const) {
      const value = Number(nextLine[key])
      if (!Number.isFinite(value) || value < 0) {
        setError("أدخل مبلغاً صالحاً غير سالب.")
        endBusy()
        return false
      }
      nextLine[key] = amount(value)
    }
    if (!sheet.percentageEnabled) {
      nextLine.debtDeduction = nextLine.debtCollectionEnabled === false
        ? 0
        : calculateDebtDeduction(lineTotal(nextLine), nextLine.advance, nextLine.debtBalance)
    }
    if (sheet.percentageEnabled && sheet.percentageBase === "manual") {
      nextLine.percentageDeduction = calculatePercentage(sheet, nextLine, nextLine.percentageBase)
    }
    if (nextLine.advance + nextLine.debtDeduction + (sheet.percentageEnabled ? nextLine.percentageDeduction : 0) > lineTotal(nextLine)) {
      setError("لا يمكن أن يتجاوز مجموع المقبوض وخصم الدين إجمالي الراتب.")
      endBusy()
      return false
    }
    if (nextLine.debtBalance !== null && nextLine.debtDeduction > nextLine.debtBalance) {
      setError("لا يمكن أن يتجاوز خصم الدين رصيد الدين المسجل للموظف.")
      endBusy()
      return false
    }
    if (sheet.percentageEnabled) {
      const percentageBase = sheet.percentageBase === "salary"
        ? lineTotal(nextLine)
        : sheet.percentageBase === "debt"
          ? nextLine.debtBalance ?? 0
          : nextLine.percentageBase
      nextLine.percentageDeduction = sheet.percentageBase === "debt" && nextLine.debtCollectionEnabled === false
        ? 0
        : calculatePercentage(sheet, nextLine, percentageBase)
    }
    try {
      const updated = { ...sheet, lines: sheet.lines.map(item => item.employeeId === employeeId ? nextLine : item) }
      const saved = await saveData({
        ...dataRef.current,
        sheets: dataRef.current.sheets.map(item => item.id === sheetId ? updated : item),
      }, "تم حفظ التعديل.")
      if (saved) {
        const key = `${sheetId}:${employeeId}`
        setLineDrafts(current => {
          const next = { ...current }
          delete next[key]
          return next
        })
      }
      return saved
    } finally {
      endBusy()
    }
  }

  const saveLineDraft = (sheet: PayrollSheet, line: PayrollLine) => {
    const key = `${sheet.id}:${line.employeeId}`
    const draft = lineDrafts[key]
    if (!draft || sheet.percentageEnabled) return
    if (!draft.attendanceDays.trim()) {
      setError(`أدخل عدد أيام حضور ${line.name}.`)
      return
    }

    const dailyRate = evaluateArithmeticExpression(draft.dailyRate)
    if (dailyRate === null || dailyRate < 0) {
      setError(`أدخل أجراً يومياً صحيحاً للموظف ${line.name}.`)
      return
    }
    void updateLine(sheet.id, line.employeeId, {
      dailyRate,
      advance: Number(normalizeNumber(draft.advance || "0")),
      bonus: Number(normalizeNumber(draft.bonus || "0")),
      attendanceDays: Number(normalizeNumber(draft.attendanceDays)),
    })
  }

  const saveEmployeePayrollRate = async () => {
    if (!employeePayrollDialog) return
    const dailyRate = evaluateArithmeticExpression(employeePayrollDialog.dailyRate)
    if (dailyRate === null || dailyRate < 0) {
      setError("أدخل أجراً يومياً صالحاً وغير سالب.")
      return
    }
    if (await updateLine(employeePayrollDialog.sheetId, employeePayrollDialog.employeeId, { dailyRate })) {
      setEmployeePayrollDialog(current => current ? { ...current, dailyRate: String(dailyRate) } : null)
    }
  }

  const savePercentageBaseDraft = (sheet: PayrollSheet, line: PayrollLine) => {
    const key = `${sheet.id}:${line.employeeId}`
    const draft = lineDrafts[key]
    if (!draft || !sheet.percentageEnabled) return
    if (!draft.attendanceDays.trim()) {
      setError(`أدخل عدد أيام حضور ${line.name}.`)
      return
    }
    const dailyRate = evaluateArithmeticExpression(draft.dailyRate)
    if (dailyRate === null || dailyRate < 0) {
      setError(`أدخل أجراً يومياً صحيحاً للموظف ${line.name}.`)
      return
    }
    void updateLine(sheet.id, line.employeeId, {
      dailyRate,
      attendanceDays: Number(normalizeNumber(draft.attendanceDays)),
      bonus: Number(normalizeNumber(draft.bonus || "0")),
      ...(sheet.percentageBase === "manual" ? { percentageBase: Number(draft.percentageBase) || 0 } : {}),
    })
  }
  const openPercentage = (sheet: PayrollSheet) => {
    setPercentageDraft({ sheetId: sheet.id, percentage: String(sheet.percentage || 10), base: sheet.percentageBase })
  }

  const applyPercentage = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!percentageDraft || !startBusy()) return
    const percentage = Number(normalizeNumber(percentageDraft.percentage))
    const sheet = dataRef.current.sheets.find(item => item.id === percentageDraft.sheetId)
    if (!sheet || sheet.locked) {
      setError("الكشف غير موجود أو مثبت. فك التثبيت أولاً لتعديله.")
      endBusy()
      return
    }
    if (!Number.isFinite(percentage) || percentage <= 0 || percentage > 100) {
      setError("أدخل نسبة صحيحة أكبر من صفر ولا تزيد على 100%.")
      endBusy()
      return
    }
    const hasManualAmounts = sheet.lines.some(line => line.advance > 0 || line.debtDeduction > 0)
    const apply = async () => {
      if (!startBusy()) return
      const updated: PayrollSheet = {
        ...sheet,
        percentageEnabled: true,
        percentage,
        percentageBase: percentageDraft.base,
        lines: sheet.lines.map(line => {
          const base = percentageDraft.base === "salary"
            ? lineTotal(line)
            : percentageDraft.base === "debt"
              ? line.debtBalance ?? 0
              : line.percentageBase
          return {
            ...line,
            advance: 0,
            debtDeduction: 0,
            percentageBase: amount(base),
            percentageDeduction: percentageDraft.base === "debt" && line.debtCollectionEnabled === false
              ? 0
              : calculatePercentage({ ...sheet, percentage, percentageBase: percentageDraft.base }, line, amount(base)),
          }
        }),
      }
      try {
        if (await saveData({
          ...dataRef.current,
          sheets: dataRef.current.sheets.map(item => item.id === updated.id ? updated : item),
        }, `تم تطبيق خصم ${formatAmount(percentage)}% على كشف الرواتب.`)) setPercentageDraft(null)
      } finally {
        endBusy()
      }
    }
    if (hasManualAmounts) {
      endBusy()
      setOverwritePrompt(() => () => { setOverwritePrompt(null); void apply() })
    } else {
      endBusy()
      await apply()
    }
  }

  const disablePercentage = async (sheet: PayrollSheet) => {
    if (sheet.locked) {
      setError("هذا الكشف مثبت. فك التثبيت أولاً لتعديل خصوماته.")
      return
    }
    if (!window.confirm("سيتم إزالة الخصم النسبي المحسوب، ولن تُستعاد مبالغ المقبوض أو خصم الدين السابقة. هل تريد المتابعة؟")) return
    if (!startBusy()) return
    const updated = {
      ...sheet,
      percentageEnabled: false,
      percentage: 0,
      lines: sheet.lines.map(line => ({
        ...line,
        percentageDeduction: 0,
        percentageBase: 0,
        debtDeduction: line.debtCollectionEnabled === false
          ? 0
          : calculateDebtDeduction(lineTotal(line), line.advance, line.debtBalance),
      })),
    }
    try {
      await saveData({
        ...dataRef.current,
        sheets: dataRef.current.sheets.map(item => item.id === sheet.id ? updated : item),
      }, "تم إيقاف الخصم النسبي وإعادة احتساب تحصيل الدين تلقائياً. أدخل المقبوضات من جديد.")
    } finally {
      endBusy()
    }
  }

  const totals = (sheet: PayrollSheet) => {
    const gross = sheet.lines.reduce((sum, line) => sum + lineTotal(line), 0)
    const advances = sheet.lines.reduce((sum, line) => sum + line.advance, 0)
    const debtCollected = sheet.lines.reduce((sum, line) =>
      sum + line.debtDeduction + (sheet.percentageEnabled && sheet.percentageBase === "debt" ? line.percentageDeduction : 0), 0)
    const percentage = sheet.lines.reduce((sum, line) => sum + line.percentageDeduction, 0)
    return {
      gross: amount(gross),
      advances: amount(advances),
      debtCollected: amount(debtCollected),
      percentage: amount(percentage),
      receivedTotal: amount(advances + debtCollected + (sheet.percentageBase === "debt" ? 0 : percentage)),
      remaining: amount(gross - advances - debtCollected - (sheet.percentageEnabled && sheet.percentageBase !== "debt" ? percentage : 0)),
    }
  }

  const prepareSheetForExport = async (sheet: PayrollSheet): Promise<PayrollSheet | null> => {
    if (busyRef.current) {
      setError("انتظر انتهاء العملية الحالية قبل تصدير كشف الرواتب.")
      return null
    }
    for (const line of sheet.lines) {
      const draft = lineDrafts[`${sheet.id}:${line.employeeId}`]
      if (!draft) continue
      if (!draft.attendanceDays.trim()) {
        setError(`أدخل عدد أيام حضور ${line.name} قبل تصدير الكشف.`)
        return null
      }
      const dailyRate = evaluateArithmeticExpression(draft.dailyRate)
      if (dailyRate === null || dailyRate < 0) {
        setError(`أدخل أجراً يومياً صحيحاً للموظف ${line.name} قبل تصدير الكشف.`)
        return null
      }
      const updates = sheet.percentageEnabled
        ? {
          dailyRate,
          attendanceDays: Number(normalizeNumber(draft.attendanceDays)),
          ...(sheet.percentageBase === "manual"
            ? { percentageBase: Number(normalizeNumber(draft.percentageBase || "0")) }
            : {}),
        }
        : {
          dailyRate,
          advance: Number(normalizeNumber(draft.advance || "0")),
          attendanceDays: Number(normalizeNumber(draft.attendanceDays)),
        }
      if (!await updateLine(sheet.id, line.employeeId, updates)) return null
    }
    const currentSheet = dataRef.current.sheets.find(item => item.id === sheet.id)
    if (!currentSheet) {
      setError("تعذر العثور على كشف الرواتب المطلوب تصديره.")
      return null
    }
    return currentSheet
  }

  const exportSheetExcel = async (sheet: PayrollSheet) => {
    const currentSheet = await prepareSheetForExport(sheet)
    if (!currentSheet) return
    try {
      const headers = [
        "رقم الموظف", "اسم الموظف", "رقم الجوال", "الأجر اليومي", "أيام الحضور",
        "الراتب الأساسي", "البونص", "إجمالي المستحق", "رقم الزبون", "رصيد الدين السابق",
        "المقبوض مقدماً", "المحصل من الدين", "رصيد الدين المتبقي",
        ...(currentSheet.percentageEnabled && currentSheet.percentageBase === "manual" ? ["أساس الخصم"] : []),
        ...(currentSheet.percentageEnabled ? ["الخصم النسبي"] : []),
        "الصافي المستحق",
      ]
      const rows = currentSheet.lines.map(line => {
        const debtPercentage = currentSheet.percentageEnabled && currentSheet.percentageBase === "debt" ? line.percentageDeduction : 0
        const salaryPercentage = currentSheet.percentageEnabled && currentSheet.percentageBase === "salary" ? line.percentageDeduction : 0
        const manualPercentage = currentSheet.percentageEnabled && currentSheet.percentageBase === "manual" ? line.percentageDeduction : 0
        const debtCollected = line.debtCollectionEnabled === false ? 0 : line.debtDeduction + debtPercentage
        const net = amount(lineTotal(line) - line.advance - debtCollected - salaryPercentage - manualPercentage)
        return [
          line.employeeNumber,
          line.name,
          line.phone ?? "",
          line.dailyRate,
          getAttendanceDays(line, currentSheet),
          line.gross,
          lineBonus(line),
          lineTotal(line),
          line.customerNumber || "",
          line.debtBalance ?? "",
          line.advance,
          debtCollected,
          line.debtBalance === null ? "" : amount(Math.max(0, line.debtBalance - debtCollected)),
          ...(currentSheet.percentageEnabled && currentSheet.percentageBase === "manual" ? [line.percentageBase] : []),
          ...(currentSheet.percentageEnabled ? [line.percentageDeduction] : []),
          net,
        ]
      })
      const summary = totals(currentSheet)
      const totalRow: Array<string | number> = [
        "", "الإجمالي", "", "",
        currentSheet.lines.reduce((sum, line) => sum + getAttendanceDays(line, currentSheet), 0),
        currentSheet.lines.reduce((sum, line) => sum + line.gross, 0),
        currentSheet.lines.reduce((sum, line) => sum + lineBonus(line), 0),
        summary.gross,
        "", "",
        summary.advances,
        summary.debtCollected,
        "",
        ...(currentSheet.percentageEnabled && currentSheet.percentageBase === "manual"
          ? [currentSheet.lines.reduce((sum, line) => sum + line.percentageBase, 0)]
          : []),
        ...(currentSheet.percentageEnabled ? [summary.percentage] : []),
        summary.remaining,
      ]
      const worksheet = XLSX.utils.aoa_to_sheet([
        ["كشف رواتب الموظفين"],
        ["اسم الكشف", currentSheet.name],
        ["الفترة", `${currentSheet.periodFrom} - ${currentSheet.periodTo}`],
        ["موعد الصرف", currentSheet.paymentDate],
        [],
        headers,
        ...rows,
        [],
        totalRow,
      ])
      worksheet["!cols"] = headers.map(header => ({ wch: Math.max(14, Math.min(28, header.length + 4)) }))
      const workbook = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(workbook, worksheet, "كشف الرواتب")
      const safeName = currentSheet.name.trim().replace(/[\\/:*?"<>|]+/g, "-") || "كشف-رواتب"
      XLSX.writeFile(workbook, `${safeName}.xlsx`)
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر تصدير كشف Excel: ${cause.message}` : "تعذر تصدير كشف Excel.")
    }
  }

  const printSheet = (sheet: PayrollSheet) => {
    if (busyRef.current) {
      setError("انتظر انتهاء العملية الحالية قبل طباعة كشف الرواتب.")
      return
    }
    const printWindow = window.open("", "_blank")
    if (!printWindow) {
      setError("تعذر فتح نافذة الطباعة. اسمح بالنوافذ المنبثقة ثم أعد المحاولة.")
      return
    }
    void (async () => {
      try {
        const currentSheet = await prepareSheetForExport(sheet)
        if (!currentSheet) {
          printWindow.close()
          return
        }
        const headers = [
          "اسم الموظف", "الأجر اليومي", "أيام الحضور", "الراتب الأساسي", "بونص",
          "إجمالي المستحق", "رقم الزبون", "رصيد الدين", "المقبوض مقدماً",
          "المحصل من الدين", ...(currentSheet.percentageEnabled && currentSheet.percentageBase === "manual" ? ["أساس الخصم"] : []),
          ...(currentSheet.percentageEnabled ? ["الخصم النسبي"] : []), "الصافي",
        ]
        const rows = currentSheet.lines.map(line => {
          const debtPercentage = currentSheet.percentageEnabled && currentSheet.percentageBase === "debt" ? line.percentageDeduction : 0
          const salaryPercentage = currentSheet.percentageEnabled && currentSheet.percentageBase === "salary" ? line.percentageDeduction : 0
          const manualPercentage = currentSheet.percentageEnabled && currentSheet.percentageBase === "manual" ? line.percentageDeduction : 0
          const debtCollected = line.debtCollectionEnabled === false ? 0 : line.debtDeduction + debtPercentage
          const net = amount(lineTotal(line) - line.advance - debtCollected - salaryPercentage - manualPercentage)
          return [
            line.name,
            formatDailyRate(line.dailyRate),
            String(getAttendanceDays(line, currentSheet)),
            formatAmount(line.gross),
            formatAmount(lineBonus(line)),
            formatAmount(lineTotal(line)),
            line.customerNumber || "—",
            line.debtBalance === null ? "غير مرتبط" : formatAmount(line.debtBalance),
            formatAmount(line.advance),
            formatAmount(debtCollected),
            ...(currentSheet.percentageEnabled && currentSheet.percentageBase === "manual" ? [formatAmount(line.percentageBase)] : []),
            ...(currentSheet.percentageEnabled ? [formatAmount(line.percentageDeduction)] : []),
            formatAmount(net),
          ]
        })
        const summary = totals(currentSheet)
        const totalRow = [
          "الإجمالي",
          "—",
          String(currentSheet.lines.reduce((sum, line) => sum + getAttendanceDays(line, currentSheet), 0)),
          formatAmount(currentSheet.lines.reduce((sum, line) => sum + line.gross, 0)),
          formatAmount(currentSheet.lines.reduce((sum, line) => sum + lineBonus(line), 0)),
          formatAmount(summary.gross),
          "—",
          "—",
          formatAmount(summary.advances),
          formatAmount(summary.debtCollected),
          ...(currentSheet.percentageEnabled && currentSheet.percentageBase === "manual"
            ? [formatAmount(currentSheet.lines.reduce((sum, line) => sum + line.percentageBase, 0))]
            : []),
          ...(currentSheet.percentageEnabled ? [formatAmount(summary.percentage)] : []),
          formatAmount(summary.remaining),
        ]
        const tableRows = [rows, [totalRow]].flat()
          .map((row, index) => `<tr${index === rows.length ? ' class="total"' : ""}>${row.map(value => `<td>${escapeHtml(value)}</td>`).join("")}</tr>`)
          .join("")
        const safeTitle = escapeHtml(currentSheet.name)
        const html = `<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>${safeTitle}</title><style>${REPORT_LOGO_PRINT_STYLES}
          @page{size:A4 landscape;margin:10mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;color:#10243c;font-size:10px;margin:0}
          h1{text-align:center;font-size:20px;margin:0 0 8px}p{text-align:center;margin:4px 0 14px}
          table{width:100%;border-collapse:collapse;table-layout:auto}th,td{border:1px solid #78909c;padding:5px 4px;text-align:right;vertical-align:middle;overflow-wrap:anywhere}
          th{background:#e2eef3;font-weight:bold}.total td{background:#dceff7;font-weight:bold}thead{display:table-header-group}tr{break-inside:avoid;page-break-inside:avoid}
          </style></head><body>${reportLogoMarkup()}<h1>تقرير رواتب الموظفين</h1><p>${safeTitle}</p>
          <p>الفترة: ${escapeHtml(formatDate(currentSheet.periodFrom))} – ${escapeHtml(formatDate(currentSheet.periodTo))} · موعد الصرف: ${escapeHtml(formatDate(currentSheet.paymentDate))}</p>
          <table><thead><tr>${headers.map(header => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${tableRows}</tbody></table>
          </body></html>`
        printWindow.document.open()
        printWindow.document.write(html)
        printWindow.document.close()
        let printStarted = false
        const startPrinting = () => {
          if (printStarted || printWindow.closed) return
          printStarted = true
          printWindow.focus()
          printWindow.print()
        }
        const logo = printWindow.document.querySelector<HTMLImageElement>(".report-print-logo img")
        if (logo && !logo.complete) {
          logo.addEventListener("load", startPrinting, { once: true })
          logo.addEventListener("error", startPrinting, { once: true })
        } else {
          window.setTimeout(startPrinting, 250)
        }
      } catch (cause) {
        printWindow.close()
        setError(cause instanceof Error ? `تعذرت طباعة كشف الرواتب: ${cause.message}` : "تعذرت طباعة كشف الرواتب.")
      }
    })()
  }

  if (loadedKey !== payrollKey) {
    return <main dir="rtl" className="min-h-full bg-[#f4f7fb] p-8 text-center text-sm text-slate-600">{error || "جارٍ تحميل بيانات الرواتب..."}</main>
  }

  return (
    <main dir="rtl" className="payroll-page min-h-full bg-[#f4f7fb] px-4 py-7 text-slate-900 sm:px-7 lg:px-10">
      <div className="mx-auto max-w-7xl space-y-6">
        <header className="flex flex-wrap items-center justify-between gap-4 rounded-3xl bg-[#10243c] px-6 py-7 text-white shadow-xl sm:px-9">
          <div>
            <div className="inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/5 px-3 py-1.5 text-xs text-cyan-100"><Banknote className="h-4 w-4" />كشوف مستقلة لكل فترة</div>
            <h1 className="mt-3 text-2xl font-extrabold sm:text-3xl">كشف رواتب الموظفين</h1>
            <p className="mt-2 text-sm text-slate-300">راتب اليوم × عدد الأيام، مع المقبوضات ودفعات الديون وخصومات الفترة.</p>
          </div>
          <div className="flex flex-wrap gap-2">
            <button type="button" onClick={() => void refreshDebts()} className="inline-flex items-center gap-2 rounded-xl border border-white/20 px-4 py-2.5 text-sm font-semibold text-white hover:bg-white/10"><RotateCw className="h-4 w-4" />تحديث كشف الديون</button>
            <button type="button" onClick={() => setTab("employees")} className="inline-flex items-center gap-2 rounded-xl border border-white/20 px-4 py-2.5 text-sm font-semibold text-white hover:bg-white/10"><Users className="h-4 w-4" />إدارة الموظفين</button>
            <button type="button" onClick={() => { setSheetForm(emptySheetForm()); setCreatingSheet(true) }} disabled={!data.employees.length} className="inline-flex items-center gap-2 rounded-xl bg-cyan-500 px-4 py-2.5 text-sm font-bold text-white hover:bg-cyan-400 disabled:opacity-50"><FilePlus2 className="h-4 w-4" />كشف رواتب جديد</button>
          </div>
        </header>

        {error && <div role="alert" className="flex items-start justify-between gap-3 rounded-xl border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-800"><span>{error}</span><button type="button" onClick={() => setError("")} aria-label="إغلاق"><X className="h-4 w-4" /></button></div>}
        {notice && <div role="status" className="rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-800">{notice}</div>}
        {!debtDatasetReady && <div role="status" className="rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-900">جارٍ تحميل كشف ديون الزبائن المرتبط.</div>}

        <nav className="flex gap-2 rounded-2xl border bg-white p-2 shadow-sm">
          <button type="button" onClick={() => setTab("sheets")} className={`rounded-xl px-4 py-2.5 text-sm font-semibold ${tab === "sheets" ? "bg-[#10243c] text-white" : "text-slate-600 hover:bg-slate-100"}`}>كشوف الرواتب ({data.sheets.length})</button>
          <button type="button" onClick={() => setTab("employees")} className={`rounded-xl px-4 py-2.5 text-sm font-semibold ${tab === "employees" ? "bg-[#10243c] text-white" : "text-slate-600 hover:bg-slate-100"}`}>الموظفون ({data.employees.length})</button>
        </nav>

        {tab === "employees" && <section className="overflow-hidden rounded-2xl border bg-white shadow-sm">
          <header className="flex flex-wrap items-center justify-between gap-3 border-b p-4 sm:p-5"><div><h2 className="font-bold">سجل الموظفين</h2><p className="mt-1 text-xs text-slate-500">رقم الموظف مستقل عن رقم الزبون المستخدم لجلب رصيد الدين.</p></div><button type="button" onClick={() => setEmployeeDraft(emptyEmployee())} className="inline-flex items-center gap-2 rounded-xl bg-cyan-800 px-4 py-2.5 text-sm font-semibold text-white hover:bg-cyan-900"><Plus className="h-4 w-4" />إضافة موظف</button></header>
          {data.employees.length ? <div className="overflow-x-auto"><table className="w-full min-w-[900px] text-right text-xs"><thead className="bg-slate-50 text-slate-600"><tr>{["رقم الموظف", "الاسم", "الأجر اليومي", "رقم الجوال", "رقم الزبون للديون", "رصيد الدين الحالي", "الإجراءات"].map(label => <th key={label} className="px-4 py-3 font-semibold">{label}</th>)}</tr></thead><tbody className="divide-y">{data.employees.map(employee => <tr key={employee.id} className="hover:bg-slate-50/60"><td className="px-4 py-3 font-bold">{employee.employeeNumber}</td><td className="px-4 py-3 font-semibold">{employee.name}</td><td className="px-4 py-3">{formatDailyRate(employee.dailyRate)}</td><td className="px-4 py-3" dir="ltr">{employee.phone || "—"}</td><td className="px-4 py-3">{employee.customerNumber || "—"}</td><td className="px-4 py-3 font-semibold">{employee.customerNumber && debts.has(employee.customerNumber) ? formatAmount(debts.get(employee.customerNumber)!) : "لا يوجد رصيد مرتبط"}</td><td className="px-4 py-3"><div className="flex gap-2"><button type="button" onClick={() => setEmployeeDraft({ ...employee, phone: employee.phone ?? "", dailyRate: String(employee.dailyRate) })} title="تعديل" className="rounded-lg border p-2 hover:bg-slate-100"><Pencil className="h-4 w-4" /></button><button type="button" onClick={() => void deleteEmployee(employee)} disabled={busy} title="حذف" className="rounded-lg p-2 text-rose-700 hover:bg-rose-50 disabled:opacity-50"><Trash2 className="h-4 w-4" /></button></div></td></tr>)}</tbody></table></div> : <p className="p-12 text-center text-sm text-slate-500">أضف الموظفين أولاً لتتمكن من إنشاء كشف رواتب.</p>}
        </section>}

        {tab === "sheets" && (data.sheets.length ? <section className="space-y-4">{data.sheets.map(sheet => {
          const summary = totals(sheet)
          const expanded = expandedSheetId === sheet.id
          return <article key={sheet.id} className="overflow-hidden rounded-2xl border bg-white shadow-sm">
            <header className="flex flex-wrap items-center gap-3 p-4 sm:p-5">
              <button type="button" onClick={() => setExpandedSheetId(expanded ? null : sheet.id)} aria-expanded={expanded} className="flex min-w-0 flex-1 items-center justify-between gap-3 text-right">
                <span className="flex min-w-0 items-center gap-3"><CalendarDays className="h-5 w-5 shrink-0 text-cyan-800" /><span className="min-w-0"><span className="block truncate font-bold">{sheet.name}</span><span className="mt-1 block text-xs text-slate-500">{formatDate(sheet.periodFrom)} – {formatDate(sheet.periodTo)} · حتى {sheet.days} يومًا · موعد الصرف {formatDate(sheet.paymentDate)} · أُنشئ {new Date(sheet.createdAt).toLocaleString("ar", { dateStyle: "short", timeStyle: "short" })} · {sheet.lines.length} موظف</span></span></span><ChevronDown className={`h-5 w-5 shrink-0 transition ${expanded ? "rotate-180" : ""}`} />
              </button>
              <div className="flex flex-wrap gap-1.5"><span className="rounded-lg bg-slate-100 px-2.5 py-2 text-xs font-semibold">المستحق {formatAmount(summary.gross)}</span><span className="rounded-lg bg-emerald-50 px-2.5 py-2 text-xs font-semibold text-emerald-800">المقبوض {formatAmount(summary.advances)}</span><span className="rounded-lg bg-amber-50 px-2.5 py-2 text-xs font-semibold text-amber-800">تحصيل الدين {formatAmount(summary.debtCollected)}</span><span className="rounded-lg bg-cyan-50 px-2.5 py-2 text-xs font-semibold text-cyan-900">الصافي {formatAmount(summary.remaining)}</span></div>
              <div className="flex flex-wrap gap-1">
                <button type="button" onClick={() => setSheetDraft({ ...sheet })} disabled={busy || sheet.locked} title="تعديل اسم الكشف أو الفترة" className="rounded-lg border p-2 hover:bg-slate-100 disabled:cursor-not-allowed disabled:opacity-50"><Pencil className="h-4 w-4" /></button>
                <button type="button" onClick={() => void toggleSheetLock(sheet)} disabled={busy} title={sheet.locked ? "فك تثبيت الكشف" : "تثبيت الكشف"} className={`inline-flex items-center gap-1 rounded-lg border px-2.5 py-2 text-xs font-semibold disabled:cursor-not-allowed disabled:opacity-50 ${sheet.locked ? "border-amber-200 bg-amber-50 text-amber-900 hover:bg-amber-100" : "border-cyan-200 bg-cyan-50 text-cyan-900 hover:bg-cyan-100"}`}>
                  {sheet.locked ? <Unlock className="h-4 w-4" /> : <Lock className="h-4 w-4" />}{sheet.locked ? "فك التثبيت" : "تثبيت"}
                </button>
                <button type="button" onClick={() => void exportSheetExcel(sheet)} disabled={busy} title="تنزيل كشف الرواتب Excel" className="inline-flex items-center gap-1.5 rounded-lg border px-2.5 py-2 text-xs font-semibold hover:bg-slate-100 disabled:opacity-50"><FileSpreadsheet className="h-4 w-4 text-emerald-700" />Excel</button>
                <button type="button" onClick={() => void printSheet(sheet)} disabled={busy} title="طباعة أو حفظ كشف الرواتب PDF" className="inline-flex items-center gap-1.5 rounded-lg border px-2.5 py-2 text-xs font-semibold hover:bg-slate-100 disabled:opacity-50"><Printer className="h-4 w-4" />PDF</button>
                <button type="button" onClick={() => void deleteSheet(sheet)} disabled={busy || sheet.locked} title={sheet.locked ? "فك التثبيت قبل الحذف" : "حذف كشف الرواتب"} className="rounded-lg p-2 text-rose-700 hover:bg-rose-50 disabled:cursor-not-allowed disabled:opacity-50"><Trash2 className="h-4 w-4" /></button>
              </div>
            </header>
            {expanded && <div className="border-t">
              {sheet.locked && <div role="status" className="flex items-center gap-2 border-b bg-amber-50 px-4 py-3 text-xs font-semibold text-amber-900"><Lock className="h-4 w-4 shrink-0" />هذا الكشف مثبت وجميع بياناته للقراءة فقط. أدخل اسم المستخدم الحالي لفك التثبيت.</div>}
              <div className="grid gap-3 border-b bg-slate-50 p-4 text-xs sm:grid-cols-2 lg:grid-cols-5"><div>إجمالي الراتب المستحق: <strong>{formatAmount(summary.gross)}</strong></div><div>المقبوض مقدماً: <strong>{formatAmount(summary.advances)}</strong></div><div>المحصل من الديون: <strong>{formatAmount(summary.debtCollected)}</strong></div><div>إجمالي المقبوض والمحصل: <strong>{formatAmount(summary.receivedTotal)}</strong></div><div>الصافي المتبقي للصرف: <strong>{formatAmount(summary.remaining)}</strong></div></div>
              <p className="border-b bg-cyan-50 px-4 py-2.5 text-xs text-cyan-950">أدخل أيام الحضور والمقبوض لكل موظف. يُحسب المستحق من الأجر اليومي × أيام الحضور، ويُحسب تحصيل الدين تلقائياً من المستحق بعد طرح المقبوض، بحد أقصى رصيد الدين.</p>
              <div className="flex flex-wrap items-center justify-between gap-3 border-b p-4"><div><h3 className="text-sm font-bold">{sheet.percentageEnabled ? `خصم نسبة ${formatAmount(sheet.percentage)}%` : "المقبوض وتحصيل الدين التلقائي"}</h3><p className="mt-1 text-xs text-slate-500">{sheet.percentageEnabled ? `أساس النسبة: ${sheet.percentageBase === "salary" ? "إجمالي راتب الفترة" : sheet.percentageBase === "debt" ? "رصيد الدين" : "مبلغ الأساس اليدوي"}` : "تحصيل الدين = المستحق − المقبوض، بما لا يتجاوز رصيد الدين المرتبط."}</p></div>{sheet.percentageEnabled ? <button type="button" onClick={() => void disablePercentage(sheet)} disabled={busy || sheet.locked} className="rounded-lg border border-rose-200 px-3 py-2 text-xs font-semibold text-rose-800 hover:bg-rose-50 disabled:opacity-50">إيقاف النسبة</button> : <button type="button" onClick={() => openPercentage(sheet)} disabled={busy || sheet.locked} className="rounded-lg border border-cyan-200 bg-cyan-50 px-3 py-2 text-xs font-semibold text-cyan-900 hover:bg-cyan-100 disabled:opacity-50">تطبيق خصم نسبة</button>}</div>
              {sheet.percentageEnabled && <div className="border-b bg-cyan-50 px-4 py-3 text-xs text-cyan-950">قيمة النسبة الإجمالية: <strong>{formatAmount(summary.percentage)}</strong>{sheet.percentageBase === "debt" && <span> · تُحسب ضمن المحصل من الديون</span>}</div>}
              <div className="overflow-x-auto"><table className="w-full min-w-[1180px] text-right text-xs"><thead className="bg-slate-50 text-slate-600"><tr>{["رقم الموظف", "اسم الموظف", "الأجر اليومي", "أيام الحضور", "الراتب الأساسي", "بونص", "إجمالي المستحق", "رقم الزبون", "رصيد الدين", "المقبوض مقدماً", "المحصل من الدين", ...(sheet.percentageEnabled && sheet.percentageBase === "manual" ? ["أساس الخصم"] : []), ...(sheet.percentageEnabled ? ["الخصم النسبي"] : []), "الصافي"].map(label => <th key={label} className="px-3 py-3 font-semibold">{label}</th>)}</tr></thead><tbody className="divide-y">{sheet.lines.map((line, lineIndex) => {
                const lineKey = `${sheet.id}:${line.employeeId}`
                const draft = lineDrafts[lineKey] ?? {
                  dailyRate: String(line.dailyRate),
                  advance: String(line.advance ?? 0),
                  bonus: String(lineBonus(line)),
                  percentageBase: String(line.percentageBase ?? 0),
                  attendanceDays: String(getAttendanceDays(line, sheet)),
                }
                const debtPercentage = sheet.percentageEnabled && sheet.percentageBase === "debt" ? line.percentageDeduction : 0
                const salaryPercentage = sheet.percentageEnabled && sheet.percentageBase === "salary" ? line.percentageDeduction : 0
                const manualPercentage = sheet.percentageEnabled && sheet.percentageBase === "manual" ? line.percentageDeduction : 0
                const debtTotal = line.debtDeduction + debtPercentage
                const net = amount(lineTotal(line) - line.advance - debtTotal - salaryPercentage - manualPercentage)
                return <tr key={line.employeeId} onBlur={event => {
                  if (event.relatedTarget instanceof Element && event.relatedTarget.closest("[data-payroll-employee-dialog]")) return
                  if (event.currentTarget.contains(event.relatedTarget as Node | null)) return
                  window.setTimeout(() => {
                    const latestSheet = dataRef.current.sheets.find(item => item.id === sheet.id)
                    if (latestSheet?.locked) return
                    if (sheet.percentageEnabled) savePercentageBaseDraft(sheet, line)
                    else saveLineDraft(sheet, line)
                  }, 0)
                }} className="hover:bg-slate-50/50">
                  <td className="px-3 py-3 font-bold">{line.employeeNumber}</td>
                  <td className="px-3 py-3 font-semibold"><button type="button" onClick={() => {
                    if (!sheet.locked && lineDrafts[lineKey]) {
                      if (sheet.percentageEnabled) savePercentageBaseDraft(sheet, line)
                      else saveLineDraft(sheet, line)
                    }
                    setEmployeePayrollDialog({ sheetId: sheet.id, employeeId: line.employeeId, dailyRate: draft.dailyRate })
                  }} className="text-cyan-900 underline decoration-dotted underline-offset-2 hover:text-cyan-700">{line.name}</button></td>
                  <td className="px-3 py-3">
                    {sheet.locked
                      ? <span className="inline-flex min-w-24 justify-center rounded-lg border bg-slate-100 px-2 py-1.5 text-slate-600">{formatDailyRate(line.dailyRate)}</span>
                      : <input
                        aria-label={`الأجر اليومي ${line.name}`}
                        title="يمكن إدخال عملية مثل 490+10"
                        data-payroll-column="dailyRate"
                        data-payroll-row={lineIndex}
                        onKeyDown={event => handlePayrollGridKeyDown(event, "dailyRate", lineIndex)}
                        type="text"
                        inputMode="decimal"
                        value={draft.dailyRate}
                        onChange={event => setLineDrafts(current => ({
                          ...current,
                          [lineKey]: { ...draft, dailyRate: event.target.value },
                        }))}
                        className="w-28 rounded-lg border px-2 py-1.5"
                      />}
                  </td>
                  <td className="px-3 py-3">{sheet.locked ? <span className="inline-flex min-w-20 justify-center rounded-lg border bg-slate-100 px-2 py-1.5 text-slate-600">{getAttendanceDays(line, sheet)}</span> : <input aria-label={`أيام حضور ${line.name}`} title={`من صفر إلى ${sheet.days} أيام ويمكن استخدام الكسور للاحتساب بالساعات`} data-payroll-column="attendanceDays" data-payroll-row={lineIndex} onKeyDown={event => handlePayrollGridKeyDown(event, "attendanceDays", lineIndex)} type="number" min="0" max={sheet.days} step="0.01" value={draft.attendanceDays} onChange={event => setLineDrafts(current => ({ ...current, [lineKey]: { ...draft, attendanceDays: event.target.value } }))} className="w-20 rounded-lg border px-2 py-1.5" />}</td>
                  <td className="px-3 py-3 font-bold">{formatAmount(line.gross)}</td>
                  <td className="px-3 py-3">{sheet.locked ? <span className="inline-flex min-w-24 justify-center rounded-lg border bg-slate-100 px-2 py-1.5 text-slate-600">{formatAmount(lineBonus(line))}</span> : <input aria-label={`بونص ${line.name}`} data-payroll-column="bonus" data-payroll-row={lineIndex} onKeyDown={event => handlePayrollGridKeyDown(event, "bonus", lineIndex)} type="number" min="0" step="0.01" value={draft.bonus} onChange={event => setLineDrafts(current => ({ ...current, [lineKey]: { ...draft, bonus: event.target.value } }))} className="w-24 rounded-lg border px-2 py-1.5" />}</td>
                  <td className="px-3 py-3 font-bold">{formatAmount(lineTotal(line))}</td>
                  <td className="px-3 py-3">{line.customerNumber || "—"}</td>
                  <td className="px-3 py-3">{line.debtBalance === null ? "غير مرتبط" : formatAmount(line.debtBalance)}</td>
                  <td className="px-3 py-3">{sheet.percentageEnabled || sheet.locked ? <span className="inline-flex min-w-28 justify-center rounded-lg border bg-slate-100 px-2 py-1.5 text-slate-600">{formatAmount(line.advance)}</span> : <input aria-label={`المقبوض من ${line.name}`} data-payroll-column="advance" data-payroll-row={lineIndex} onKeyDown={event => handlePayrollGridKeyDown(event, "advance", lineIndex)} type="number" min="0" max={lineTotal(line)} step="0.01" value={draft.advance} onChange={event => setLineDrafts(current => ({ ...current, [lineKey]: { ...draft, advance: event.target.value } }))} className="w-28 rounded-lg border px-2 py-1.5" />}</td>
                  <td className="px-3 py-3"><span className="inline-flex min-w-28 justify-center rounded-lg border bg-slate-100 px-2 py-1.5 text-slate-600">{formatAmount(line.debtDeduction + debtPercentage)}</span></td>
                  {sheet.percentageEnabled && sheet.percentageBase === "manual" && <td className="px-3 py-3">{sheet.locked ? <span className="inline-flex min-w-28 justify-center rounded-lg border bg-slate-100 px-2 py-1.5 text-slate-600">{formatAmount(line.percentageBase)}</span> : <input aria-label={`أساس النسبة من ${line.name}`} data-payroll-column="percentageBase" data-payroll-row={lineIndex} onKeyDown={event => handlePayrollGridKeyDown(event, "percentageBase", lineIndex)} type="number" min="0" step="0.01" value={draft.percentageBase} onChange={event => setLineDrafts(current => ({ ...current, [lineKey]: { ...draft, percentageBase: event.target.value } }))} className="w-28 rounded-lg border px-2 py-1.5" />}</td>}
                  {sheet.percentageEnabled && <td className="px-3 py-3 font-semibold">{formatAmount(line.percentageDeduction)}</td>}
                  <td className="px-3 py-3 font-extrabold text-cyan-900">{formatAmount(net)}</td>
                </tr>
              })}</tbody></table></div>
            </div>}
          </article>
        })}</section> : <section className="rounded-2xl border bg-white p-12 text-center shadow-sm"><CalendarDays className="mx-auto h-10 w-10 text-cyan-700" /><p className="mt-3 font-bold">لا توجد كشوف رواتب محفوظة بعد.</p><p className="mt-1 text-sm text-slate-500">{data.employees.length ? "أنشئ كشفاً جديداً وحدد الفترة وعدد الأيام." : "أضف الموظفين أولاً ثم أنشئ كشف الرواتب."}</p>{!data.employees.length && <button type="button" onClick={() => setTab("employees")} className="mt-4 rounded-xl bg-cyan-800 px-4 py-2.5 text-sm font-semibold text-white">إضافة الموظفين</button>}</section>)}
      </div>

      {employeePayrollDialog && (() => {
        const selectedSheet = data.sheets.find(item => item.id === employeePayrollDialog.sheetId)
        const selectedLine = selectedSheet?.lines.find(item => item.employeeId === employeePayrollDialog.employeeId)
        if (!selectedSheet || !selectedLine) return null
        const debtCollectionEnabled = selectedLine.debtCollectionEnabled !== false
        const debtCollected = debtCollectionEnabled
          ? selectedLine.debtDeduction
            + (selectedSheet.percentageEnabled && selectedSheet.percentageBase === "debt" ? selectedLine.percentageDeduction : 0)
          : 0
        const salaryPercentage = selectedSheet.percentageEnabled && selectedSheet.percentageBase === "salary" ? selectedLine.percentageDeduction : 0
        const manualPercentage = selectedSheet.percentageEnabled && selectedSheet.percentageBase === "manual" ? selectedLine.percentageDeduction : 0
        const netPay = amount(lineTotal(selectedLine) - selectedLine.advance - debtCollected - salaryPercentage - manualPercentage)
        const phone = data.employees.find(item => item.id === selectedLine.employeeId)?.phone || selectedLine.phone || ""
        const addedMessage = `مرحباً ${selectedLine.name}، تمت إضافة راتبك في كشف ${selectedSheet.name}. إجمالي الراتب ${formatAmount(lineTotal(selectedLine))}، وصافي المستحق بعد المقبوض وتسديد الدين ${formatAmount(netPay)}.`
        const previousDebtBalance = selectedLine.debtBalance
        const currentDebtBalance = previousDebtBalance === null
          ? null
          : amount(Math.max(0, previousDebtBalance - debtCollected))
        const receivedMessage = [
          `عزيزي الموظف: ${selectedLine.name}`,
          "",
          `نود إعلامك بتوثيق وتفاصيل مستحقاتك المالية عن كشف: ${selectedSheet.name}`,
          "",
          `* إجمالي المستحق: ${formatAmount(lineTotal(selectedLine))} شيكل`,
          `* المقبوض مقدماً: ${formatAmount(selectedLine.advance)} شيكل`,
          `* المحصل من الدين (المسدد): ${formatAmount(debtCollected)} شيكل`,
          `* رصيد الدين السابق: ${previousDebtBalance === null ? "غير محدد" : `${formatAmount(previousDebtBalance)} شيكل`}`,
          `* رصيد الدين الحالي المتبقي: ${currentDebtBalance === null ? "غير محدد" : `${formatAmount(currentDebtBalance)} شيكل`}`,
          `* يتبقى لك مبلغ: ${formatAmount(netPay)} شيكل`,
          "اذا كنت تجد خلل يرجى مراجعة المختص",
          "",
          "شاكرين جهودكم، ومع خالص تمنياتنا لكم بالتوفيق.",
        ].join("\n")
        const validPhone = whatsappPhone(phone).length >= 8 && whatsappPhone(phone).length <= 15
        const draftRate = evaluateArithmeticExpression(employeePayrollDialog.dailyRate)
        const hasUnsavedRate = draftRate === null || draftRate !== selectedLine.dailyRate
        return <div role="dialog" aria-modal="true" aria-labelledby="payroll-employee-action-title" data-payroll-employee-dialog className="fixed inset-0 z-[85] flex items-center justify-center bg-slate-950/50 p-4">
          <div className="w-full max-w-lg space-y-4 rounded-2xl bg-white p-5 shadow-2xl sm:p-7">
            <div className="flex items-start justify-between gap-3">
              <div><h2 id="payroll-employee-action-title" className="text-lg font-bold">{selectedLine.name}</h2><p className="mt-1 text-xs text-slate-500">تعديل راتب الموظف في هذا الكشف فقط أو تجهيز رسالة واتساب.</p></div>
              <button type="button" onClick={() => setEmployeePayrollDialog(null)} aria-label="إغلاق" className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5" /></button>
            </div>
            <label className="block space-y-1.5 text-xs font-semibold">الأجر اليومي في هذا الكشف
              <input autoFocus type="text" inputMode="decimal" disabled={selectedSheet.locked} value={employeePayrollDialog.dailyRate} onChange={event => setEmployeePayrollDialog({ ...employeePayrollDialog, dailyRate: event.target.value })} className={INPUT_CLASS} />
              <span className="block font-normal text-slate-500">تقبل الخانة عملية مثل 490+10. لا يتغير أجر الموظف في الكشوف الأخرى.</span>
            </label>
            <label className="flex items-center gap-2 rounded-xl border border-slate-200 p-3 text-xs font-semibold text-slate-700">
              <input
                type="checkbox"
                checked={debtCollectionEnabled}
                disabled={selectedSheet.locked || busy}
                onChange={event => void updateLine(selectedSheet.id, selectedLine.employeeId, { debtCollectionEnabled: event.target.checked })}
                className="h-4 w-4 accent-cyan-800"
              />
              تحصيل الدين من راتب هذا الموظف
              <span className="font-normal text-slate-500">(أزل التحديد لإبقاء المبلغ مستحقاً له)</span>
            </label>
            <div className="grid gap-2 rounded-xl bg-cyan-50 p-3 text-xs text-cyan-950 sm:grid-cols-2">
              <p>إجمالي الراتب: <strong>{formatAmount(lineTotal(selectedLine))}</strong></p>
              <p>الصافي المستحق: <strong>{formatAmount(netPay)}</strong></p>
              <p>المقبوض مقدماً: <strong>{formatAmount(selectedLine.advance)}</strong></p>
              <p>المسدد من الدين: <strong>{formatAmount(debtCollected)}</strong></p>
              <p>المتبقي للموظف: <strong>{formatAmount(netPay)}</strong></p>
            </div>
            {!validPhone && <p className="rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-900">أضف رقم جوال صالحاً للموظف من سجل الموظفين لتجهيز رسائل واتساب.</p>}
            {hasUnsavedRate && <p className="rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-900">احفظ الأجر اليومي الجديد أولاً لتحديث المبلغ في الرسالة.</p>}
            <div className="flex flex-wrap gap-2">
              <button type="button" disabled={busy || selectedSheet.locked} onClick={() => void saveEmployeePayrollRate()} className="rounded-xl bg-cyan-800 px-4 py-2.5 text-sm font-semibold text-white disabled:opacity-50">حفظ الأجر في هذا الكشف</button>
              <button type="button" disabled={busy || !validPhone || hasUnsavedRate} onClick={() => {
                if (!openEmployeeWhatsApp(phone, addedMessage)) setError("رقم جوال الموظف غير صالح لفتح واتساب.")
              }} className="inline-flex items-center gap-2 rounded-xl border border-emerald-300 px-3 py-2.5 text-sm font-semibold text-emerald-800 disabled:opacity-50"><MessageCircle className="h-4 w-4" />إرسال تفاصيل الراتب</button>
              <button type="button" disabled={busy || !validPhone || hasUnsavedRate} onClick={() => {
                if (!openEmployeeWhatsApp(phone, receivedMessage)) setError("رقم جوال الموظف غير صالح لفتح واتساب.")
              }} className="inline-flex items-center gap-2 rounded-xl border border-emerald-300 px-3 py-2.5 text-sm font-semibold text-emerald-800 disabled:opacity-50"><MessageCircle className="h-4 w-4" />تأكيد استلام الراتب</button>
            </div>
            <p className="text-[11px] text-slate-500">تفتح الأزرار رسالة جاهزة في واتساب؛ راجعها وأرسلها بنفسك. تأكيد الاستلام رسالة فقط ولا يغيّر قيمة المقبوض في الكشف.</p>
          </div>
        </div>
      })()}

      {employeeDraft && <div role="dialog" aria-modal="true" aria-labelledby="payroll-employee-title" className="fixed inset-0 z-[80] flex items-start justify-center overflow-y-auto bg-slate-950/50 p-4 sm:items-center"><form onSubmit={event => void saveEmployee(event)} className="my-4 w-full max-w-xl space-y-4 rounded-2xl bg-white p-5 shadow-2xl sm:p-7">
        <div className="flex items-center justify-between"><div><h2 id="payroll-employee-title" className="text-lg font-bold">{employeeDraft.id ? "تعديل موظف" : "إضافة موظف"}</h2><p className="mt-1 text-xs text-slate-500">رقم الزبون اختياري، ويُستخدم فقط لعرض رصيد الدين.</p></div><button type="button" onClick={() => setEmployeeDraft(null)} aria-label="إغلاق" className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <div className="grid gap-3 sm:grid-cols-2"><label className="space-y-1.5 text-xs font-semibold">رقم الموظف<input required value={employeeDraft.employeeNumber} onChange={event => setEmployeeDraft({ ...employeeDraft, employeeNumber: event.target.value })} className={INPUT_CLASS} /></label><label className="space-y-1.5 text-xs font-semibold">اسم الموظف<input required value={employeeDraft.name} onChange={event => setEmployeeDraft({ ...employeeDraft, name: event.target.value })} className={INPUT_CLASS} /></label><label className="space-y-1.5 text-xs font-semibold">المبلغ المستحق عن كل يوم<input required type="text" inputMode="decimal" placeholder="مثال: 490+10 أو 12.375" value={employeeDraft.dailyRate} onChange={event => setEmployeeDraft({ ...employeeDraft, dailyRate: event.target.value })} className={INPUT_CLASS} /><span className="font-normal text-slate-500">{evaluateArithmeticExpression(employeeDraft.dailyRate) === null ? "أدخل مبلغاً أو عملية حسابية صحيحة." : `الأجر اليومي الناتج: ${formatDailyRate(evaluateArithmeticExpression(employeeDraft.dailyRate)!)}`}</span></label><label className="space-y-1.5 text-xs font-semibold">رقم الجوال (واتساب)<input type="tel" inputMode="tel" dir="ltr" value={employeeDraft.phone ?? ""} onChange={event => setEmployeeDraft({ ...employeeDraft, phone: event.target.value })} className={INPUT_CLASS} placeholder="05xxxxxxxx" /></label><label className="space-y-1.5 text-xs font-semibold">رقم الزبون في كشف الديون<input value={employeeDraft.customerNumber} onChange={event => setEmployeeDraft({ ...employeeDraft, customerNumber: event.target.value })} className={INPUT_CLASS} /></label></div>
        {employeeDraft.customerNumber && <p className="rounded-xl bg-amber-50 px-3 py-2 text-xs text-amber-900">رصيد الدين الحالي: {debts.has(employeeDraft.customerNumber) ? formatAmount(debts.get(employeeDraft.customerNumber)!) : "لا يوجد رصيد مرتبط"}؛ سيُحتسب التحصيل تلقائياً من الراتب حتى هذا الرصيد.</p>}
        <div className="flex justify-end gap-2 border-t pt-4"><button type="button" onClick={() => setEmployeeDraft(null)} disabled={busy} className="rounded-xl border px-4 py-2.5 text-sm font-semibold">إلغاء</button><button type="submit" disabled={busy} className="rounded-xl bg-cyan-800 px-5 py-2.5 text-sm font-semibold text-white disabled:opacity-50">{busy ? "جارٍ الحفظ..." : "حفظ الموظف"}</button></div>
      </form></div>}

      {creatingSheet && <div role="dialog" aria-modal="true" aria-labelledby="payroll-sheet-title" className="fixed inset-0 z-[80] flex items-start justify-center overflow-y-auto bg-slate-950/50 p-4 sm:items-center"><form onSubmit={event => void createSheet(event)} className="my-4 w-full max-w-2xl space-y-4 rounded-2xl bg-white p-5 shadow-2xl sm:p-7">
        <div className="flex items-center justify-between"><div><h2 id="payroll-sheet-title" className="text-lg font-bold">إنشاء كشف رواتب جديد</h2><p className="mt-1 text-xs text-slate-500">يُحفظ كل كشف مستقلاً عن الكشوف السابقة.</p></div><button type="button" onClick={() => setCreatingSheet(false)} aria-label="إغلاق" className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <div className="grid gap-3 sm:grid-cols-2"><label className="space-y-1.5 text-xs font-semibold sm:col-span-2">اسم كشف الرواتب<input required value={sheetForm.name} onChange={event => setSheetForm({ ...sheetForm, name: event.target.value })} placeholder="مثال: رواتب الأسبوع الأول - نوفمبر" className={INPUT_CLASS} /></label><label className="space-y-1.5 text-xs font-semibold">الفترة من<input required type="date" value={sheetForm.periodFrom} onChange={event => setSheetForm({ ...sheetForm, periodFrom: event.target.value })} className={INPUT_CLASS} /></label><label className="space-y-1.5 text-xs font-semibold">الفترة إلى<input required type="date" value={sheetForm.periodTo} onChange={event => setSheetForm({ ...sheetForm, periodTo: event.target.value })} className={INPUT_CLASS} /></label><label className="space-y-1.5 text-xs font-semibold">تاريخ / موعد الصرف<input required type="date" value={sheetForm.paymentDate} onChange={event => setSheetForm({ ...sheetForm, paymentDate: event.target.value })} className={INPUT_CLASS} /></label><label className="space-y-1.5 text-xs font-semibold">عدد أيام احتساب الراتب<input required type="number" min="1" max="366" step="1" value={sheetForm.days} onChange={event => setSheetForm({ ...sheetForm, days: Number(event.target.value) })} className={INPUT_CLASS} /><span className="font-normal text-slate-500">المبلغ اليومي × عدد الأيام (مثلاً 7 أو 14 أو 30).</span></label></div>
        <p className="rounded-xl bg-cyan-50 px-3 py-2 text-xs text-cyan-950">سيُضاف {data.employees.length} موظفاً إلى الكشف الحالي. تعديل بيانات الموظف لاحقاً لا يغيّر الكشوفات الأخرى.</p>
        <div className="flex justify-end gap-2 border-t pt-4"><button type="button" disabled={busy} onClick={() => setCreatingSheet(false)} className="rounded-xl border px-4 py-2.5 text-sm font-semibold">إلغاء</button><button type="submit" disabled={busy} className="rounded-xl bg-cyan-800 px-5 py-2.5 text-sm font-semibold text-white disabled:opacity-50">{busy ? "جارٍ الإنشاء..." : "إنشاء وحفظ الكشف"}</button></div>
      </form></div>}

      {sheetDraft && <div role="dialog" aria-modal="true" aria-labelledby="payroll-edit-sheet-title" className="fixed inset-0 z-[80] flex items-start justify-center overflow-y-auto bg-slate-950/50 p-4 sm:items-center"><form onSubmit={event => void editSheetMetadata(event)} className="my-4 w-full max-w-2xl space-y-4 rounded-2xl bg-white p-5 shadow-2xl sm:p-7">
        <div className="flex items-center justify-between"><div><h2 id="payroll-edit-sheet-title" className="text-lg font-bold">تعديل بيانات الكشف</h2><p className="mt-1 text-xs text-slate-500">تعديل عدد الأيام يعيد احتساب الإجمالي حسب الأجر اليومي المحفوظ لكل موظف.</p></div><button type="button" onClick={() => setSheetDraft(null)} aria-label="إغلاق" className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <div className="grid gap-3 sm:grid-cols-2"><label className="space-y-1.5 text-xs font-semibold sm:col-span-2">اسم الكشف<input required value={sheetDraft.name} onChange={event => setSheetDraft({ ...sheetDraft, name: event.target.value })} className={INPUT_CLASS} /></label><label className="space-y-1.5 text-xs font-semibold">من<input required type="date" value={sheetDraft.periodFrom} onChange={event => setSheetDraft({ ...sheetDraft, periodFrom: event.target.value })} className={INPUT_CLASS} /></label><label className="space-y-1.5 text-xs font-semibold">إلى<input required type="date" value={sheetDraft.periodTo} onChange={event => setSheetDraft({ ...sheetDraft, periodTo: event.target.value })} className={INPUT_CLASS} /></label><label className="space-y-1.5 text-xs font-semibold">موعد الصرف<input required type="date" value={sheetDraft.paymentDate} onChange={event => setSheetDraft({ ...sheetDraft, paymentDate: event.target.value })} className={INPUT_CLASS} /></label><label className="space-y-1.5 text-xs font-semibold">عدد الأيام<input required type="number" min="1" max="366" step="1" value={sheetDraft.days} onChange={event => setSheetDraft({ ...sheetDraft, days: Number(event.target.value) })} className={INPUT_CLASS} /></label></div>
        <div className="flex justify-end gap-2 border-t pt-4"><button type="button" onClick={() => setSheetDraft(null)} disabled={busy} className="rounded-xl border px-4 py-2.5 text-sm font-semibold">إلغاء</button><button type="submit" disabled={busy} className="rounded-xl bg-cyan-800 px-5 py-2.5 text-sm font-semibold text-white disabled:opacity-50">حفظ التعديلات</button></div>
      </form></div>}

      {unlockDraft && <div role="dialog" aria-modal="true" aria-labelledby="payroll-unlock-title" className="fixed inset-0 z-[90] flex items-center justify-center bg-slate-950/60 p-4"><form onSubmit={event => void unlockSheet(event)} className="w-full max-w-md space-y-4 rounded-2xl bg-white p-5 shadow-2xl sm:p-7">
        <div className="flex items-center justify-between"><div><h2 id="payroll-unlock-title" className="text-lg font-bold">فك تثبيت كشف الرواتب</h2><p className="mt-1 text-xs text-slate-500">أدخل كلمة السر للمتابعة.</p></div><button type="button" onClick={() => { setUnlockDraft(null); setUnlockError("") }} aria-label="إغلاق" className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <label className="block space-y-1.5 text-xs font-semibold">كلمة السر<input autoFocus required type="password" autoComplete="current-password" value={unlockDraft.password} onChange={event => { setUnlockDraft({ ...unlockDraft, password: event.target.value }); setUnlockError("") }} className={INPUT_CLASS} /></label>
        <p className="rounded-xl bg-cyan-50 px-3 py-2 text-xs text-cyan-950">كلمة السر الافتراضية هي اسم المستخدم الحالي.</p>
        {unlockError && <p role="alert" className="text-sm text-rose-700">{unlockError}</p>}
        <div className="flex justify-end gap-2 border-t pt-4"><button type="button" onClick={() => { setUnlockDraft(null); setUnlockError("") }} disabled={busy} className="rounded-xl border px-4 py-2.5 text-sm font-semibold">إلغاء</button><button type="submit" disabled={busy} className="rounded-xl bg-cyan-800 px-5 py-2.5 text-sm font-semibold text-white disabled:opacity-50">{busy ? "جارٍ التحقق..." : "فك التثبيت"}</button></div>
      </form></div>}

      {percentageDraft && <div role="dialog" aria-modal="true" aria-labelledby="payroll-percentage-title" className="fixed inset-0 z-[85] flex items-center justify-center bg-slate-950/50 p-4"><form onSubmit={event => void applyPercentage(event)} className="w-full max-w-lg space-y-4 rounded-2xl bg-white p-5 shadow-2xl sm:p-7">
        <div className="flex items-center justify-between"><div><h2 id="payroll-percentage-title" className="text-lg font-bold">تطبيق خصم بنسبة مئوية</h2><p className="mt-1 text-xs text-slate-500">يستبدل طريقة تحصيل الدين التلقائية وأي مبالغ مقبوضة مدخلة.</p></div><button type="button" onClick={() => setPercentageDraft(null)} aria-label="إغلاق" className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <label className="block space-y-1.5 text-xs font-semibold">النسبة %<input required type="number" min="0.01" max="100" step="0.01" value={percentageDraft.percentage} onChange={event => setPercentageDraft({ ...percentageDraft, percentage: event.target.value })} className={INPUT_CLASS} /></label>
        <label className="block space-y-1.5 text-xs font-semibold">أساس احتساب النسبة<select value={percentageDraft.base} onChange={event => setPercentageDraft({ ...percentageDraft, base: event.target.value as PayrollSheet["percentageBase"] })} className={INPUT_CLASS}><option value="debt">رصيد الدين في كشف الزبائن</option><option value="salary">إجمالي راتب الفترة</option><option value="manual">مبلغ أساس أدخله لكل موظف</option></select></label>
        <p className="rounded-xl bg-amber-50 px-3 py-2 text-xs leading-5 text-amber-900">سيطلب النظام تأكيداً قبل استبدال تحصيل الدين التلقائي أو أي مقبوضات مدخلة بحساب النسبة.</p>
        <div className="flex justify-end gap-2 border-t pt-4"><button type="button" onClick={() => setPercentageDraft(null)} className="rounded-xl border px-4 py-2.5 text-sm font-semibold">إلغاء</button><button type="submit" disabled={busy} className="rounded-xl bg-cyan-800 px-5 py-2.5 text-sm font-semibold text-white disabled:opacity-50">احتساب وتطبيق</button></div>
      </form></div>}

      {overwritePrompt && <div role="alertdialog" aria-modal="true" aria-labelledby="payroll-overwrite-title" className="fixed inset-0 z-[95] flex items-center justify-center bg-slate-950/60 p-4"><div className="w-full max-w-lg space-y-4 rounded-2xl bg-white p-6 shadow-2xl"><h2 id="payroll-overwrite-title" className="text-lg font-bold text-rose-900">تأكيد استبدال البيانات</h2><p className="text-sm leading-6 text-slate-700">تحتوي بعض أسطر الكشف على مبالغ مقبوضة أو تحصيل دين محسوب تلقائياً. تطبيق النسبة سيستبدل هذه القيم بحساب الخصم النسبي. هل تريد المتابعة؟</p><div className="flex justify-end gap-2"><button type="button" onClick={() => setOverwritePrompt(null)} className="rounded-xl border px-4 py-2.5 text-sm font-semibold">إلغاء دون حذف</button><button type="button" onClick={() => overwritePrompt()} className="inline-flex items-center gap-2 rounded-xl bg-rose-700 px-4 py-2.5 text-sm font-semibold text-white hover:bg-rose-800"><Check className="h-4 w-4" />نعم، احذف وطبّق النسبة</button></div></div></div>}
    </main>
  )
}
