"use client"

import { useEffect, useMemo, useRef, useState, type ChangeEvent } from "react"
import { ArrowDownLeft, ArrowUpRight, ChevronLeft, ChevronRight, FileSpreadsheet, Printer, RefreshCw, Search, Trash2, Upload } from "lucide-react"
import * as XLSX from "xlsx"
import { appendTransferData, loadTransferData, replaceTransferData, type TransferLoadProgress } from "../lib/transfer-storage"
import {
  EMPTY_TRANSFER_DATA,
  canonicalTransferDedupeKey,
  canonicalTransferReferenceKey,
  createTransferSearchMatcher,
  getTransferMonth,
  parseTransferRows,
  readTransferFile,
  type TransferColumns,
  type TransferData,
  type TransferImport,
  type TransferRecord,
  type TransferSheet,
  type TransferSource,
} from "../lib/transfers"
import { formatMonthKey } from "../lib/table-view"
import { TASWIYAH_LOGO_PATH } from "../lib/report-branding"
import { useGlobalLoading } from "./global-loading"

const SOURCES: Array<{ value: TransferSource; label: string }> = [
  { value: "bank", label: "كشف البنك" },
  { value: "wallet", label: "محفظة بال بي" },
  { value: "platform", label: "منصة بال بي" },
]

const COLUMN_OPTIONS: Array<{ key: keyof TransferColumns; label: string; required?: boolean }> = [
  { key: "date", label: "التاريخ", required: true },
  { key: "reference", label: "الرقم المرجعي" },
  { key: "name", label: "الاسم" },
  { key: "amount", label: "المبلغ" },
  { key: "debit", label: "مبلغ الصادر / المدفوع" },
  { key: "credit", label: "مبلغ الوارد / المستلم" },
  { key: "direction", label: "نوع الحركة" },
  { key: "details", label: "التفاصيل / الإيضاحات" },
  { key: "currency", label: "العملة" },
]

const PAGE_SIZE = 10
const DEFAULT_FILTERS = {
  search: "",
  source: "all" as TransferSource | "all",
  direction: "all" as TransferRecord["direction"] | "all",
  account: "all",
  sort: "date-asc" as "date-asc" | "date-desc" | "amount-asc" | "amount-desc",
  month: "all",
  dateFrom: "",
  dateTo: "",
  minimum: "",
  maximum: "",
}

type TransferFilters = typeof DEFAULT_FILTERS
type ReviewDecision = "approve" | "reject"
type TransferReviewItem = {
  kind: "duplicate" | "mismatch"
  rowNumber: number
  record: TransferRecord
  matchedRecord: TransferRecord
  matchedRowNumber: number | null
  key: string
}

function inputClass() {
  return "w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-normal outline-none transition focus:border-cyan-600 focus:ring-2 focus:ring-cyan-100"
}

function labelClass() {
  return "block space-y-1.5 text-xs font-semibold text-slate-700"
}

function formatAmount(value: number) {
  return new Intl.NumberFormat("ar", { maximumFractionDigits: 2 }).format(value)
}

function formatCurrencyTotals(values: Map<string, number>) {
  if (values.size === 0) return formatAmount(0)
  return [...values.entries()]
    .sort(([left], [right]) => left.localeCompare(right, "ar"))
    .map(([currency, amount]) => `${formatAmount(amount)} ${currency || "عملة غير محددة"}`)
    .join(" · ")
}

function sumCurrencies(records: TransferRecord[], direction: TransferRecord["direction"]) {
  const totals = new Map<string, number>()
  for (const record of records) {
    if (record.direction !== direction) continue
    const currency = record.currency || "غير محددة"
    totals.set(currency, (totals.get(currency) ?? 0) + record.amount)
  }
  return totals
}

function sourceLabel(source: TransferSource) {
  return SOURCES.find(item => item.value === source)?.label ?? source
}

function transferAccountLabel(record: Pick<TransferRecord, "accountName" | "currency" | "source">) {
  const account = record.accountName || sourceLabel(record.source)
  return record.currency ? `${account} - ${record.currency}` : account
}

function amountValue(value: string) {
  let normalized = value
    .replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, digit => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)))
    .trim()
  if (!normalized) return null
  if (normalized.includes(",") && normalized.includes(".")) {
    normalized = normalized.lastIndexOf(",") > normalized.lastIndexOf(".")
      ? normalized.replace(/\./g, "").replace(",", ".")
      : normalized.replace(/,/g, "")
  } else if (normalized.includes(",")) {
    normalized = /,\d{1,2}$/.test(normalized) ? normalized.replace(",", ".") : normalized.replace(/,/g, "")
  }
  const parsed = Number(normalized)
  return Number.isFinite(parsed) ? parsed : null
}

export function TransfersManagement({ username }: { username: string }) {
  const normalizedUsername = username.normalize("NFKC").trim().toLocaleLowerCase("en-US")
  const storageKey = `transfers_ledger:${normalizedUsername}`
  const [data, setData] = useState<TransferData>(EMPTY_TRANSFER_DATA)
  const dataRef = useRef<TransferData>(EMPTY_TRANSFER_DATA)
  const [loadedKey, setLoadedKey] = useState<string | null>(null)
  const [loading, setLoading] = useState(true)
  const [loadProgress, setLoadProgress] = useState<TransferLoadProgress>({
    loadedRecords: 0,
    loadedChunks: 0,
    totalChunks: 0,
  })
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState("")
  const [notice, setNotice] = useState("")
  const [source, setSource] = useState<TransferSource>("bank")
  const [institutionName, setInstitutionName] = useState("بنك فلسطين")
  const [accountType, setAccountType] = useState("جاري")
  const [importCurrency, setImportCurrency] = useState("شيكل")
  const [file, setFile] = useState<File | null>(null)
  const [sheet, setSheet] = useState<TransferSheet | null>(null)
  const [columns, setColumns] = useState<TransferColumns | null>(null)
  const [reviewDecisions, setReviewDecisions] = useState<Record<string, ReviewDecision>>({})
  const [search, setSearch] = useState("")
  const [sourceFilter, setSourceFilter] = useState<TransferSource | "all">("all")
  const [directionFilter, setDirectionFilter] = useState<TransferRecord["direction"] | "all">("all")
  const [accountFilter, setAccountFilter] = useState("all")
  const [sortOrder, setSortOrder] = useState<"date-asc" | "date-desc" | "amount-asc" | "amount-desc">("date-asc")
  const [monthFilter, setMonthFilter] = useState("all")
  const [dateFrom, setDateFrom] = useState("")
  const [dateTo, setDateTo] = useState("")
  const [minimum, setMinimum] = useState("")
  const [maximum, setMaximum] = useState("")
  const [page, setPage] = useState(0)
  const [printRows, setPrintRows] = useState<TransferRecord[] | null>(null)
  const [appliedFilters, setAppliedFilters] = useState<TransferFilters>(DEFAULT_FILTERS)
  const [filteredRecords, setFilteredRecords] = useState<TransferRecord[]>([])
  const [searching, setSearching] = useState(false)

  const loadProgressPercent = loadProgress.totalChunks
    ? Math.round((loadProgress.loadedChunks / loadProgress.totalChunks) * 100)
    : undefined
  useGlobalLoading(
    loading,
    "جاري تحميل سجل الحوالات",
    `تم تحميل ${loadProgress.loadedRecords.toLocaleString("ar")} حوالة`,
    loadProgressPercent,
  )
  useGlobalLoading(busy, "جاري معالجة ملف الحوالات...")

  useEffect(() => {
    if (!printRows) return
    const beforePrint = () => document.body.classList.add("transfer-printing")
    const afterPrint = () => {
      document.body.classList.remove("transfer-printing")
      setPrintRows(null)
    }
    window.addEventListener("beforeprint", beforePrint)
    window.addEventListener("afterprint", afterPrint)
    const timer = window.setTimeout(() => window.print(), 0)
    return () => {
      window.clearTimeout(timer)
      window.removeEventListener("beforeprint", beforePrint)
      window.removeEventListener("afterprint", afterPrint)
      document.body.classList.remove("transfer-printing")
    }
  }, [printRows])

  useEffect(() => {
    let active = true
    setLoading(true)
    setLoadProgress({ loadedRecords: 0, loadedChunks: 0, totalChunks: 0 })
    setLoadedKey(null)
    setData(EMPTY_TRANSFER_DATA)
    dataRef.current = EMPTY_TRANSFER_DATA
    setError("")
    setNotice("")
    const load = async () => {
      try {
        const { data: stored, warning } = await loadTransferData(storageKey, progress => {
          if (active) setLoadProgress(progress)
        })
        if (!Array.isArray(stored.records) || !Array.isArray(stored.imports)) {
          throw new Error("بيانات الحوالات المحفوظة غير صالحة.")
        }
        if (active) {
          dataRef.current = stored
          setData(stored)
          setNotice(warning ?? "")
          setLoadedKey(storageKey)
        }
      } catch (cause) {
        if (active) setError(cause instanceof Error ? `تعذر تحميل بيانات الحوالات: ${cause.message}` : "تعذر تحميل بيانات الحوالات.")
      } finally {
        if (active) setLoading(false)
      }
    }
    void load()
    return () => { active = false }
  }, [storageKey])

  const monthOptions = useMemo(() => {
    const months = new Set(data.records.map(getTransferMonth).filter((month): month is string => !!month))
    return [...months].sort()
  }, [data.records])

  useEffect(() => {
    let active = true
    let offset = 0
    const matches: TransferRecord[] = []
    const queryMatches = createTransferSearchMatcher(appliedFilters.search)
    const minimumAmount = amountValue(appliedFilters.minimum)
    const maximumAmount = amountValue(appliedFilters.maximum)
    setSearching(true)

    const processBatch = () => {
      if (!active) return
      const end = Math.min(offset + 2500, data.records.length)
      for (; offset < end; offset++) {
        const record = data.records[offset]
        if (appliedFilters.source !== "all" && record.source !== appliedFilters.source) continue
        if (appliedFilters.direction !== "all" && record.direction !== appliedFilters.direction) continue
        if (appliedFilters.account !== "all" && transferAccountLabel(record) !== appliedFilters.account) continue
        if (appliedFilters.month !== "all" && getTransferMonth(record) !== appliedFilters.month) continue
        if (appliedFilters.dateFrom && record.date < appliedFilters.dateFrom) continue
        if (appliedFilters.dateTo && record.date > appliedFilters.dateTo) continue
        if (minimumAmount !== null && record.amount < minimumAmount) continue
        if (maximumAmount !== null && record.amount > maximumAmount) continue
        if (!queryMatches(record)) continue
        matches.push(record)
      }
      if (offset < data.records.length) {
        window.setTimeout(processBatch, 0)
        return
      }
      matches.sort((left, right) => {
        if (appliedFilters.sort === "amount-asc") return left.amount - right.amount || left.date.localeCompare(right.date)
        if (appliedFilters.sort === "amount-desc") return right.amount - left.amount || left.date.localeCompare(right.date)
        if (appliedFilters.sort === "date-desc") return right.date.localeCompare(left.date) || left.id.localeCompare(right.id)
        return left.date.localeCompare(right.date) || left.id.localeCompare(right.id)
      })
      setFilteredRecords(matches)
      setSearching(false)
    }

    processBatch()
    return () => { active = false }
  }, [appliedFilters, data.records])

  const pageCount = Math.max(1, Math.ceil(filteredRecords.length / PAGE_SIZE))
  const visibleRecords = filteredRecords.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE)
  const totals = useMemo(() => {
    const result = {
      count: 0,
      incomingCount: 0,
      outgoingCount: 0,
      incomingAmounts: new Map<string, number>(),
      outgoingAmounts: new Map<string, number>(),
    }
    for (const record of data.records) {
      result.count++
      const amounts = record.direction === "incoming" ? result.incomingAmounts : result.outgoingAmounts
      if (record.direction === "incoming") result.incomingCount++
      else result.outgoingCount++
      const currency = record.currency || "غير محددة"
      amounts.set(currency, (amounts.get(currency) ?? 0) + record.amount)
    }
    return result
  }, [data.records])

  const perSource = useMemo(() => SOURCES.map(item => {
    const records = data.records.filter(record => record.source === item.value)
    return {
      ...item,
      count: records.length,
      incoming: sumCurrencies(records, "incoming"),
      outgoing: sumCurrencies(records, "outgoing"),
    }
  }), [data.records])
  const accountOptions = useMemo(
    () => [...new Set(data.records.map(transferAccountLabel))].sort((a, b) => a.localeCompare(b, "ar")),
    [data.records],
  )
  const accountName = [institutionName.trim(), accountType.trim()].filter(Boolean).join(" - ")

  const parsedPreview = useMemo(() => {
    if (!sheet || !columns || columns.date === null
      || (columns.amount === null && columns.debit === null && columns.credit === null)) return null
    return parseTransferRows(sheet.rows, columns, source, accountName.trim(), importCurrency.trim())
  }, [accountName, columns, importCurrency, sheet, source])

  const duplicateReview = useMemo<TransferReviewItem[]>(() => {
    if (!parsedPreview) return []
    const savedByKey = new Map(data.records.map(record => [canonicalTransferDedupeKey(record), record]))
    const savedByReference = new Map(data.records
      .filter(record => !!record.reference)
      .map(record => [canonicalTransferReferenceKey(record), record]))
    const rowByKey = new Map(parsedPreview.recordRows.map(item => [item.record.dedupeKey, item.rowNumber]))
    const existing = parsedPreview.records.flatMap((record): TransferReviewItem[] => {
      const rowNumber = rowByKey.get(record.dedupeKey) ?? 0
      const exactMatch = savedByKey.get(canonicalTransferDedupeKey(record))
      if (exactMatch) return [{
        kind: "duplicate" as const,
        rowNumber,
        record,
        matchedRecord: exactMatch,
        matchedRowNumber: null,
        key: `saved:${record.dedupeKey}`,
      }]
      const referenceMatch = savedByReference.get(canonicalTransferReferenceKey(record))
      return referenceMatch ? [{
        kind: "mismatch" as const,
        rowNumber,
        record,
        matchedRecord: referenceMatch,
        matchedRowNumber: null,
        key: `saved-mismatch:${record.dedupeKey}`,
      }] : []
    })
    return [
      ...parsedPreview.duplicates.map((duplicate): TransferReviewItem => ({
        ...duplicate,
        kind: "duplicate" as const,
        key: `file:${duplicate.rowNumber}:${duplicate.record.dedupeKey}`,
      })),
      ...parsedPreview.mismatches.map((mismatch): TransferReviewItem => ({
        ...mismatch,
        kind: "mismatch" as const,
        key: `file-mismatch:${mismatch.rowNumber}:${mismatch.record.dedupeKey}`,
      })),
      ...existing,
    ].sort((left, right) => left.rowNumber - right.rowNumber)
  }, [data.records, parsedPreview])
  const pendingReviewCount = duplicateReview.filter(item => !reviewDecisions[item.key]).length
  const reviewByRowAndKey = useMemo(
    () => new Map(duplicateReview.map(item => [`${item.rowNumber}:${item.record.dedupeKey}`, item])),
    [duplicateReview],
  )
  const filtersDirty = JSON.stringify({
    search,
    source: sourceFilter,
    direction: directionFilter,
    account: accountFilter,
    sort: sortOrder,
    month: monthFilter,
    dateFrom,
    dateTo,
    minimum,
    maximum,
  }) !== JSON.stringify(appliedFilters)

  const saveImport = async (
    newRecords: TransferRecord[],
    imported: TransferImport,
    successMessage: string,
  ) => {
    try {
      const { data: next, cleanupWarning } = await appendTransferData(storageKey, dataRef.current, newRecords, imported)
      dataRef.current = next
      setData(next)
      setNotice(cleanupWarning ? `${successMessage} ${cleanupWarning}` : successMessage)
      setError("")
      return true
    } catch (cause) {
      setError(cause instanceof Error
        ? `تعذر حفظ البيانات؛ لم يتم اعتماد التعديل: ${cause.message}`
        : "تعذر حفظ البيانات؛ لم يتم اعتماد التعديل.")
      return false
    }
  }

  const deleteImport = async (item: TransferImport) => {
    if (busy) return
    const linkedCount = dataRef.current.records.filter(record => record.importId === item.id).length
    const legacyWarning = linkedCount === 0 && item.importedCount > 0
      ? "\nتنبيه: هذا الملف أُضيف قبل ربط الحركات بملفاتها، لذلك سيُحذف سجل الملف فقط ولن يمكن تحديد حركاته القديمة بأمان."
      : ""
    if (!window.confirm(`حذف الملف «${item.accountName || item.fileName}»؟ سيتم حذف ${linkedCount.toLocaleString("ar")} حوالة مرتبطة به.${legacyWarning}`)) return
    setBusy(true)
    setError("")
    setNotice("")
    try {
      const next: TransferData = {
        records: dataRef.current.records.filter(record => record.importId !== item.id),
        imports: dataRef.current.imports.filter(imported => imported.id !== item.id),
      }
      const { cleanupWarning } = await replaceTransferData(storageKey, next)
      dataRef.current = next
      setData(next)
      setPage(0)
      setNotice(cleanupWarning
        ? `تم حذف الملف والحركات المرتبطة. ${cleanupWarning}`
        : `تم حذف الملف والحركات المرتبطة (${linkedCount.toLocaleString("ar")}).`)
    } catch (cause) {
      setError(cause instanceof Error
        ? `تعذر حذف الملف؛ لم يتم اعتماد التعديل: ${cause.message}`
        : "تعذر حذف الملف؛ لم يتم اعتماد التعديل.")
    } finally {
      setBusy(false)
    }
  }

  const clearAllTransfers = async () => {
    if (busy) return
    const recordCount = dataRef.current.records.length
    const importCount = dataRef.current.imports.length
    if (!window.confirm(`تأكيد حذف جميع بيانات منصة الحوالات؟ سيتم حذف ${recordCount.toLocaleString("ar")} حوالة و${importCount.toLocaleString("ar")} ملف استيراد نهائياً.`)) return
    setBusy(true)
    setError("")
    setNotice("")
    try {
      const { cleanupWarning } = await replaceTransferData(storageKey, EMPTY_TRANSFER_DATA)
      dataRef.current = EMPTY_TRANSFER_DATA
      setData(EMPTY_TRANSFER_DATA)
      setPage(0)
      setNotice(cleanupWarning
        ? `تم حذف جميع الحوالات وملفات الاستيراد. ${cleanupWarning}`
        : "تم حذف جميع الحوالات وملفات الاستيراد من المنصة.")
    } catch (cause) {
      setError(cause instanceof Error
        ? `تعذر حذف بيانات الحوالات؛ لم يتم اعتماد التعديل: ${cause.message}`
        : "تعذر حذف بيانات الحوالات؛ لم يتم اعتماد التعديل.")
    } finally {
      setBusy(false)
    }
  }

  const handleFile = async (event: ChangeEvent<HTMLInputElement>) => {
    const selected = event.target.files?.[0] ?? null
    event.target.value = ""
    setFile(selected)
    setSheet(null)
    setColumns(null)
    setReviewDecisions({})
    setError("")
    setNotice("")
    if (!selected) return
    setBusy(true)
    try {
      const loaded = await readTransferFile(selected)
      setSheet(loaded)
      setColumns(loaded.columns)
      setNotice(`تمت قراءة ${loaded.rows.length.toLocaleString("ar")} صفاً. راجع مطابقة الأعمدة قبل الاستيراد.`)
    } catch (cause) {
      setFile(null)
      setError(cause instanceof Error ? cause.message : "تعذرت قراءة الملف.")
    } finally {
      setBusy(false)
    }
  }

  const importFile = async () => {
    if (!file || !sheet || !columns || loadedKey !== storageKey || busy) return
    if (!institutionName.trim() || !accountType.trim() || !importCurrency.trim()) {
      setError("أدخل اسم البنك أو المحفظة ونوع الحساب والعملة قبل الاستيراد.")
      return
    }
    if (pendingReviewCount > 0) {
      setError(`راجع كل الحركات المكررة أو غير المتطابقة واختر اعتماداً أو رفضاً لكل حركة قبل الاستيراد. المتبقي: ${pendingReviewCount.toLocaleString("ar")}.`)
      return
    }
    setBusy(true)
    setError("")
    setNotice("")
    try {
      const parsed = parseTransferRows(sheet.rows, columns, source, accountName.trim(), importCurrency.trim())
      const newRecords: TransferRecord[] = []
      let rejectedCount = 0
      const rowByRecordKey = new Map(parsed.recordRows.map(item => [item.record.dedupeKey, item.rowNumber]))
      const importId = typeof globalThis.crypto?.randomUUID === "function" ? globalThis.crypto.randomUUID() : `${Date.now()}`
      for (const record of parsed.records) {
        const rowNumber = rowByRecordKey.get(record.dedupeKey) ?? 0
        const review = reviewByRowAndKey.get(`${rowNumber}:${record.dedupeKey}`)
        if (review && reviewDecisions[review.key] !== "approve") {
          rejectedCount++
          continue
        }
        newRecords.push({
          ...record,
          importId,
          accountName: accountName.trim(),
          currency: importCurrency.trim(),
          ...(review ? {
            id: typeof globalThis.crypto?.randomUUID === "function" ? globalThis.crypto.randomUUID() : `${Date.now()}-${record.id}`,
            dedupeKey: `${record.dedupeKey}:approved:${record.id}`,
          } : {}),
        })
      }
      for (const duplicate of parsed.duplicates) {
        const key = `file:${duplicate.rowNumber}:${duplicate.record.dedupeKey}`
        if (reviewDecisions[key] !== "approve") {
          rejectedCount++
          continue
        }
        newRecords.push({
          ...duplicate.record,
          importId,
          accountName: accountName.trim(),
          currency: importCurrency.trim(),
          id: typeof globalThis.crypto?.randomUUID === "function" ? globalThis.crypto.randomUUID() : `${Date.now()}-${duplicate.rowNumber}`,
          dedupeKey: `${duplicate.record.dedupeKey}:approved:${duplicate.rowNumber}`,
        })
      }
      for (const mismatch of parsed.mismatches) {
        const key = `file-mismatch:${mismatch.rowNumber}:${mismatch.record.dedupeKey}`
        if (reviewDecisions[key] !== "approve") {
          rejectedCount++
          continue
        }
        newRecords.push({
          ...mismatch.record,
          importId,
          accountName: accountName.trim(),
          currency: importCurrency.trim(),
          id: typeof globalThis.crypto?.randomUUID === "function" ? globalThis.crypto.randomUUID() : `${Date.now()}-${mismatch.rowNumber}`,
          dedupeKey: `${mismatch.record.dedupeKey}:approved:${mismatch.rowNumber}`,
        })
      }
      if (newRecords.length === 0) {
        throw new Error("لم تُضف أي حوالة؛ تم رفض جميع الحركات أو لم توجد بيانات صالحة.")
      }

      const importedAt = new Date().toISOString()
      const imported: TransferImport = {
        id: importId,
        source,
        accountName: accountName.trim(),
        currency: importCurrency.trim(),
        fileName: file.name,
        importedAt,
        importedCount: newRecords.length,
        duplicateCount: rejectedCount,
      }
      const successMessage = `تم استيراد ${newRecords.length.toLocaleString("ar")} حوالة.`
        + (rejectedCount ? ` تم رفض ${rejectedCount.toLocaleString("ar")} حركة بعد المراجعة.` : "")
        + (parsed.invalidCount ? ` تعذر اعتماد ${parsed.invalidCount.toLocaleString("ar")} صفاً غير مكتمل.` : "")
      if (await saveImport(newRecords, imported, successMessage)) {
        setFile(null)
        setSheet(null)
        setColumns(null)
        setReviewDecisions({})
        setPage(0)
      }
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذر استيراد الحوالات.")
    } finally {
      setBusy(false)
    }
  }

  const updateColumn = (key: keyof TransferColumns, value: string) => {
    setReviewDecisions({})
    setColumns(current => current ? { ...current, [key]: value === "" ? null : Number(value) } : current)
  }

  const exportReviewWorkbook = () => {
    if (!parsedPreview || (!duplicateReview.length && !parsedPreview.invalidRows.length)) return
    try {
      const workbook = XLSX.utils.book_new()
      if (duplicateReview.length) {
        const reviewRows = duplicateReview.map(item => [
          item.kind === "duplicate" ? "مكررة: تطابق المرجع والمبلغ" : "غير متطابقة: المرجع نفسه بمبلغ مختلف",
          item.rowNumber,
          reviewDecisions[item.key] === "approve" ? "اعتماد" : reviewDecisions[item.key] === "reject" ? "رفض" : "بانتظار القرار",
          item.record.reference,
          item.record.amount,
          item.record.currency,
          item.record.date,
          item.record.name,
          item.matchedRowNumber ?? "حركة محفوظة",
          item.matchedRecord.reference,
          item.matchedRecord.amount,
          item.matchedRecord.currency,
          item.matchedRecord.date,
          item.matchedRecord.name,
          accountName,
        ])
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
          ["الحالة", "صف الملف", "قرار المراجعة", "مرجع الحركة", "مبلغ الحركة", "العملة", "تاريخها", "اسمها", "الصف المطابق", "المرجع المطابق", "المبلغ المطابق", "عملة المطابق", "تاريخ المطابق", "اسم المطابق", "الحساب"],
          ...reviewRows,
        ]), "مراجعة الحركات")
      }
      if (parsedPreview.invalidRows.length) {
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
          ["رقم الصف", "سبب عدم الصلاحية", ...(sheet?.headers ?? [])],
          ...parsedPreview.invalidRows.map(item => [item.rowNumber, item.reason, ...item.values.map(value => String(value ?? ""))]),
        ]), "صفوف غير صالحة")
      }
      const fileBase = file?.name.replace(/\.[^.]+$/, "") || "كشف-الحوالات"
      const safeName = fileBase.replace(/[\\/:*?"<>|]+/g, "-")
      XLSX.writeFile(workbook, `مراجعة-${safeName}.xlsx`)
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر تنزيل ملف المراجعة: ${cause.message}` : "تعذر تنزيل ملف المراجعة.")
    }
  }

  const resetFilters = () => {
    setSearch("")
    setSourceFilter("all")
    setDirectionFilter("all")
    setAccountFilter("all")
    setMonthFilter("all")
    setDateFrom("")
    setDateTo("")
    setMinimum("")
    setMaximum("")
    setSortOrder("date-asc")
    setAppliedFilters(DEFAULT_FILTERS)
    setPage(0)
  }

  const applyFilters = () => {
    setAppliedFilters({
      search,
      source: sourceFilter,
      direction: directionFilter,
      account: accountFilter,
      sort: sortOrder,
      month: monthFilter,
      dateFrom,
      dateTo,
      minimum,
      maximum,
    })
    setPage(0)
  }

  const exportFilteredRecords = () => {
    if (!filteredRecords.length) return
    const workbook = XLSX.utils.book_new()
    const worksheet = XLSX.utils.json_to_sheet(filteredRecords.map(record => ({
      "التاريخ": record.date,
      "اسم الكشف": transferAccountLabel(record),
      "الاسم": record.name,
      "الرقم المرجعي": record.reference,
      "الاتجاه": record.direction === "incoming" ? "وارد" : "صادر",
      "المبلغ": record.amount,
      "العملة": record.currency,
      "التفاصيل": record.details,
    })))
    XLSX.utils.book_append_sheet(workbook, worksheet, "الحوالات المفلترة")
    XLSX.writeFile(workbook, `الحوالات-المفلترة-${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  if (loading) {
    return null
  }

  if (loadedKey !== storageKey && error) {
    return <div role="alert" className="m-6 rounded-xl border border-red-200 bg-red-50 p-4 text-sm text-red-800">{error}</div>
  }

  return (
    <div className="mx-auto max-w-[1500px] space-y-6 p-4 md:p-6">
      <header className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">منصة الحوالات</h1>
          <p className="mt-1 text-sm text-slate-500">استيراد كشوف البنك والمحافظ، حفظها تراكميّاً، ومنع تكرار الحوالات.</p>
        </div>
        <div className="rounded-xl border border-cyan-100 bg-cyan-50 px-3 py-2 text-xs font-medium text-cyan-900">
          تُحفظ البيانات حسب حساب المستخدم
        </div>
      </header>

      {error && <div role="alert" className="rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800">{error}</div>}
      {notice && <div role="status" className="rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-800">{notice}</div>}

      <section className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">
        <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
          <div className="text-xs font-semibold text-slate-500">إجمالي الحوالات المحفوظة</div>
          <div className="mt-2 text-2xl font-bold text-slate-900">{totals.count.toLocaleString("ar")}</div>
        </div>
        <div className="rounded-2xl border border-emerald-100 bg-emerald-50 p-4 shadow-sm">
          <div className="flex items-center gap-2 text-xs font-semibold text-emerald-800"><ArrowDownLeft className="h-4 w-4" /> الوارد</div>
          <div className="mt-2 text-base font-bold text-emerald-950">{formatCurrencyTotals(totals.incomingAmounts)}</div>
          <div className="mt-1 text-xs text-emerald-800">{totals.incomingCount.toLocaleString("ar")} حوالة</div>
        </div>
        <div className="rounded-2xl border border-rose-100 bg-rose-50 p-4 shadow-sm">
          <div className="flex items-center gap-2 text-xs font-semibold text-rose-800"><ArrowUpRight className="h-4 w-4" /> الصادر</div>
          <div className="mt-2 text-base font-bold text-rose-950">{formatCurrencyTotals(totals.outgoingAmounts)}</div>
          <div className="mt-1 text-xs text-rose-800">{totals.outgoingCount.toLocaleString("ar")} حوالة</div>
        </div>
        {perSource.map(item => (
          <div key={item.value} className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
            <div className="text-xs font-semibold text-slate-500">{item.label}</div>
            <div className="mt-2 text-xl font-bold text-slate-900">{item.count.toLocaleString("ar")}</div>
            <div className="mt-1 text-[11px] text-slate-500">وارد {formatCurrencyTotals(item.incoming)} · صادر {formatCurrencyTotals(item.outgoing)}</div>
          </div>
        ))}
      </section>

      <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
        <div className="mb-4 flex items-center gap-2">
          <Upload className="h-5 w-5 text-cyan-700" />
          <h2 className="text-lg font-bold text-slate-900">إضافة كشف جديد</h2>
        </div>
        <div className="grid gap-4 md:grid-cols-2">
          <label className={labelClass()}>
            <span>نوع المصدر</span>
            <select className={inputClass()} value={source} onChange={event => {
              const nextSource = event.target.value as TransferSource
              setSource(nextSource)
              setReviewDecisions({})
              setInstitutionName(nextSource === "bank" ? "بنك فلسطين" : nextSource === "wallet" ? "بال بي" : "جوال بي")
              setAccountType(nextSource === "bank" ? "جاري" : "محفظة")
            }}>
              {SOURCES.map(item => <option key={item.value} value={item.value}>{item.label}</option>)}
            </select>
          </label>
          <label className={labelClass()}>
            <span>اسم البنك أو المحفظة</span>
            <input className={inputClass()} required list="transfer-institutions" value={institutionName} onChange={event => { setInstitutionName(event.target.value); setReviewDecisions({}) }} placeholder="اكتب اسم البنك أو المحفظة" />
            <datalist id="transfer-institutions">
              <option value="بنك فلسطين" />
              <option value="البنك الإسلامي الفلسطيني" />
              <option value="البنك العربي الإسلامي" />
              <option value="بنك القدس" />
              <option value="بال بي" />
              <option value="جوال بي" />
            </datalist>
          </label>
          <label className={labelClass()}>
            <span>نوع الحساب</span>
            <input className={inputClass()} required list="transfer-account-types" value={accountType} onChange={event => { setAccountType(event.target.value); setReviewDecisions({}) }} placeholder="جاري أو توفير" />
            <datalist id="transfer-account-types">
              <option value="جاري" />
              <option value="توفير" />
              <option value="محفظة" />
            </datalist>
          </label>
          <label className={labelClass()}>
            <span>عملة هذا الكشف</span>
            <input className={inputClass()} required list="transfer-currencies" value={importCurrency} onChange={event => { setImportCurrency(event.target.value); setReviewDecisions({}) }} placeholder="شيكل أو دولار أو يورو" />
            <datalist id="transfer-currencies">
              <option value="شيكل" />
              <option value="دولار" />
              <option value="يورو" />
              <option value="دينار" />
            </datalist>
          </label>
          <label className={labelClass()}>
            <span>ملف Excel أو CSV</span>
            <input className={inputClass()} type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.csv" onChange={event => void handleFile(event)} />
          </label>
        </div>
        {busy && <div className="mt-4 flex items-center gap-2 text-sm text-cyan-800"><RefreshCw className="h-4 w-4 animate-spin" /> جارٍ معالجة ملف الحوالات...</div>}
        {sheet && columns && file && (
          <div className="mt-5 space-y-4 rounded-xl border border-slate-200 bg-slate-50 p-4">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="text-sm font-semibold text-slate-800">{file.name} · {sheet.rows.length.toLocaleString("ar")} صفاً</div>
              <div className="text-xs text-slate-500">راجع تعيين الأعمدة، خصوصاً التاريخ والمبلغ والرقم المرجعي.</div>
            </div>
            <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
              {COLUMN_OPTIONS.map(option => (
                <label key={option.key} className={labelClass()}>
                  <span>{option.label}{option.required ? " *" : ""}</span>
                  <select
                    className={inputClass()}
                    value={columns[option.key] ?? ""}
                    onChange={event => updateColumn(option.key, event.target.value)}
                  >
                    <option value="">غير محدد</option>
                    {sheet.headers.map((header, index) => <option key={`${index}-${header}`} value={index}>{header}</option>)}
                  </select>
                </label>
              ))}
            </div>
            {parsedPreview && (duplicateReview.length > 0 || parsedPreview.invalidCount > 0) && (
              <div className="space-y-4 rounded-xl border border-amber-200 bg-amber-50 p-4">
                <div className="flex flex-wrap items-start justify-between gap-3">
                  <div>
                  <h3 className="text-sm font-bold text-amber-950">مراجعة الحركات قبل الاستيراد</h3>
                  <p className="mt-1 text-xs text-amber-900">
                    {duplicateReview.filter(item => item.kind === "duplicate").length.toLocaleString("ar")} مكررة · {duplicateReview.filter(item => item.kind === "mismatch").length.toLocaleString("ar")} غير متطابقة · {parsedPreview.invalidCount.toLocaleString("ar")} صفوف غير صالحة
                    {parsedPreview.invalidCount ? " · تظهر جميع الصفوف غير الصالحة" : ""}
                  </p>
                  {duplicateReview.length > 0 && <p className="mt-1 text-xs font-semibold text-amber-950">لم يُستورد بعد: اختر اعتماد أو رفض لكل حركة ({pendingReviewCount.toLocaleString("ar")} بانتظار القرار).</p>}
                  </div>
                  <button type="button" onClick={exportReviewWorkbook} className="inline-flex items-center gap-2 rounded-lg border border-emerald-300 bg-white px-3 py-2 text-xs font-semibold text-emerald-800 hover:bg-emerald-50"><FileSpreadsheet className="h-4 w-4" />تنزيل ملف المراجعة Excel</button>
                </div>
                {duplicateReview.length > 0 && (
                  <div className="space-y-2">
                    <h4 className="text-xs font-bold text-slate-800">الحركات المتكررة أو التي لها المرجع نفسه ومبلغ مختلف</h4>
                    <div className="max-h-64 space-y-2 overflow-auto">
                      {duplicateReview.map(item => {
                        const decision = reviewDecisions[item.key]
                        return <div key={item.key} className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-amber-200 bg-white p-3 text-xs">
                          <div className="space-y-1">
                            <p className="font-semibold text-slate-900">
                              {item.kind === "duplicate" ? "مكررة" : "مبلغ غير متطابق"} · صف {item.rowNumber.toLocaleString("ar")}: {item.record.name || item.record.details || "حركة بدون اسم"} · {item.record.date} · المرجع {item.record.reference || "—"} · {formatAmount(item.record.amount)} {item.record.currency}
                            </p>
                            <p className="text-slate-600">
                              مطابق لـ {item.matchedRowNumber === null
                                ? "حركة محفوظة مسبقاً"
                                : `الحركة في صف ${item.matchedRowNumber.toLocaleString("ar")} من نفس الملف`}
                              : {item.matchedRecord.name || item.matchedRecord.details || "بدون اسم"} · {item.matchedRecord.date} · المرجع {item.matchedRecord.reference || "—"} · {formatAmount(item.matchedRecord.amount)} {item.matchedRecord.currency}
                            </p>
                            <p className={`font-semibold ${decision === "approve" ? "text-emerald-700" : decision === "reject" ? "text-rose-700" : "text-amber-800"}`}>
                              {decision === "approve" ? "القرار: اعتماد" : decision === "reject" ? "القرار: رفض" : "بانتظار القرار"}
                            </p>
                          </div>
                          <div className="flex gap-2">
                            <button type="button" onClick={() => setReviewDecisions(current => ({ ...current, [item.key]: "approve" }))} className={`rounded-lg px-3 py-2 font-semibold ${decision === "approve" ? "bg-emerald-700 text-white" : "border border-emerald-300 bg-emerald-50 text-emerald-800 hover:bg-emerald-100"}`}>اعتماد الحركة</button>
                            <button type="button" onClick={() => setReviewDecisions(current => ({ ...current, [item.key]: "reject" }))} className={`rounded-lg px-3 py-2 font-semibold ${decision === "reject" ? "bg-rose-700 text-white" : "border border-rose-300 bg-rose-50 text-rose-800 hover:bg-rose-100"}`}>رفض الحركة</button>
                          </div>
                        </div>
                      })}
                    </div>
                  </div>
                )}
                {parsedPreview.invalidCount > 0 && (
                  <div>
                    <h4 className="mb-2 text-xs font-bold text-slate-800">الصفوف غير الصالحة ({parsedPreview.invalidRows.length.toLocaleString("ar")})</h4>
                    <div className="max-h-64 overflow-auto rounded-lg border border-amber-200 bg-white">
                      <table className="w-full min-w-[560px] text-right text-xs">
                        <thead className="bg-slate-100 text-slate-700"><tr><th className="px-3 py-2">رقم الصف</th><th className="px-3 py-2">سبب الرفض</th><th className="px-3 py-2">بيانات الصف</th></tr></thead>
                        <tbody>{parsedPreview.invalidRows.map(item => (
                          <tr key={item.rowNumber} className="border-t">
                            <td className="px-3 py-2">{item.rowNumber.toLocaleString("ar")}</td>
                            <td className="px-3 py-2 text-rose-800">{item.reason}</td>
                            <td className="max-w-[32rem] truncate px-3 py-2" title={item.values.map(value => String(value ?? "")).join(" | ")}>{item.values.map(value => String(value ?? "")).filter(Boolean).join(" · ") || "—"}</td>
                          </tr>
                        ))}</tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            )}
            <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white">
              <table className="w-full min-w-[600px] text-right text-xs">
                <thead className="bg-slate-100 text-slate-600"><tr>{sheet.headers.map((header, index) => <th key={`${header}-${index}`} className="px-3 py-2 font-semibold">{header}</th>)}</tr></thead>
                <tbody>{sheet.rows.slice(0, 3).map((row, rowIndex) => (
                  <tr key={rowIndex} className="border-t border-slate-100">
                    {sheet.headers.map((_, columnIndex) => <td key={columnIndex} className="max-w-64 truncate px-3 py-2">{String(row[columnIndex] ?? "")}</td>)}
                  </tr>
                ))}</tbody>
              </table>
            </div>
            <button type="button" disabled={busy || searching || pendingReviewCount > 0 || loadedKey !== storageKey} onClick={() => void importFile()} className="inline-flex items-center gap-2 rounded-xl bg-cyan-700 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-cyan-800 disabled:cursor-not-allowed disabled:opacity-50">
              <FileSpreadsheet className="h-4 w-4" /> استيراد الحوالات ومنع التكرار
            </button>
          </div>
        )}
      </section>

      <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
        <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
          <div>
            <h2 className="text-lg font-bold text-slate-900">سجل الحوالات</h2>
            <p className="mt-1 text-xs text-slate-500">ابحث بكلمة أو أكثر من الاسم بأي ترتيب، بالعربية أو الإنجليزية.</p>
          </div>
          <div className="flex flex-wrap gap-2">
            <button type="button" disabled={busy || (!data.records.length && !data.imports.length)} onClick={() => void clearAllTransfers()} className="inline-flex items-center gap-2 rounded-lg border border-rose-200 px-3 py-2 text-xs font-semibold text-rose-700 hover:bg-rose-50 disabled:opacity-50"><Trash2 className="h-4 w-4" />تنظيف وحذف الكل</button>
            <button type="button" onClick={resetFilters} className="rounded-lg border border-slate-200 px-3 py-2 text-xs font-semibold text-slate-700 hover:bg-slate-50">مسح الفلاتر</button>
          </div>
        </div>
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
          <label className={`${labelClass()} sm:col-span-2`}>
            <span>بحث بالاسم أو الرقم المرجعي أو التاريخ أو المبلغ</span>
            <div className="relative">
              <Search className="absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
              <input className={`${inputClass()} pr-9`} value={search} onChange={event => setSearch(event.target.value)} onKeyDown={event => { if (event.key === "Enter") applyFilters() }} placeholder="ابحث بكلمة من الاسم أو أكثر، بأي ترتيب..." />
            </div>
          </label>
          <label className={labelClass()}>
            <span>المصدر</span>
            <select className={inputClass()} value={sourceFilter} onChange={event => setSourceFilter(event.target.value as TransferSource | "all")}>
              <option value="all">جميع المصادر</option>
              {SOURCES.map(item => <option key={item.value} value={item.value}>{item.label}</option>)}
            </select>
          </label>
          <label className={labelClass()}>
            <span>الحساب / اسم الكشف</span>
            <select className={inputClass()} value={accountFilter} onChange={event => setAccountFilter(event.target.value)}>
              <option value="all">جميع الحسابات</option>
              {accountOptions.map(account => <option key={account} value={account}>{account}</option>)}
            </select>
          </label>
          <label className={labelClass()}>
            <span>اتجاه الحركة</span>
            <select className={inputClass()} value={directionFilter} onChange={event => setDirectionFilter(event.target.value as TransferRecord["direction"] | "all")}>
              <option value="all">الوارد والصادر</option>
              <option value="incoming">الوارد فقط</option>
              <option value="outgoing">الصادر فقط</option>
            </select>
          </label>
          <label className={labelClass()}>
            <span>الشهر</span>
            <select className={inputClass()} value={monthFilter} onChange={event => setMonthFilter(event.target.value)}>
              <option value="all">جميع الأشهر</option>
              {monthOptions.map(month => <option key={month} value={month}>{formatMonthKey(month)}</option>)}
            </select>
          </label>
          <label className={labelClass()}>
            <span>ترتيب النتائج</span>
            <select className={inputClass()} value={sortOrder} onChange={event => setSortOrder(event.target.value as typeof sortOrder)}>
              <option value="date-asc">التاريخ: الأقدم أولاً</option>
              <option value="date-desc">التاريخ: الأحدث أولاً</option>
              <option value="amount-asc">المبلغ: الأقل أولاً</option>
              <option value="amount-desc">المبلغ: الأعلى أولاً</option>
            </select>
          </label>
          <label className={labelClass()}><span>من تاريخ</span><input className={inputClass()} type="date" value={dateFrom} onChange={event => setDateFrom(event.target.value)} /></label>
          <label className={labelClass()}><span>إلى تاريخ</span><input className={inputClass()} type="date" value={dateTo} onChange={event => setDateTo(event.target.value)} /></label>
          <label className={labelClass()}><span>المبلغ من</span><input className={inputClass()} inputMode="decimal" value={minimum} onChange={event => setMinimum(event.target.value)} placeholder="مثال: 10" /></label>
          <label className={labelClass()}><span>المبلغ إلى</span><input className={inputClass()} inputMode="decimal" value={maximum} onChange={event => setMaximum(event.target.value)} placeholder="مثال: 500" /></label>
        </div>

        <div className="mt-5 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500">
          <div className="flex items-center gap-3">
            <button type="button" disabled={searching} onClick={applyFilters} className="rounded-lg bg-cyan-800 px-4 py-2 font-semibold text-white hover:bg-cyan-900 disabled:opacity-50"><Search className="ml-1 inline h-4 w-4" />بحث</button>
            <span>{searching ? "جارٍ البحث على دفعات..." : filtersDirty ? "عدّلت الفلاتر؛ اضغط بحث لتطبيقها." : `النتائج: ${filteredRecords.length.toLocaleString("ar")} حوالة`}</span>
          </div>
          <div className="flex flex-wrap gap-2">
            <button type="button" disabled={searching || filtersDirty || !filteredRecords.length} onClick={exportFilteredRecords} className="inline-flex items-center gap-2 rounded-lg border border-emerald-700 px-3 py-2 font-semibold text-emerald-800 disabled:opacity-50"><FileSpreadsheet className="h-4 w-4" />تنزيل Excel</button>
            <button type="button" disabled={searching || filtersDirty || !filteredRecords.length} onClick={() => setPrintRows(filteredRecords)} className="inline-flex items-center gap-2 rounded-lg bg-cyan-800 px-3 py-2 font-semibold text-white disabled:opacity-50"><Printer className="h-4 w-4" />طباعة الكشف المفلتر ({filteredRecords.length.toLocaleString("ar")})</button>
          </div>
        </div>
        <div className="mt-3 overflow-x-auto rounded-xl border border-slate-200">
          <table className="w-full min-w-[760px] text-right text-sm">
            <thead className="bg-slate-50 text-xs font-semibold text-slate-600">
              <tr>
                <th className="px-3 py-3">التاريخ</th><th className="px-3 py-3">الحساب / الكشف</th><th className="px-3 py-3">الاسم</th>
                <th className="px-3 py-3">الرقم المرجعي</th><th className="px-3 py-3">الاتجاه</th><th className="px-3 py-3">المبلغ</th>
              </tr>
            </thead>
            <tbody>
              {visibleRecords.map(record => <TransferTableRow key={record.id} record={record} />)}
              {visibleRecords.length === 0 && <tr><td colSpan={6} className="px-4 py-12 text-center text-sm text-slate-500">لا توجد حوالات مطابقة للفلاتر المحددة.</td></tr>}
            </tbody>
          </table>
        </div>
        <div className="mt-4 flex items-center justify-between gap-3">
          <button type="button" disabled={page === 0} onClick={() => setPage(current => Math.max(0, current - 1))} className="inline-flex items-center gap-1 rounded-lg border border-slate-200 px-3 py-2 text-xs font-semibold disabled:opacity-40"><ChevronRight className="h-4 w-4" /> السابق</button>
          <span className="text-xs text-slate-500">صفحة {Math.min(page + 1, pageCount).toLocaleString("ar")} من {pageCount.toLocaleString("ar")}</span>
          <button type="button" disabled={page + 1 >= pageCount} onClick={() => setPage(current => Math.min(pageCount - 1, current + 1))} className="inline-flex items-center gap-1 rounded-lg border border-slate-200 px-3 py-2 text-xs font-semibold disabled:opacity-40">التالي <ChevronLeft className="h-4 w-4" /></button>
        </div>
      </section>

      {data.imports.length > 0 && (
        <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
          <h2 className="mb-3 text-base font-bold text-slate-900">سجل ملفات الاستيراد</h2>
          <div className="divide-y divide-slate-100">
            {data.imports.map(item => (
              <div key={item.id} className="flex flex-wrap items-center justify-between gap-2 py-2.5 text-xs">
                <div className="min-w-0"><span className="font-semibold text-slate-800">{item.accountName ? `${item.accountName} - ${item.currency || "عملة غير محددة"}` : item.fileName}</span><span className="mr-2 text-slate-500">· {item.fileName} · {sourceLabel(item.source)}</span></div>
                <div className="flex items-center gap-3 text-slate-500">
                  <span>{new Date(item.importedAt).toLocaleString("ar")} · أضيفت {item.importedCount.toLocaleString("ar")} · لم تُستورد {item.duplicateCount.toLocaleString("ar")}</span>
                  <button type="button" disabled={busy} onClick={() => void deleteImport(item)} title="حذف الملف والحوالات المرتبطة به" className="rounded-lg p-2 text-rose-700 hover:bg-rose-50 disabled:opacity-50"><Trash2 className="h-4 w-4" /></button>
                </div>
              </div>
            ))}
          </div>
        </section>
      )}
      {printRows && <section className="transfer-print-report hidden bg-white p-4 text-black" dir="rtl">
        <div className="report-print-logo"><img src={TASWIYAH_LOGO_PATH} alt="شعار تسوية هب" /></div>
        <header className="mb-4 text-center">
          <h1 className="text-lg font-bold">كشف الحوالات المفلتر</h1>
          <p className="mt-1 text-xs">{printRows.length.toLocaleString("ar")} حوالة · تاريخ الطباعة {new Date().toLocaleDateString("ar")}</p>
        </header>
        <div className="mb-3 grid grid-cols-2 gap-2 text-xs">
          <p>إجمالي الوارد: {formatCurrencyTotals(sumCurrencies(printRows, "incoming"))}</p>
          <p>إجمالي الصادر: {formatCurrencyTotals(sumCurrencies(printRows, "outgoing"))}</p>
        </div>
        <table className="w-full border-collapse text-right text-xs">
          <thead><tr>{["التاريخ", "الحساب / الكشف", "الاسم", "الرقم المرجعي", "الاتجاه", "المبلغ"].map(label => <th key={label} className="border border-slate-500 bg-white px-2 py-1.5">{label}</th>)}</tr></thead>
          <tbody>{printRows.map(record => <tr key={record.id}>
            <td className="border border-slate-400 px-2 py-1.5">{record.date}</td>
            <td className="border border-slate-400 px-2 py-1.5">{transferAccountLabel(record)}</td>
            <td className="border border-slate-400 px-2 py-1.5">{record.name || "—"}</td>
            <td className="border border-slate-400 px-2 py-1.5">{record.reference || "—"}</td>
            <td className="border border-slate-400 px-2 py-1.5">{record.direction === "incoming" ? "وارد" : "صادر"}</td>
            <td className="border border-slate-400 px-2 py-1.5">{formatAmount(record.amount)} {record.currency}</td>
          </tr>)}</tbody>
        </table>
      </section>}
    </div>
  )
}

function TransferTableRow({ record }: { record: TransferRecord }) {
  const incoming = record.direction === "incoming"
  return (
    <tr className="border-t border-slate-100 align-top hover:bg-slate-50">
      <td className="whitespace-nowrap px-3 py-3 text-xs text-slate-700">{record.date}</td>
      <td className="whitespace-nowrap px-3 py-3 text-xs text-slate-600">{transferAccountLabel(record)}</td>
      <td className="max-w-48 px-3 py-3 font-medium text-slate-900">{record.name || "—"}</td>
      <td className="whitespace-nowrap px-3 py-3 font-mono text-xs text-slate-700">{record.reference || "—"}</td>
      <td className={`whitespace-nowrap px-3 py-3 text-xs font-semibold ${incoming ? "text-emerald-700" : "text-rose-700"}`}>{incoming ? "وارد" : "صادر"}</td>
      <td className={`whitespace-nowrap px-3 py-3 font-semibold ${incoming ? "text-emerald-800" : "text-rose-800"}`}>{formatAmount(record.amount)} {record.currency}</td>
    </tr>
  )
}
