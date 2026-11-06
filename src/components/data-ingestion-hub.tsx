"use client"

import { useCallback, useEffect, useMemo, useState } from "react"
import * as XLSX from "xlsx"
import { ChevronLeft, ChevronRight, Download, File, FileSpreadsheet, FolderOpen, Pencil, Plus, RotateCw, Table, Trash2, Upload, WalletCards, X } from "lucide-react"
import {
  centralStorageDelete,
  centralStorageGet,
  centralStorageSet,
  deleteStoredFile,
  downloadStoredFile,
  fetchStoredFile,
  listStoredFiles,
  uploadStoredFile,
  type StoredFile,
} from "../lib/central-storage"
import { classifyExpenseByLabel, parseIngestionFile, upgradeParsedData, type DataCategory, type ExpenseSubCategory, type ParsedFileData } from "../lib/data-parser"
import {
  INGESTION_METADATA_KEY,
  INGESTION_PARSED_DATA_KEY_PREFIX,
  type IngestionFileMetadata,
} from "../lib/data-ingestion"
import { compareTableValues, formatMonthKey, getMonthKey, getTableDateColumn } from "../lib/table-view"
import { useGlobalLoading } from "./global-loading"

type IngestionFile = IngestionFileMetadata

const MAX_FILE_SIZE = 25 * 1024 * 1024
const PREVIEW_PAGE_SIZE = 100

const categories: Array<{
  id: DataCategory
  title: string
  description: string
  icon: typeof FileSpreadsheet
  color: string
}> = [
  {
    id: "SALES_AND_CATEGORIES",
    title: "المبيعات وربحية الأقسام",
    description: "ملفات المبيعات العامة والكاشير وأرباح الأقسام",
    icon: FileSpreadsheet,
    color: "text-blue-700 bg-blue-50",
  },
  {
    id: "ELECTRONIC_WALLETS",
    title: "المحافظ والتطبيقات الإلكترونية",
    description: "كشوفات Jawwal Pay والتطبيقات والمحافظ",
    icon: WalletCards,
    color: "text-violet-700 bg-violet-50",
  },
  {
    id: "OPERATING_EXPENSES",
    title: "المصاريف التشغيلية المفرّقة",
    description: "مصروفات التغليف والعمال والوقود والصيانة",
    icon: FolderOpen,
    color: "text-amber-700 bg-amber-50",
  },
  {
    id: "DEBTS_AND_LEDGERS",
    title: "حركة الديون والتحصيلات",
    description: "دفاتر الزبائن وفواتير الموردين والمدفوعات",
    icon: File,
    color: "text-emerald-700 bg-emerald-50",
  },
]

const subCategoryLabels: Record<ExpenseSubCategory, string> = {
  packaging_supplies: "التغليف والمستهلكات",
  staff_meals: "طعام العمال والضيافة",
  fuel_and_energy: "الوقود والطاقة",
  maintenance: "الصيانة",
  general_expenses: "مصاريف عامة",
}

async function saveIngestionIndex(files: IngestionFile[]) {
  const index = files.map(({ parsedData: _parsedData, ...metadata }) => metadata)
  await centralStorageSet(INGESTION_METADATA_KEY, index)
}

function fileExtension(fileName: string): IngestionFile["fileType"] | null {
  const extension = fileName.split(".").pop()?.toLocaleLowerCase("en")
  return extension === "xlsx" || extension === "csv" || extension === "pdf" ? extension : null
}

function formatBytes(size: number) {
  if (size < 1024) return `${size} بايت`
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} كيلوبايت`
  return `${(size / (1024 * 1024)).toFixed(1)} ميغابايت`
}

export function DataIngestionHub() {
  const [files, setFiles] = useState<IngestionFile[]>([])
  const [storedFiles, setStoredFiles] = useState<StoredFile[]>([])
  const [busy, setBusy] = useState(false)
  const [progress, setProgress] = useState("")
  const [loading, setLoading] = useState(true)
  const [modalOpen, setModalOpen] = useState(false)
  const [activeCategory, setActiveCategory] = useState<DataCategory>("SALES_AND_CATEGORIES")
  const [editingFile, setEditingFile] = useState<IngestionFile | null>(null)
  const [customLabel, setCustomLabel] = useState("")
  const [selectedFile, setSelectedFile] = useState<File | null>(null)
  const [message, setMessage] = useState("")
  const [error, setError] = useState("")
  const [previewFileId, setPreviewFileId] = useState<string | null>(null)
  const [previewMonth, setPreviewMonth] = useState("all")
  const [previewSortColumn, setPreviewSortColumn] = useState("")
  const [previewSortDescending, setPreviewSortDescending] = useState(false)
  const [previewSearch, setPreviewSearch] = useState("")
  const [previewPage, setPreviewPage] = useState(0)

  useGlobalLoading(loading || busy, loading ? "جاري تحميل ملفات مركز الرفع..." : "جاري معالجة الملفات...", progress || undefined)

  const storedFilesById = useMemo(() => new Map(storedFiles.map(file => [file.id, file])), [storedFiles])
  const previewFile = files.find(file => file.fileId === previewFileId)
  const previewRows = previewFile?.parsedData?.rawParsedRows ?? []
  const previewHeaders = useMemo(() => Array.from(new Set(previewRows.flatMap(row => Object.keys(row)))), [previewRows])
  const previewDateColumn = getTableDateColumn(previewHeaders)
  const activePreviewSortColumn = previewSortColumn || previewHeaders[previewDateColumn] || ""
  const previewMonths = useMemo(() => {
    const months = new Set<string>()
    let hasUndated = false
    for (const row of previewRows) {
      const month = previewDateColumn < 0 ? null : getMonthKey(row[previewHeaders[previewDateColumn]])
      if (month) months.add(month)
      else hasUndated = true
    }
    return [...Array.from(months).sort(), ...(hasUndated ? ["undated"] : [])]
  }, [previewDateColumn, previewHeaders, previewRows])
  const filteredPreviewRows = useMemo(() => {
    const query = previewSearch.trim().toLocaleLowerCase("ar")
    const rows = previewRows.filter(row => {
      const rowMonth = previewDateColumn < 0 ? null : getMonthKey(row[previewHeaders[previewDateColumn]])
      const matchesMonth = previewMonth === "all" || (previewMonth === "undated" ? !rowMonth : rowMonth === previewMonth)
      return matchesMonth && (!query || previewHeaders.some(header =>
        String(row[header] ?? "").toLocaleLowerCase("ar").includes(query),
      ))
    })
    const column = previewHeaders.indexOf(activePreviewSortColumn)
    if (column < 0) return rows
    const header = previewHeaders[column]
    return rows.sort((left, right) =>
      compareTableValues(left[header], right[header], column === previewDateColumn) * (previewSortDescending ? -1 : 1),
    )
  }, [activePreviewSortColumn, previewDateColumn, previewHeaders, previewMonth, previewRows, previewSearch, previewSortDescending])
  const previewPageCount = Math.max(1, Math.ceil(filteredPreviewRows.length / PREVIEW_PAGE_SIZE))
  const pagePreviewRows = filteredPreviewRows.slice(previewPage * PREVIEW_PAGE_SIZE, (previewPage + 1) * PREVIEW_PAGE_SIZE)

  const refresh = useCallback(async () => {
    setLoading(true)
    setError("")
    try {
      const [savedFiles, savedMetadata] = await Promise.all([
        listStoredFiles(),
        centralStorageGet<IngestionFile[]>(INGESTION_METADATA_KEY),
      ])
      const metadata = Array.isArray(savedMetadata) ? savedMetadata : []
      const migratedInlineData = metadata.some(file => !!file.parsedData)
      const hydratedFiles = await Promise.all(metadata.map(async file => {
        const savedData = file.parsedData
          ? upgradeParsedData(file.parsedData)
          : upgradeParsedData(await centralStorageGet<unknown>(`${INGESTION_PARSED_DATA_KEY_PREFIX}${file.fileId}`))
        return savedData ? { ...file, parsedData: savedData } : { ...file, parsedData: undefined }
      }))
      if (migratedInlineData) {
        await Promise.all(hydratedFiles.flatMap(file => file.parsedData
          ? [centralStorageSet(`${INGESTION_PARSED_DATA_KEY_PREFIX}${file.fileId}`, file.parsedData)]
          : []))
        await saveIngestionIndex(hydratedFiles)
      }
      setStoredFiles(savedFiles)
      setFiles(hydratedFiles)
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذر تحميل ملفات مركز الرفع.")
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => { void refresh() }, [refresh])

  const openAddDialog = (category: DataCategory) => {
    setActiveCategory(category)
    setEditingFile(null)
    setCustomLabel("")
    setSelectedFile(null)
    setError("")
    setMessage("")
    setProgress("")
    setModalOpen(true)
  }

  const openEditDialog = (file: IngestionFile) => {
    setActiveCategory(file.categoryType)
    setEditingFile(file)
    setCustomLabel(file.customLabel)
    setSelectedFile(null)
    setError("")
    setMessage("")
    setProgress("")
    setModalOpen(true)
  }

  const closeDialog = () => {
    if (busy) return
    setModalOpen(false)
    setEditingFile(null)
    setSelectedFile(null)
    setProgress("")
  }

  const saveFile = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    const label = customLabel.trim()
    if (!label) {
      setError("أدخل اسماً توضيحياً للملف.")
      return
    }
    if (!editingFile && !selectedFile) {
      setError("اختر ملف Excel أو CSV أو PDF لإرفاقه.")
      return
    }

    setBusy(true)
    setError("")
    setMessage("")
    let uploadedFileName: string | null = null
    try {
      if (editingFile) {
        let parsedData = editingFile.parsedData
        if (activeCategory !== editingFile.categoryType || label !== editingFile.customLabel) {
          const stored = storedFilesById.get(editingFile.fileId)
          if (!stored) throw new Error(`الملف «${editingFile.customLabel}» غير موجود في مكتبة الملفات.`)
          setProgress("جارٍ إعادة تحليل الملف بعد تعديل بياناته...")
          const source = await fetchStoredFile(editingFile.fileId, stored.fileName, stored.mimeType)
          parsedData = await parseIngestionFile(source, activeCategory, label, setProgress)
          await centralStorageSet(`${INGESTION_PARSED_DATA_KEY_PREFIX}${editingFile.fileId}`, parsedData)
        }
        const updated = files.map(file => file.fileId === editingFile.fileId
          ? {
              ...file,
              categoryType: activeCategory,
              customLabel: label,
              subCategory: activeCategory === "OPERATING_EXPENSES" ? classifyExpenseByLabel(label) : null,
              parsedData,
            }
          : file)
        await saveIngestionIndex(updated)
        setFiles(updated)
        setMessage("تم تحديث اسم الملف وتصنيفه.")
      } else if (selectedFile) {
        const extension = fileExtension(selectedFile.name)
        if (!extension) throw new Error("الصيغ المدعومة هي XLSX وCSV وPDF فقط.")
        if (selectedFile.size > MAX_FILE_SIZE) throw new Error("الحد الأقصى لحجم الملف 25 ميغابايت.")
        setProgress("جارٍ تحليل محتوى الملف...")
        const parsedData = await parseIngestionFile(selectedFile, activeCategory, label, setProgress)
        setProgress("اكتمل التحليل، جارٍ حفظ الملف...")
        const fileId = await uploadStoredFile(selectedFile)
        uploadedFileName = selectedFile.name
        const newFile: IngestionFile = {
          fileId,
          categoryType: activeCategory,
          customLabel: label,
          fileName: selectedFile.name,
          fileType: extension,
          uploadTimestamp: new Date().toISOString(),
          subCategory: activeCategory === "OPERATING_EXPENSES" ? classifyExpenseByLabel(label) : null,
          parsedData,
        }
        const updated = [newFile, ...files]
        await centralStorageSet(`${INGESTION_PARSED_DATA_KEY_PREFIX}${fileId}`, parsedData)
        await saveIngestionIndex(updated)
        uploadedFileName = null
        setFiles(updated)
        setStoredFiles(current => [{
          id: fileId,
          fileName: selectedFile.name,
          mimeType: selectedFile.type || "application/octet-stream",
          size: selectedFile.size,
          uploadedAt: newFile.uploadTimestamp,
        }, ...current])
        setMessage(parsedData.amountDetected
          ? `تم رفع الملف وتحليل ${parsedData.summary.recordCount.toLocaleString("ar")} سطراً.`
          : "تم رفع الملف واستخراج بياناته؛ لم يتم تحديد عمود مبلغ بثقة.")
      }
      setModalOpen(false)
      setEditingFile(null)
      setSelectedFile(null)
    } catch (cause) {
      const detail = cause instanceof Error ? cause.message : "تعذر حفظ الملف."
      setError(uploadedFileName
        ? `تم رفع «${uploadedFileName}»، لكن تعذر حفظ بيانات تصنيفه: ${detail}`
        : detail)
    } finally {
      setBusy(false)
      setProgress("")
    }
  }

  const reprocessFile = async (file: IngestionFile) => {
    setBusy(true)
    setError("")
    setMessage("")
    setProgress(`جارٍ استرجاع «${file.customLabel}»...`)
    try {
      const stored = storedFilesById.get(file.fileId)
      if (!stored) throw new Error(`الملف «${file.customLabel}» غير موجود في مكتبة الملفات.`)
      const source = await fetchStoredFile(file.fileId, stored.fileName, stored.mimeType)
      const parsedData = await parseIngestionFile(source, file.categoryType, file.customLabel, setProgress)
      const updated = files.map(item => item.fileId === file.fileId ? { ...item, parsedData } : item)
      await centralStorageSet(`${INGESTION_PARSED_DATA_KEY_PREFIX}${file.fileId}`, parsedData)
      setFiles(updated)
      setMessage(parsedData.amountDetected
        ? `اكتمل تحليل «${file.customLabel}»: ${parsedData.summary.recordCount.toLocaleString("ar")} سطراً.`
        : `اكتمل استخراج «${file.customLabel}»، لكن لم يتم تحديد مبلغ بثقة.`)
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : `تعذر تحليل «${file.customLabel}».`)
    } finally {
      setBusy(false)
      setProgress("")
    }
  }

  const removeFile = async (file: IngestionFile) => {
    if (!window.confirm(`هل تريد حذف «${file.customLabel}» نهائياً؟`)) return
    setError("")
    setMessage("")
    try {
      const updated = files.filter(item => item.fileId !== file.fileId)
      await deleteStoredFile(file.fileId)
      await saveIngestionIndex(updated)
      await centralStorageDelete(`${INGESTION_PARSED_DATA_KEY_PREFIX}${file.fileId}`)
      setFiles(updated)
      setStoredFiles(current => current.filter(item => item.id !== file.fileId))
      setMessage("تم حذف الملف.")
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذر حذف الملف.")
    }
  }

  const downloadFile = async (file: IngestionFile) => {
    const storedFile = storedFilesById.get(file.fileId)
    if (!storedFile) {
      setError(`الملف «${file.customLabel}» غير موجود في مكتبة الملفات.`)
      return
    }
    setError("")
    try {
      await downloadStoredFile(storedFile)
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذر تنزيل الملف.")
    }
  }

  const openDataPreview = (file: IngestionFile) => {
    const rows = file.parsedData?.rawParsedRows ?? []
    const headers = Array.from(new Set(rows.flatMap(row => Object.keys(row))))
    const dateColumn = getTableDateColumn(headers)
    const months = new Set<string>()
    let hasUndated = false
    for (const row of rows) {
      const month = dateColumn < 0 ? null : getMonthKey(row[headers[dateColumn]])
      if (month) months.add(month)
      else hasUndated = true
    }
    const options = [...Array.from(months).sort(), ...(hasUndated ? ["undated"] : [])]
    setPreviewFileId(file.fileId)
    setPreviewMonth(options.includes("2026-01") ? "2026-01" : options[0] ?? "all")
    setPreviewSortColumn(headers[getTableDateColumn(headers)] ?? "")
    setPreviewSortDescending(false)
    setPreviewSearch("")
    setPreviewPage(0)
  }

  const changePreviewSort = (header: string) => {
    setPreviewSortDescending(current => previewSortColumn === header ? !current : false)
    setPreviewSortColumn(header)
    setPreviewPage(0)
  }

  const downloadPreview = () => {
    if (!previewHeaders.length || !filteredPreviewRows.length) return
    const monthName = previewMonth === "all" ? "كل-الأشهر" : previewMonth === "undated" ? "بدون-تاريخ" : previewMonth
    const sheet = XLSX.utils.json_to_sheet(filteredPreviewRows, { header: previewHeaders })
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, sheet, "البيانات المعروضة")
    XLSX.writeFile(workbook, `${previewFile?.customLabel ?? "البيانات"}-${monthName}.xlsx`)
  }

  const filesByCategory = (category: DataCategory) => files.filter(file => file.categoryType === category)
  const selectedCategory = categories.find(category => category.id === activeCategory)
  const formatAmount = (amount: number, currency = "ILS") => `${new Intl.NumberFormat("ar", { maximumFractionDigits: 2 }).format(amount)} ${currency}`

  return (
    <section className="mx-auto max-w-7xl p-5 sm:p-8" dir="rtl">
      <header className="mb-7 flex flex-wrap items-center justify-between gap-4">
        <div>
          <p className="text-sm font-semibold text-primary">إدارة ملفات التقارير</p>
          <h1 className="mt-1 text-2xl font-bold sm:text-3xl">مركز رفع البيانات</h1>
          <p className="mt-2 max-w-2xl text-sm leading-6 text-muted-foreground">
            ارفع ملفاتك ضمن أقسامها، وامنح كل ملف اسماً واضحاً لتسهيل التعرف على تصنيفه المالي لاحقاً.
          </p>
        </div>
        <button
          type="button"
          onClick={() => void refresh()}
          disabled={loading}
          className="rounded-xl border bg-card px-4 py-2.5 text-sm font-semibold hover:bg-muted disabled:opacity-50"
        >
          {loading ? "جارٍ التحديث..." : "تحديث الملفات"}
        </button>
      </header>

      {message && <p role="status" className="mb-4 rounded-xl bg-emerald-50 px-4 py-3 text-sm text-emerald-800">{message}</p>}
      {error && <p role="alert" className="mb-4 rounded-xl bg-destructive/10 px-4 py-3 text-sm text-destructive">{error}</p>}

      <div className="grid gap-5 xl:grid-cols-2">
        {categories.map(category => {
          const Icon = category.icon
          const categoryFiles = filesByCategory(category.id)
          return (
            <article key={category.id} className="overflow-hidden rounded-2xl border bg-card shadow-sm">
              <div className="flex flex-wrap items-start justify-between gap-4 border-b p-5">
                <div className="flex min-w-0 items-start gap-3">
                  <span className={`flex h-11 w-11 shrink-0 items-center justify-center rounded-xl ${category.color}`}>
                    <Icon className="h-5 w-5" />
                  </span>
                  <div>
                    <h2 className="font-bold">{category.title}</h2>
                    <p className="mt-1 text-xs leading-5 text-muted-foreground">{category.description}</p>
                  </div>
                </div>
                <span className="rounded-full bg-muted px-2.5 py-1 text-xs font-semibold text-muted-foreground">
                  {categoryFiles.length} ملف
                </span>
              </div>

              <div className="p-4 sm:p-5">
                <button
                  type="button"
                  onClick={() => openAddDialog(category.id)}
                  className="flex w-full items-center justify-center gap-2 rounded-xl border border-dashed border-primary/40 bg-primary/[0.03] px-4 py-3 text-sm font-semibold text-primary transition-colors hover:bg-primary/[0.08]"
                >
                  <Plus className="h-4 w-4" />
                  إضافة ملف إلى هذا القسم
                </button>

                {categoryFiles.length === 0 ? (
                  <p className="py-7 text-center text-sm text-muted-foreground">
                    {loading ? "جارٍ تحميل الملفات..." : "لا توجد ملفات مضافة لهذا القسم بعد."}
                  </p>
                ) : (
                  <ul className="mt-3 divide-y">
                    {categoryFiles.map(file => (
                      <li key={file.fileId} className="flex flex-wrap items-center gap-3 py-4">
                        <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-muted text-primary">
                          <FileSpreadsheet className="h-4 w-4" />
                        </span>
                        <div className="min-w-0 flex-1">
                          <p className="truncate text-sm font-semibold">{file.customLabel}</p>
                          <p className="mt-1 truncate text-xs text-muted-foreground">
                            {file.fileName} · {file.fileType.toUpperCase()}
                            {file.subCategory && ` · ${subCategoryLabels[file.subCategory]}`}
                          </p>
                          {file.parsedData ? (
                            <div className="mt-1 text-[11px]">
                              <p className={file.parsedData.amountDetected ? "text-emerald-700" : "text-amber-700"}>
                                تم تحليل {file.parsedData.summary.recordCount.toLocaleString("ar")} سطراً
                                {file.parsedData.amountDetected && ` · الإجمالي ${formatAmount(file.parsedData.summary.totalAmount, file.parsedData.summary.currency)}`}
                                {file.parsedData.truncated && " · عُرضت عينة من الصفوف"}
                              </p>
                              {file.parsedData.extractionNote && <p className="mt-1 text-amber-700">{file.parsedData.extractionNote}</p>}
                              {file.parsedData.amountColumn && <p className="mt-1 text-muted-foreground">عمود المبلغ: {file.parsedData.amountColumn}</p>}
                              {file.parsedData.details.cashAmount !== undefined && <p className="mt-1 text-muted-foreground">النقد: {formatAmount(file.parsedData.details.cashAmount, file.parsedData.summary.currency)}</p>}
                              {file.parsedData.details.app1Amount !== undefined && <p className="mt-1 text-muted-foreground">التطبيق 1: {formatAmount(file.parsedData.details.app1Amount, file.parsedData.summary.currency)}</p>}
                              {file.parsedData.details.app2Amount !== undefined && <p className="mt-1 text-muted-foreground">التطبيق 2: {formatAmount(file.parsedData.details.app2Amount, file.parsedData.summary.currency)}</p>}
                              {file.parsedData.details.app3Amount !== undefined && <p className="mt-1 text-muted-foreground">التطبيق 3: {formatAmount(file.parsedData.details.app3Amount, file.parsedData.summary.currency)}</p>}
                              {file.parsedData.details.newDebtsAmount !== undefined && <p className="mt-1 text-muted-foreground">الديون الجديدة: {formatAmount(file.parsedData.details.newDebtsAmount, file.parsedData.summary.currency)}</p>}
                              {file.parsedData.details.totalExpenseAmount !== undefined && <p className="mt-1 text-muted-foreground">المصاريف: {formatAmount(file.parsedData.details.totalExpenseAmount, file.parsedData.summary.currency)}</p>}
                              {file.parsedData.details.actualStatementTotal !== undefined && <p className="mt-1 text-muted-foreground">المقبوض في الكشف: {formatAmount(file.parsedData.details.actualStatementTotal, file.parsedData.summary.currency)}</p>}
                              {file.parsedData.details.collectedDebtsAmount !== undefined && <p className="mt-1 text-muted-foreground">تحصيل الديون: {formatAmount(file.parsedData.details.collectedDebtsAmount, file.parsedData.summary.currency)}</p>}
                              {file.parsedData.details.vendorPaymentsMade !== undefined && <p className="mt-1 text-muted-foreground">مدفوعات الموردين: {formatAmount(file.parsedData.details.vendorPaymentsMade, file.parsedData.summary.currency)}</p>}
                              {file.parsedData.details.litersCount !== undefined && <p className="mt-1 text-muted-foreground">اللترات: {new Intl.NumberFormat("ar", { maximumFractionDigits: 2 }).format(file.parsedData.details.litersCount)}{file.parsedData.details.costPerLiter !== undefined && ` · سعر اللتر ${formatAmount(file.parsedData.details.costPerLiter, file.parsedData.summary.currency)}`}</p>}
                              {file.parsedData.categoriesBreakdown?.length ? (
                                <ul className="mt-2 space-y-1 border-t pt-2 text-muted-foreground">
                                  {file.parsedData.categoriesBreakdown.map(item => (
                                    <li key={item.categoryName}>
                                      {item.categoryName}: مبيعات {formatAmount(item.sales, file.parsedData!.summary.currency)} · تكلفة {formatAmount(item.cogs, file.parsedData!.summary.currency)} · ربح {formatAmount(item.profit, file.parsedData!.summary.currency)}
                                    </li>
                                  ))}
                                </ul>
                              ) : null}
                            </div>
                          ) : (
                            <p className="mt-1 text-[11px] text-amber-700">لم يحلل محتوى الملف بعد</p>
                          )}
                        </div>
                        {file.parsedData && <button
                          type="button"
                          onClick={() => openDataPreview(file)}
                          className="rounded-lg p-2 text-primary hover:bg-primary/10"
                          aria-label={`عرض بيانات ${file.customLabel}`}
                          title="عرض البيانات وفرزها حسب الشهر أو العمود"
                        >
                          <Table className="h-4 w-4" />
                        </button>}
                        <button
                          type="button"
                          onClick={() => void reprocessFile(file)}
                          disabled={busy}
                          className="rounded-lg p-2 text-muted-foreground hover:bg-muted hover:text-foreground disabled:opacity-50"
                          aria-label={`تحليل ${file.customLabel}`}
                          title="تحليل أو إعادة تحليل الملف"
                        >
                          <RotateCw className={`h-4 w-4 ${busy ? "animate-spin" : ""}`} />
                        </button>
                        <button
                          type="button"
                          onClick={() => void downloadFile(file)}
                          className="rounded-lg p-2 text-muted-foreground hover:bg-muted hover:text-foreground"
                          aria-label={`تنزيل ${file.customLabel}`}
                          title="تنزيل الملف"
                        >
                          <Download className="h-4 w-4" />
                        </button>
                        <button
                          type="button"
                          onClick={() => openEditDialog(file)}
                          className="rounded-lg p-2 text-muted-foreground hover:bg-muted hover:text-foreground"
                          aria-label={`تعديل ${file.customLabel}`}
                          title="تعديل الاسم أو القسم"
                        >
                          <Pencil className="h-4 w-4" />
                        </button>
                        <button
                          type="button"
                          onClick={() => void removeFile(file)}
                          className="rounded-lg p-2 text-destructive hover:bg-destructive/10"
                          aria-label={`حذف ${file.customLabel}`}
                          title="حذف الملف"
                        >
                          <Trash2 className="h-4 w-4" />
                        </button>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            </article>
          )
        })}
      </div>

      {previewFile && previewFile.parsedData && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center bg-slate-950/60 p-3 sm:p-6" role="presentation" onMouseDown={event => {
          if (event.target === event.currentTarget) setPreviewFileId(null)
        }}>
          <section role="dialog" aria-modal="true" aria-labelledby="data-preview-title" className="flex max-h-full w-full max-w-7xl flex-col overflow-hidden rounded-2xl border bg-card shadow-2xl">
            <header className="flex flex-wrap items-center justify-between gap-3 border-b p-4 sm:p-5">
              <div className="min-w-0">
                <h2 id="data-preview-title" className="truncate text-lg font-bold">بيانات الملف: {previewFile.customLabel}</h2>
                <p className="mt-1 text-xs text-muted-foreground">اختر الشهر للعرض، واضغط عنوان أي عمود للفرز تصاعدياً أو تنازلياً. الفرز يدعم التاريخ والاسم والمبلغ.</p>
              </div>
              <button type="button" onClick={() => setPreviewFileId(null)} aria-label="إغلاق عرض البيانات" className="rounded-lg p-2 text-muted-foreground hover:bg-muted"><X className="h-5 w-5" /></button>
            </header>
            <div className="flex flex-wrap items-end justify-between gap-3 border-b p-4">
              <div className="flex flex-wrap items-end gap-2">
                <label className="text-xs font-semibold">الشهر حسب التاريخ الفعلي
                  <select value={previewMonth} onChange={event => { setPreviewMonth(event.target.value); setPreviewPage(0) }} className="mt-1 block min-w-44 rounded-lg border bg-background px-3 py-2 text-sm">
                    {previewMonths.map(month => <option key={month} value={month}>{month === "undated" ? "حركات بلا تاريخ" : formatMonthKey(month)}</option>)}
                    <option value="all">كل الأشهر</option>
                  </select>
                </label>
                <label className="text-xs font-semibold">بحث بالاسم أو البيانات
                  <input value={previewSearch} onChange={event => { setPreviewSearch(event.target.value); setPreviewPage(0) }} placeholder="اكتب اسم أو مبلغاً..." className="mt-1 block w-56 rounded-lg border bg-background px-3 py-2 text-sm font-normal" />
                </label>
                {previewMonths.length > 1 && previewMonth !== "all" && <div className="flex gap-1">
                  <button type="button" disabled={previewMonths.indexOf(previewMonth) <= 0} onClick={() => { setPreviewMonth(previewMonths[previewMonths.indexOf(previewMonth) - 1]); setPreviewPage(0) }} aria-label="الشهر السابق" className="rounded-lg border p-2 disabled:opacity-40"><ChevronRight className="h-4 w-4" /></button>
                  <button type="button" disabled={previewMonths.indexOf(previewMonth) < 0 || previewMonths.indexOf(previewMonth) >= previewMonths.length - 1} onClick={() => { setPreviewMonth(previewMonths[previewMonths.indexOf(previewMonth) + 1]); setPreviewPage(0) }} aria-label="الشهر التالي" className="rounded-lg border p-2 disabled:opacity-40"><ChevronLeft className="h-4 w-4" /></button>
                </div>}
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-xs text-muted-foreground">{filteredPreviewRows.length.toLocaleString("ar")} سجل</span>
                <button type="button" onClick={downloadPreview} disabled={!filteredPreviewRows.length} className="inline-flex items-center gap-2 rounded-lg bg-emerald-700 px-3 py-2 text-xs font-semibold text-white hover:bg-emerald-800 disabled:opacity-50"><Download className="h-4 w-4" />تنزيل المعروض Excel</button>
              </div>
            </div>
            {previewDateColumn < 0 && <p className="border-b bg-amber-50 px-4 py-2 text-xs text-amber-800">لم يتم العثور على عمود تاريخ في البيانات؛ يمكنك عرض كل السجلات وفرزها حسب أي عمود.</p>}
            <div className="min-h-0 flex-1 overflow-auto">
              {previewHeaders.length ? <table className="w-full min-w-[760px] border-collapse text-right text-xs">
                <thead className="sticky top-0 z-10 bg-muted"><tr>{previewHeaders.map(header => <th key={header} className="whitespace-nowrap border-b px-3 py-3"><button type="button" onClick={() => changePreviewSort(header)} className="inline-flex items-center gap-1 font-semibold hover:text-primary">{header}{previewSortColumn === header && <span>{previewSortDescending ? "↓" : "↑"}</span>}</button></th>)}</tr></thead>
                <tbody>{pagePreviewRows.map((row, rowIndex) => <tr key={`${previewPage}-${rowIndex}`} className="border-b hover:bg-muted/40">{previewHeaders.map(header => <td key={header} className="max-w-96 whitespace-nowrap px-3 py-2">{String(row[header] ?? "")}</td>)}</tr>)}</tbody>
              </table> : <p className="p-10 text-center text-sm text-muted-foreground">لم تُستخرج صفوف جدول من هذا الملف.</p>}
              {previewHeaders.length > 0 && pagePreviewRows.length === 0 && <p className="p-10 text-center text-sm text-muted-foreground">لا توجد سجلات تطابق الشهر أو البحث المحدد.</p>}
            </div>
            <footer className="flex items-center justify-between border-t px-4 py-3 text-xs">
              <span>الصفحة {(previewPage + 1).toLocaleString("ar")} من {previewPageCount.toLocaleString("ar")} · يعرض حتى {PREVIEW_PAGE_SIZE} سجل في الصفحة</span>
              <div className="flex gap-2"><button type="button" disabled={previewPage <= 0} onClick={() => setPreviewPage(page => Math.max(0, page - 1))} className="rounded-lg border px-3 py-1.5 disabled:opacity-40">السابق</button><button type="button" disabled={previewPage >= previewPageCount - 1} onClick={() => setPreviewPage(page => Math.min(previewPageCount - 1, page + 1))} className="rounded-lg border px-3 py-1.5 disabled:opacity-40">التالي</button></div>
            </footer>
          </section>
        </div>
      )}

      {modalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4" role="presentation" onMouseDown={event => {
          if (event.target === event.currentTarget) closeDialog()
        }}>
          <form
            onSubmit={event => void saveFile(event)}
            className="w-full max-w-lg rounded-2xl border bg-card p-5 shadow-xl sm:p-6"
            role="dialog"
            aria-modal="true"
            aria-labelledby="ingestion-dialog-title"
          >
            <div className="mb-5 flex items-start justify-between gap-3">
              <div>
                <h2 id="ingestion-dialog-title" className="text-lg font-bold">
                  {editingFile ? "تعديل بيانات الملف" : `إضافة ملف إلى ${selectedCategory?.title ?? "القسم"}`}
                </h2>
                <p className="mt-1 text-sm text-muted-foreground">اكتب اسماً يصف محتوى الملف لتحديد تصنيفه المالي.</p>
              </div>
              <button type="button" onClick={closeDialog} disabled={busy} aria-label="إغلاق" className="rounded-lg p-2 text-muted-foreground hover:bg-muted">
                <X className="h-4 w-4" />
              </button>
            </div>

            <label className="mb-2 block text-sm font-semibold" htmlFor="ingestion-label">الاسم التوضيحي</label>
            <input
              id="ingestion-label"
              value={customLabel}
              onChange={event => setCustomLabel(event.target.value)}
              maxLength={120}
              placeholder="مثال: مصاريف نيلون وتغليف"
              className="mb-4 h-11 w-full rounded-xl border bg-background px-3 text-sm outline-none focus:border-primary focus:ring-2 focus:ring-primary/20"
              required
            />

            <label className="mb-2 block text-sm font-semibold" htmlFor="ingestion-category">القسم</label>
            <select
              id="ingestion-category"
              value={activeCategory}
              onChange={event => setActiveCategory(event.target.value as DataCategory)}
              className="mb-4 h-11 w-full rounded-xl border bg-background px-3 text-sm outline-none focus:border-primary focus:ring-2 focus:ring-primary/20"
            >
              {categories.map(category => <option key={category.id} value={category.id}>{category.title}</option>)}
            </select>

            {!editingFile && (
              <label
                className="mb-1 flex min-h-32 cursor-pointer flex-col items-center justify-center rounded-xl border border-dashed bg-muted/30 p-4 text-center transition-colors hover:border-primary/50 hover:bg-primary/[0.03]"
                onDragOver={event => event.preventDefault()}
                onDrop={event => {
                  event.preventDefault()
                  setSelectedFile(event.dataTransfer.files[0] ?? null)
                  setError("")
                }}
              >
                <Upload className="mb-2 h-5 w-5 text-primary" />
                <span className="text-sm font-semibold">{selectedFile ? selectedFile.name : "اختر ملفاً أو اسحبه إلى هنا"}</span>
                <span className="mt-1 text-xs text-muted-foreground">Excel XLSX أو CSV أو PDF · حتى 25 ميغابايت</span>
                <input
                  type="file"
                  accept=".xlsx,.csv,.pdf"
                  className="sr-only"
                  onChange={event => {
                    setSelectedFile(event.target.files?.[0] ?? null)
                    setError("")
                  }}
                />
              </label>
            )}
            {progress && <p role="status" className="mt-3 rounded-lg bg-primary/5 px-3 py-2 text-sm text-primary">{progress}</p>}
            {error && <p role="alert" className="mt-3 rounded-lg bg-destructive/10 px-3 py-2 text-sm text-destructive">{error}</p>}
            <div className="mt-5 flex flex-wrap items-center justify-between gap-3">
              {selectedFile && <span className="text-xs text-muted-foreground">{formatBytes(selectedFile.size)}</span>}
              <div className="ms-auto flex gap-2">
                <button type="button" onClick={closeDialog} disabled={busy} className="rounded-xl border px-4 py-2.5 text-sm font-semibold hover:bg-muted">
                  إلغاء
                </button>
                <button type="submit" disabled={busy} className="inline-flex items-center gap-2 rounded-xl bg-primary px-4 py-2.5 text-sm font-semibold text-primary-foreground disabled:opacity-50">
                  {busy ? "جارٍ التحليل والحفظ..." : editingFile ? "حفظ التعديلات" : "تحليل ورفع الملف"}
                </button>
              </div>
            </div>
          </form>
        </div>
      )}
    </section>
  )
}
