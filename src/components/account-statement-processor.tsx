import { useEffect, useMemo, useRef, useState, type ChangeEvent, type DragEvent } from "react"
import * as XLSX from "xlsx"
import { AlertTriangle, ArrowDownToLine, ChevronDown, ChevronLeft, ChevronRight, FileSpreadsheet, Plus, RefreshCw, ShieldCheck, Sparkles, Trash2, Upload, X } from "lucide-react"
import { cleanStatementText, extractStatementRows, isStatementNoiseRow, mergeSelectedStatements, removeStatementBalances, type StatementCell, type StatementFormat } from "../lib/account-statement-processor"
import { compareTableValues, formatMonthKey, getMonthKey, getTableDateColumn, normalizeTableHeader } from "../lib/table-view"
import { useGlobalLoading } from "./global-loading"

type StatementDocument = {
  id: string
  fileName: string
  format: StatementFormat
  headers: string[]
  rows: StatementCell[][]
  selectedColumns: boolean[]
  selectedRows: boolean[]
  pages: number
}

let nextId = 0

function readFile(file: File): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => reader.result instanceof ArrayBuffer ? resolve(reader.result) : reject(new Error("تعذر قراءة الملف."))
    reader.onerror = () => reject(new Error("حدث خطأ أثناء قراءة الملف."))
    reader.readAsArrayBuffer(file)
  })
}

function cellText(value: StatementCell): string {
  if (value instanceof Date) return value.toLocaleDateString("en-GB")
  return value == null ? "" : String(value)
}

async function parseFile(file: File): Promise<StatementDocument> {
  if (!/\.(xlsx|xls|xlsm|csv)$/i.test(file.name)) throw new Error("صيغة الملف غير مدعومة. ارفع XLSX أو XLS أو XLSM أو CSV.")
  const workbook = XLSX.read(await readFile(file), { type: "array", cellDates: true })
  const statement = extractStatementRows(workbook.SheetNames.map((name) => ({
    rows: XLSX.utils.sheet_to_json<StatementCell[]>(workbook.Sheets[name], { header: 1, defval: "", raw: false }),
  })))
  if (!statement) throw new Error("لم أتعرف على ترويسة كشف الحساب في هذا الملف.")
  if (!statement.rows.length) throw new Error("تم التعرف على الأعمدة، لكن لم يتم العثور على حركات.")
  const rows = statement.rows.map((values) => Object.fromEntries(statement.headers.map((header, index) => [header, values[index] ?? ""])))
  const cleaned = removeStatementBalances(statement.headers, rows)
  return {
    id: `${Date.now()}-${nextId++}`,
    fileName: file.name,
    format: statement.format,
    headers: cleaned.headers,
    rows: cleaned.rows.map((row) => cleaned.headers.map((header) => cleanStatementText((row[header] ?? "") as StatementCell))),
    selectedColumns: cleaned.headers.map(() => true),
    selectedRows: cleaned.rows.map(() => true),
    pages: statement.headerBlocks,
  }
}

export function AccountStatementProcessor() {
  const inputRef = useRef<HTMLInputElement>(null)
  const [documents, setDocuments] = useState<StatementDocument[]>([])
  const [expanded, setExpanded] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  const [dragging, setDragging] = useState(false)
  const [errors, setErrors] = useState<string[]>([])
  const [downloaded, setDownloaded] = useState(false)
  const [filterNotice, setFilterNotice] = useState("")
  const [selectedMonth, setSelectedMonth] = useState("all")
  const monthInitialized = useRef(false)
  const [sortColumn, setSortColumn] = useState("")
  const [sortDescending, setSortDescending] = useState(false)
  useGlobalLoading(busy, "جاري قراءة ملفات الكشوف...")
  const merged = useMemo(() => mergeSelectedStatements(documents), [documents])
  const selectedCount = documents.reduce((count, item) => count + item.selectedRows.filter(Boolean).length, 0)
  const monthOptions = useMemo(() => {
    const months = new Set<string>()
    let hasUndatedRows = false
    for (const document of documents) {
      const dateColumn = getTableDateColumn(document.headers)
      if (dateColumn < 0) continue
      for (const row of document.rows) {
        const month = getMonthKey(row[dateColumn])
        if (month) months.add(month)
        else hasUndatedRows = true
      }
    }
    return [...Array.from(months).sort(), ...(hasUndatedRows ? ["undated"] : [])]
  }, [documents])
  useEffect(() => {
    if (!monthInitialized.current && monthOptions.length) {
      monthInitialized.current = true
      setSelectedMonth(monthOptions.includes("2026-01") ? "2026-01" : monthOptions[0])
    } else if (selectedMonth !== "all" && monthOptions.length && !monthOptions.includes(selectedMonth)) {
      setSelectedMonth(monthOptions.includes("2026-01") ? "2026-01" : monthOptions[0])
    }
  }, [monthOptions, selectedMonth])
  const visibleSelectedCount = useMemo(() => documents.reduce((count, document) => {
    const dateColumn = getTableDateColumn(document.headers)
    return count + document.rows.reduce((documentCount, row, index) => {
      const matchesMonth = selectedMonth === "all" || (selectedMonth === "undated"
        ? dateColumn < 0 || !getMonthKey(row[dateColumn])
        : dateColumn >= 0 && getMonthKey(row[dateColumn]) === selectedMonth)
      return documentCount + (matchesMonth && document.selectedRows[index] ? 1 : 0)
    }, 0)
  }, 0), [documents, selectedMonth])

  const mergedDateColumn = getTableDateColumn(merged.headers)
  const activeSortColumn = sortColumn || merged.headers[mergedDateColumn] || ""
  const visibleMergedRows = useMemo(() => {
    const rows = merged.rows.flatMap(row => {
      const matchesMonth = selectedMonth === "all" || (selectedMonth === "undated"
        ? mergedDateColumn < 0 || !getMonthKey(row[mergedDateColumn])
        : mergedDateColumn >= 0 && getMonthKey(row[mergedDateColumn]) === selectedMonth)
      return matchesMonth ? [row] : []
    })
    const column = merged.headers.indexOf(activeSortColumn)
    if (column < 0) return rows
    return rows.sort((left, right) => compareTableValues(left[column], right[column], column === mergedDateColumn) * (sortDescending ? -1 : 1))
  }, [activeSortColumn, merged.headers, merged.rows, mergedDateColumn, selectedMonth, sortDescending])

  const changeSort = (header: string) => {
    setSortDescending(current => activeSortColumn === header ? !current : false)
    setSortColumn(header)
  }

  const selectVisibleRows = (document: StatementDocument, selected: boolean) => {
    const dateColumn = getTableDateColumn(document.headers)
    updateDocument(document.id, item => ({
      ...item,
      selectedRows: item.selectedRows.map((current, index) => {
        const month = getMonthKey(item.rows[index]?.[dateColumn])
        const matchesMonth = selectedMonth === "all" || (selectedMonth === "undated"
          ? dateColumn < 0 || !month
          : dateColumn >= 0 && month === selectedMonth)
        return matchesMonth ? selected : current
      }),
    }))
  }

  const downloadRows = (headers: string[], rows: StatementCell[][], fileName: string) => {
    if (!headers.length || !rows.length) return
    const sheet = XLSX.utils.aoa_to_sheet([headers, ...rows.map((row) => row.map(cleanStatementText))])
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, sheet, "الكشف المنظف")
    XLSX.writeFile(workbook, fileName)
    setDownloaded(true)
  }

  const addFiles = async (fileList: FileList | File[]) => {
    const files = Array.from(fileList)
    if (!files.length) return
    setBusy(true)
    setErrors([])
    setDownloaded(false)
    setFilterNotice("")
    const loaded: StatementDocument[] = []
    const failed: string[] = []
    for (const file of files) {
      try {
        loaded.push(await parseFile(file))
      } catch (error) {
        failed.push(`${file.name}: ${error instanceof Error ? error.message : "فشل تحليل الملف."}`)
      }
    }
    if (loaded.length) {
      setDocuments((current) => [...current, ...loaded])
      setExpanded(loaded.at(-1)?.id ?? null)
    }
    setErrors(failed)
    setBusy(false)
    if (inputRef.current) inputRef.current.value = ""
  }

  const updateDocument = (id: string, updater: (document: StatementDocument) => StatementDocument) => {
    setDownloaded(false)
    setDocuments((current) => current.map((document) => document.id === id ? updater(document) : document))
  }

  const removeDocument = (id: string) => {
    setDocuments((current) => current.filter((document) => document.id !== id))
    setExpanded((current) => current === id ? null : current)
    setDownloaded(false)
  }

  const applyFiltering = () => {
    setDocuments((current) => current.map((document) => {
      const keptColumns = document.headers.map((header, index) => ({ header, index }))
        .filter(({ header }) => !/رصيد|balance/i.test(header))
      const keptRows = document.rows.map((row, index) => ({ row, index }))
        .filter(({ row }) => !isStatementNoiseRow(document.headers, row))
      const headers = keptColumns.map(({ header }) => header)
      return {
        ...document,
        headers,
        rows: keptRows.map(({ row }) => keptColumns.map(({ index }) => cleanStatementText(row[index]))),
        selectedColumns: keptColumns.map(({ index }) => document.selectedColumns[index]),
        selectedRows: keptRows.map(({ index }) => document.selectedRows[index]),
      }
    }))
    setDownloaded(false)
    setFilterNotice("تمت فلترة الأرصدة وحذف الصفوف الفارغة والإجماليات والعناوين غير الخاصة بالحركات.")
  }

  const download = () => {
    const rows = [...merged.rows]
    const column = merged.headers.indexOf(activeSortColumn)
    if (column >= 0) rows.sort((left, right) => compareTableValues(left[column], right[column], column === mergedDateColumn) * (sortDescending ? -1 : 1))
    downloadRows(merged.headers, rows, "كشف-الحسابات-المنظف.xlsx")
  }

  const downloadVisibleMonth = () => {
    if (selectedMonth === "all") {
      download()
      return
    }
    const selectedRows = documents.flatMap(document => {
      const dateColumn = getTableDateColumn(document.headers)
      return document.rows.flatMap((row, index) => {
        if (!document.selectedRows[index]) return []
        const matchesMonth = selectedMonth === "undated"
          ? dateColumn < 0 || !getMonthKey(row[dateColumn])
          : dateColumn >= 0 && getMonthKey(row[dateColumn]) === selectedMonth
        if (!matchesMonth) return []
        const mergedRow = Array.from({ length: merged.headers.length }, () => "" as StatementCell)
        document.headers.forEach((header, columnIndex) => {
          if (!document.selectedColumns[columnIndex]) return
          const target = merged.headers.findIndex(item => normalizeTableHeader(item) === normalizeTableHeader(header))
          if (target >= 0) mergedRow[target] = row[columnIndex] ?? ""
        })
        return [mergedRow]
      })
    })
    const column = merged.headers.indexOf(activeSortColumn)
    if (column >= 0) selectedRows.sort((left, right) => compareTableValues(left[column], right[column], column === mergedDateColumn) * (sortDescending ? -1 : 1))
    const monthName = selectedMonth === "undated" ? "بدون-تاريخ" : selectedMonth
    downloadRows(merged.headers, selectedRows, `كشف-الحسابات-${monthName}.xlsx`)
  }

  return (
    <main dir="rtl" className="min-h-full bg-[#f4f7fb] px-4 py-7 text-slate-900 sm:px-7 lg:px-10">
      <div className="mx-auto max-w-6xl space-y-6">
        <header className="rounded-3xl bg-[#10243c] px-6 py-8 text-white shadow-xl sm:px-9">
          <div className="mb-3 inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/5 px-3 py-1.5 text-xs text-cyan-100"><Sparkles className="h-3.5 w-3.5" />معالجة محلية وآمنة</div>
          <h1 className="text-2xl font-extrabold sm:text-3xl">معالج كشوفات الحسابات والبيانات الذكي</h1>
          <p className="mt-3 max-w-3xl text-sm leading-7 text-slate-300">ارفع كشف المحفظة أو البنك، اختر الصفوف والأعمدة التي تريدها، واجمع أي عدد من الكشوف في ملف Excel واحد.</p>
          <div className="mt-4 flex items-center gap-2 text-xs text-emerald-200"><ShieldCheck className="h-4 w-4" />تتم المعالجة على جهازك فقط</div>
        </header>

        <section className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm sm:p-7">
          <div className="mb-4">
            <h2 className="text-sm font-bold">رفع الكشوفات</h2>
            <p className="mt-1 text-xs leading-6 text-slate-500">يمكن رفع عدة ملفات معاً أو إضافة ملفات لاحقاً بزر +.</p>
          </div>
          <input ref={inputRef} type="file" multiple accept=".xlsx,.xls,.xlsm,.csv" className="hidden" onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files && void addFiles(event.target.files)} />
          <div
            onDragEnter={(event) => { event.preventDefault(); setDragging(true) }}
            onDragOver={(event) => event.preventDefault()}
            onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setDragging(false) }}
            onDrop={(event: DragEvent<HTMLDivElement>) => { event.preventDefault(); setDragging(false); void addFiles(event.dataTransfer.files) }}
            className={`flex min-h-48 flex-col items-center justify-center rounded-2xl border-2 border-dashed px-5 py-8 text-center ${dragging ? "border-cyan-500 bg-cyan-50" : "border-slate-200 bg-slate-50"}`}
          >
            <div className="mb-3 text-cyan-700">{busy ? <RefreshCw className="h-7 w-7 animate-spin" /> : documents.length ? <Plus className="h-7 w-7" /> : <Upload className="h-7 w-7" />}</div>
            <p className="text-sm font-bold">{busy ? "جارٍ قراءة الملفات..." : "اسحب ملفاً أو أكثر وأفلته هنا"}</p>
            <button type="button" disabled={busy} onClick={() => inputRef.current?.click()} className="mt-4 inline-flex items-center gap-2 rounded-xl bg-[#10243c] px-5 py-2.5 text-sm font-semibold text-white disabled:opacity-50">
              {documents.length > 0 && <Plus className="h-4 w-4" />}{documents.length ? "إضافة ملف" : "اختيار ملف"}
            </button>
            <span className="mt-2 text-[11px] text-slate-400">XLSX · XLS · XLSM · CSV</span>
          </div>
          {!!errors.length && <div role="alert" className="mt-4 space-y-1 rounded-xl border border-rose-200 bg-rose-50 p-3 text-sm text-rose-800"><AlertTriangle className="ml-2 inline h-4 w-4" />{errors.map((error) => <p key={error}>{error}</p>)}</div>}
        </section>

        {!!documents.length && <>
          <section className="flex flex-wrap items-center justify-between gap-3 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
            <div>
              <h2 className="text-sm font-bold">عرض الحركات حسب التاريخ الفعلي</h2>
              <p className="mt-1 text-xs text-slate-500">اعرض شهراً واحداً، ثم انتقل للشهر السابق أو التالي. اختر «كل الأشهر» لعرض الكشف كاملاً.</p>
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <button type="button" disabled={!monthOptions.length || selectedMonth === "all" || monthOptions.indexOf(selectedMonth) <= 0} onClick={() => setSelectedMonth(monthOptions[monthOptions.indexOf(selectedMonth) - 1])} aria-label="الشهر السابق" className="rounded-lg border p-2 disabled:opacity-40"><ChevronRight className="h-4 w-4" /></button>
              <select aria-label="شهر العرض" value={selectedMonth} onChange={event => setSelectedMonth(event.target.value)} className="rounded-lg border bg-white px-3 py-2 text-sm font-semibold">
                {monthOptions.map(month => <option key={month} value={month}>{month === "undated" ? "حركات بلا تاريخ فعلي" : formatMonthKey(month)}</option>)}
                <option value="all">كل الأشهر</option>
              </select>
              <button type="button" disabled={!monthOptions.length || selectedMonth === "all" || monthOptions.indexOf(selectedMonth) < 0 || monthOptions.indexOf(selectedMonth) >= monthOptions.length - 1} onClick={() => setSelectedMonth(monthOptions[monthOptions.indexOf(selectedMonth) + 1])} aria-label="الشهر التالي" className="rounded-lg border p-2 disabled:opacity-40"><ChevronLeft className="h-4 w-4" /></button>
              <span className="rounded-lg bg-cyan-50 px-3 py-2 text-xs font-semibold text-cyan-900">{visibleSelectedCount.toLocaleString("ar")} حركة محددة في العرض</span>
            </div>
          </section>
          <section className="space-y-3">
            <div><h2 className="text-lg font-bold">اختر ما تريد الاحتفاظ به من كل ملف</h2><p className="mt-1 text-xs text-slate-500">يُتعرّف على صفحات كشف الأستاذ المتتابعة، وتُستبعد أعمدة الرصيد وأسطر الرصيد السابق والمجموع تلقائياً. المدين والدائن هما مبالغ الحركات.</p></div>
            <div className="flex flex-wrap items-center gap-3">
              <button type="button" onClick={applyFiltering} className="inline-flex items-center gap-2 rounded-xl border border-cyan-200 bg-cyan-50 px-4 py-2 text-sm font-semibold text-cyan-900 hover:bg-cyan-100"><Sparkles className="h-4 w-4" />فلترة الأرصدة والصفوف الفارغة</button>
              {filterNotice && <span role="status" className="text-xs text-emerald-700">{filterNotice}</span>}
            </div>
            {downloaded && <div className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-emerald-200 bg-emerald-50 p-3">
              <span className="text-sm font-medium text-emerald-900">تم تنزيل الملف. يمكنك حذف الكشوفات الحالية والبدء من جديد.</span>
              <button type="button" onClick={() => { setDocuments([]); setExpanded(null); setDownloaded(false) }} className="inline-flex items-center gap-2 rounded-lg bg-rose-700 px-3 py-2 text-xs font-semibold text-white hover:bg-rose-800">
                <Trash2 className="h-4 w-4" />حذف الملفات المرفوعة
              </button>
            </div>}
            {documents.map((document, documentIndex) => {
              const isExpanded = expanded === document.id
              const dateColumn = getTableDateColumn(document.headers)
              const shownRows = document.rows
                .map((row, rowIndex) => ({ row, rowIndex }))
                .filter(({ row }) => selectedMonth === "all" || (selectedMonth === "undated"
                  ? dateColumn < 0 || !getMonthKey(row[dateColumn])
                  : dateColumn >= 0 && getMonthKey(row[dateColumn]) === selectedMonth))
              const documentSortHeader = document.headers.find(header => normalizeTableHeader(header) === normalizeTableHeader(activeSortColumn))
                ?? document.headers[dateColumn]
              const sortedColumn = document.headers.indexOf(documentSortHeader)
              if (sortedColumn >= 0) shownRows.sort((left, right) =>
                compareTableValues(left.row[sortedColumn], right.row[sortedColumn], sortedColumn === dateColumn) * (sortDescending ? -1 : 1))
              return <article key={document.id} className="overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-sm">
                <div className="flex items-center gap-3 px-4 py-4 sm:px-6">
                <button type="button" onClick={() => setExpanded(isExpanded ? null : document.id)} aria-expanded={isExpanded} className="flex min-w-0 flex-1 items-center justify-between gap-4 text-right">
                  <span className="flex min-w-0 items-center gap-3"><FileSpreadsheet className="h-5 w-5 shrink-0 text-cyan-700" /><span className="min-w-0"><span className="block truncate text-sm font-bold">{documentIndex + 1}. {document.fileName}</span><span className="mt-1 block text-xs text-slate-500">{document.format === "bank" ? "كشف بنكي" : document.format === "ledger" ? "كشف محاسبي" : "كشف بعناوين أصلية"} · {document.selectedRows.filter(Boolean).length} من {document.rows.length} حركة محددة</span></span></span>
                  <ChevronDown className={`h-5 w-5 shrink-0 transition-transform ${isExpanded ? "rotate-180" : ""}`} />
                </button>
                {downloaded && <button type="button" aria-label={`حذف ${document.fileName}`} title="حذف هذا الملف" onClick={() => removeDocument(document.id)} className="shrink-0 rounded-lg p-1.5 text-rose-700 hover:bg-rose-50"><X className="h-4 w-4" /></button>}
                </div>
                {isExpanded && <div className="border-t border-slate-100 p-4 sm:p-6">
                  <div className="mb-2 flex flex-wrap items-center justify-between gap-2"><h3 className="text-sm font-bold">الأعمدة</h3><div className="flex gap-2"><button onClick={() => updateDocument(document.id, (item) => ({ ...item, selectedColumns: item.headers.map(() => true) }))} className="rounded-lg border px-3 py-1.5 text-xs">تحديد الكل</button><button onClick={() => updateDocument(document.id, (item) => ({ ...item, selectedColumns: item.headers.map(() => false) }))} className="rounded-lg border px-3 py-1.5 text-xs">إلغاء الكل</button></div></div>
                  <div className="mb-5 flex flex-wrap gap-2">{document.headers.map((header, index) => <label key={`${header}-${index}`} className={`inline-flex cursor-pointer items-center gap-2 rounded-xl border px-3 py-2 text-sm ${document.selectedColumns[index] ? "border-cyan-200 bg-cyan-50" : "border-slate-200 text-slate-500"}`}><input type="checkbox" checked={document.selectedColumns[index]} onChange={() => updateDocument(document.id, (item) => ({ ...item, selectedColumns: item.selectedColumns.map((selected, column) => column === index ? !selected : selected) }))} className="h-4 w-4 accent-cyan-700" />{header}</label>)}</div>
                  <div className="mb-2 flex flex-wrap items-center justify-between gap-2"><h3 className="text-sm font-bold">صفوف الحركات ({shownRows.length} ظاهرة)</h3><div className="flex gap-2"><button onClick={() => selectVisibleRows(document, true)} className="rounded-lg border px-3 py-1.5 text-xs">تحديد الظاهر</button><button onClick={() => selectVisibleRows(document, false)} className="rounded-lg border px-3 py-1.5 text-xs">إلغاء الظاهر</button></div></div>
                  {dateColumn < 0 && selectedMonth !== "all" && <p className="mb-2 rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-800">لم يتم العثور على عمود تاريخ في هذا الملف، لذلك لا تظهر حركاته ضمن اختيار الشهر.</p>}
                  <div className="max-h-[360px] overflow-auto rounded-xl border border-slate-200"><table className="w-full min-w-[760px] border-collapse text-right text-xs"><thead className="sticky top-0 bg-slate-100"><tr><th className="px-3 py-2">إبقاء</th><th className="px-3 py-2">#</th>{document.headers.map((header, index) => <th key={`${header}-${index}`} className="px-3 py-2"><button type="button" onClick={() => changeSort(header)} className="inline-flex items-center gap-1 font-semibold hover:text-cyan-800">{header}{activeSortColumn === header && <span aria-label={sortDescending ? "تنازلي" : "تصاعدي"}>{sortDescending ? "↓" : "↑"}</span>}</button></th>)}</tr></thead><tbody>{shownRows.map(({ row, rowIndex }) => <tr key={rowIndex} className="border-t border-slate-100"><td className="px-3 py-2 text-center"><input type="checkbox" checked={document.selectedRows[rowIndex]} onChange={() => updateDocument(document.id, (item) => ({ ...item, selectedRows: item.selectedRows.map((selected, index) => index === rowIndex ? !selected : selected) }))} aria-label={`إبقاء الصف ${rowIndex + 1}`} /></td><td className="px-3 py-2">{rowIndex + 1}</td>{document.headers.map((header, columnIndex) => <td key={`${header}-${columnIndex}`} className="max-w-80 px-3 py-2">{cellText(row[columnIndex])}</td>)}</tr>)}</tbody></table>{shownRows.length === 0 && <p className="p-6 text-center text-xs text-slate-500">لا توجد حركات في الشهر المحدد.</p>}</div>
                </div>}
              </article>
            })}
          </section>

          <section className="overflow-hidden rounded-3xl border border-slate-200 bg-white shadow-sm">
            <div className="flex flex-wrap items-center justify-between gap-3 border-b p-5"><div><h2 className="font-bold">الملف الموحّد</h2><p className="mt-1 text-xs text-slate-500">{documents.length} ملف · {selectedCount} حركة محفوظة · {visibleMergedRows.length} حركة في العرض · الرصيد مستبعد افتراضياً</p></div><div className="flex flex-wrap gap-2"><button onClick={downloadVisibleMonth} disabled={!visibleSelectedCount || !merged.headers.length} className="inline-flex items-center gap-2 rounded-xl border border-emerald-700 px-4 py-2.5 text-sm font-semibold text-emerald-800 disabled:opacity-50"><ArrowDownToLine className="h-4 w-4" />تنزيل الشهر المعروض</button><button onClick={download} disabled={!merged.rows.length || !merged.headers.length} className="inline-flex items-center gap-2 rounded-xl bg-emerald-700 px-5 py-2.5 text-sm font-semibold text-white disabled:opacity-50"><ArrowDownToLine className="h-4 w-4" />تنزيل كل الأشهر</button></div></div>
            {!!merged.headers.length && <div className="max-h-[420px] overflow-auto"><table className="w-full min-w-[760px] border-collapse text-right text-xs"><thead className="sticky top-0 bg-slate-100"><tr>{merged.headers.map((header, index) => <th key={`${header}-${index}`} className="px-4 py-3"><button type="button" onClick={() => changeSort(header)} className="inline-flex items-center gap-1 font-semibold hover:text-cyan-800">{header}{activeSortColumn === header && <span>{sortDescending ? "↓" : "↑"}</span>}</button></th>)}</tr></thead><tbody>{visibleMergedRows.slice(0, 30).map((row, rowIndex) => <tr key={rowIndex} className="border-t"><>{merged.headers.map((header, index) => <td key={`${header}-${index}`} className="px-4 py-3">{cellText(row[index])}</td>)}</></tr>)}</tbody></table>{visibleMergedRows.length === 0 && <p className="p-6 text-center text-xs text-slate-500">لا توجد حركات مطابقة للشهر المحدد.</p>}</div>}
          </section>
        </>}
      </div>
    </main>
  )
}
