"use client"

import { useEffect, useMemo, useRef, useState, type FormEvent } from "react"
import { Archive, FileText, Fuel, Pencil, Plus, Printer, Search, Trash2, X } from "lucide-react"
import {
  centralStorageEnabled,
  centralStorageGet,
  centralStorageSet,
  deleteStoredFile,
  fetchStoredFile,
  uploadStoredFile,
} from "../lib/central-storage"
import { TASWIYAH_LOGO_PATH } from "../lib/report-branding"
import { useGlobalLoading } from "./global-loading"

type FuelType = "diesel" | "gasoline" | "engine_oil" | "other"

type FuelInvoice = {
  id: string
  invoiceNumber: string
  supplier: string
  date: string
  fuelType: FuelType
  product: string
  quantity: number
  unitPrice: number
  paidBy: string
  receivedBy: string
  location: string
  archiveFileId?: string
  archiveFileName?: string
  archiveMimeType?: string
  archiveDataUrl?: string
}

type FuelWithdrawal = {
  id: string
  invoiceId: string
  date: string
  quantity: number
  withdrawnBy: string
  purpose: string
  beneficiary: string
}

type FuelData = { invoices: FuelInvoice[]; withdrawals: FuelWithdrawal[] }
type InvoiceDraft = Omit<FuelInvoice, "id" | "archiveFileId" | "archiveFileName" | "archiveMimeType" | "archiveDataUrl"> & { id?: string }
type WithdrawalDraft = Omit<FuelWithdrawal, "id"> & { id?: string }
type Tab = "dashboard" | "invoices" | "withdrawals" | "report"

const FUEL_TYPES: Array<{ value: FuelType; label: string }> = [
  { value: "diesel", label: "سولار" },
  { value: "gasoline", label: "بنزين" },
  { value: "engine_oil", label: "زيت محرك" },
  { value: "other", label: "محروقات أخرى" },
]
const EMPTY_DATA: FuelData = { invoices: [], withdrawals: [] }

function createId() {
  return typeof globalThis.crypto?.randomUUID === "function"
    ? globalThis.crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2)}`
}

function emptyInvoiceDraft(): InvoiceDraft {
  const today = new Date().toISOString().slice(0, 10)
  return {
    invoiceNumber: "",
    supplier: "",
    date: today,
    fuelType: "diesel",
    product: "",
    quantity: 0,
    unitPrice: 0,
    paidBy: "",
    receivedBy: "",
    location: "",
  }
}

function emptyWithdrawalDraft(invoiceId = ""): WithdrawalDraft {
  return {
    invoiceId,
    date: new Date().toISOString().slice(0, 10),
    quantity: 0,
    withdrawnBy: "",
    purpose: "",
    beneficiary: "",
  }
}

function toHundredths(value: number) {
  return Math.round((value + Number.EPSILON) * 100)
}

function formatQuantity(value: number) {
  return new Intl.NumberFormat("ar", { minimumFractionDigits: 0, maximumFractionDigits: 2 }).format(value)
}

function formatMoney(value: number) {
  return new Intl.NumberFormat("ar", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value)
}

function fuelLabel(value: FuelType) {
  return FUEL_TYPES.find(type => type.value === value)?.label ?? "محروقات أخرى"
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => typeof reader.result === "string" ? resolve(reader.result) : reject(new Error("تعذر قراءة ملف الفاتورة."))
    reader.onerror = () => reject(reader.error ?? new Error("تعذر قراءة ملف الفاتورة."))
    reader.readAsDataURL(file)
  })
}

function remainingFor(invoice: FuelInvoice, withdrawals: FuelWithdrawal[]) {
  const withdrawn = withdrawals
    .filter(item => item.invoiceId === invoice.id)
    .reduce((sum, item) => sum + toHundredths(item.quantity), 0)
  return Math.max(0, toHundredths(invoice.quantity) - withdrawn) / 100
}

function inputClass() {
  return "w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-normal outline-none transition focus:border-cyan-600 focus:ring-2 focus:ring-cyan-100"
}

function labelClass() {
  return "block space-y-1.5 text-xs font-semibold text-slate-700"
}

export function FuelManagement({ username }: { username: string }) {
  const normalizedUsername = username.normalize("NFKC").trim().toLocaleLowerCase("en-US")
  const storageKey = `fuel_inventory:${normalizedUsername}`
  const [data, setData] = useState<FuelData>(EMPTY_DATA)
  const dataRef = useRef<FuelData>(EMPTY_DATA)
  const busyRef = useRef(false)
  const [loadedKey, setLoadedKey] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState("")
  const [notice, setNotice] = useState("")
  const [tab, setTab] = useState<Tab>("dashboard")
  const [invoiceDraft, setInvoiceDraft] = useState<InvoiceDraft | null>(null)
  const [invoiceFile, setInvoiceFile] = useState<File | null>(null)
  const [withdrawalDraft, setWithdrawalDraft] = useState<WithdrawalDraft | null>(null)
  const [search, setSearch] = useState("")
  const [dateFrom, setDateFrom] = useState("")
  const [dateTo, setDateTo] = useState("")
  const [archiveView, setArchiveView] = useState<{ name: string; url: string; revoke: boolean } | null>(null)

  useGlobalLoading(loadedKey !== storageKey && !error, "جاري تحميل بيانات المحروقات...")
  useGlobalLoading(busy, "جاري معالجة بيانات المحروقات...")

  useEffect(() => {
    let active = true
    setLoadedKey(null)
    setData(EMPTY_DATA)
    dataRef.current = EMPTY_DATA
    setError("")
    const load = async () => {
      try {
        const stored = centralStorageEnabled
          ? await centralStorageGet<FuelData>(storageKey) ?? EMPTY_DATA
          : JSON.parse(localStorage.getItem(storageKey) ?? "null") as FuelData | null ?? EMPTY_DATA
        if (!Array.isArray(stored.invoices) || !Array.isArray(stored.withdrawals)) {
          throw new Error("بيانات منصة المحروقات المحفوظة غير صالحة.")
        }
        if (active) {
          dataRef.current = stored
          setData(stored)
          setLoadedKey(storageKey)
        }
      } catch (cause) {
        if (active) setError(cause instanceof Error ? `تعذر تحميل بيانات المحروقات: ${cause.message}` : "تعذر تحميل بيانات المحروقات.")
      }
    }
    void load()
    return () => { active = false }
  }, [storageKey])

  useEffect(() => () => {
    if (archiveView?.revoke) URL.revokeObjectURL(archiveView.url)
  }, [archiveView])

  const invoicesById = useMemo(() => new Map(data.invoices.map(invoice => [invoice.id, invoice])), [data.invoices])
  const filteredWithdrawals = useMemo(() => data.withdrawals.filter(withdrawal => {
    if (tab === "report" && ((dateFrom && withdrawal.date < dateFrom) || (dateTo && withdrawal.date > dateTo))) return false
    const invoice = invoicesById.get(withdrawal.invoiceId)
    const query = search.trim().toLocaleLowerCase("ar")
    if (!query) return true
    return [withdrawal.withdrawnBy, withdrawal.purpose, withdrawal.beneficiary, invoice?.invoiceNumber, invoice?.supplier, invoice?.product]
      .some(value => value?.toLocaleLowerCase("ar").includes(query))
  }), [data.withdrawals, dateFrom, dateTo, invoicesById, search, tab])

  const totals = useMemo(() => FUEL_TYPES.map(type => {
    const invoices = data.invoices.filter(invoice => invoice.fuelType === type.value)
    const invoiceIds = new Set(invoices.map(invoice => invoice.id))
    const received = invoices.reduce((sum, invoice) => sum + toHundredths(invoice.quantity), 0) / 100
    const withdrawn = data.withdrawals
      .filter(item => invoiceIds.has(item.invoiceId))
      .reduce((sum, item) => sum + toHundredths(item.quantity), 0) / 100
    return { ...type, received, withdrawn, remaining: Math.max(0, toHundredths(received - withdrawn)) / 100 }
  }), [data.invoices, data.withdrawals])

  const commitData = async (next: FuelData, successMessage: string) => {
    try {
      if (centralStorageEnabled) await centralStorageSet(storageKey, next)
      else localStorage.setItem(storageKey, JSON.stringify(next))
      dataRef.current = next
      setData(next)
      setNotice(successMessage)
      setError("")
      return true
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر حفظ البيانات؛ لم يتم اعتماد التعديل: ${cause.message}` : "تعذر حفظ البيانات؛ لم يتم اعتماد التعديل.")
      return false
    }
  }

  const startBusy = () => {
    if (busyRef.current || loadedKey !== storageKey) return false
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

  const nextInvoiceNumber = () => {
    const year = new Date().getFullYear()
    const usedNumbers = new Set(dataRef.current.invoices.map(invoice => invoice.invoiceNumber))
    let index = dataRef.current.invoices.length + 1
    let number = `F-${year}-${String(index).padStart(4, "0")}`
    while (usedNumbers.has(number)) {
      index++
      number = `F-${year}-${String(index).padStart(4, "0")}`
    }
    return number
  }

  const openNewInvoice = () => {
    setInvoiceDraft({ ...emptyInvoiceDraft(), invoiceNumber: nextInvoiceNumber() })
    setInvoiceFile(null)
  }

  const submitInvoice = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!invoiceDraft || !startBusy()) return
    const quantity = Number(invoiceDraft.quantity)
    const unitPrice = Number(invoiceDraft.unitPrice)
    if (!invoiceDraft.invoiceNumber.trim() || !invoiceDraft.supplier.trim() || !invoiceDraft.product.trim()
      || !invoiceDraft.paidBy.trim() || !invoiceDraft.receivedBy.trim() || !invoiceDraft.location.trim()
      || !invoiceDraft.date || !Number.isFinite(quantity) || quantity <= 0
      || !Number.isFinite(unitPrice) || unitPrice < 0 || toHundredths(quantity) <= 0
      || Math.abs(quantity * 100 - toHundredths(quantity)) > 0.000001) {
      setError("أكمل حقول الفاتورة، وأدخل كمية موجبة حتى منزلتين عشريتين وسعراً صالحاً.")
      endBusy()
      return
    }
    const original = invoiceDraft.id ? dataRef.current.invoices.find(invoice => invoice.id === invoiceDraft.id) : undefined
    if (dataRef.current.invoices.some(invoice => invoice.invoiceNumber === invoiceDraft.invoiceNumber.trim() && invoice.id !== invoiceDraft.id)) {
      setError("رقم الفاتورة مستخدم مسبقاً؛ أدخل رقماً مختلفاً.")
      endBusy()
      return
    }
    if (original && toHundredths(quantity) < toHundredths(original.quantity) - dataRef.current.withdrawals
      .filter(item => item.invoiceId === original.id)
      .reduce((sum, item) => sum + toHundredths(item.quantity), 0)) {
      setError("لا يمكن تخفيض كمية الفاتورة إلى أقل من المسحوبات المسجلة عليها.")
      endBusy()
      return
    }
    let uploadedFileId: string | undefined
    let archiveDataUrl: string | undefined
    try {
      if (invoiceFile) {
        if (centralStorageEnabled) uploadedFileId = await uploadStoredFile(invoiceFile, `فاتورة-${invoiceDraft.invoiceNumber}-${invoiceFile.name}`)
        else {
          if (invoiceFile.size > 3 * 1024 * 1024) throw new Error("في التخزين المحلي، يجب ألا يتجاوز ملف الأرشيف 3 ميغابايت.")
          archiveDataUrl = await readFileAsDataUrl(invoiceFile)
        }
      }
      const savedInvoice: FuelInvoice = {
        ...invoiceDraft,
        id: invoiceDraft.id ?? createId(),
        invoiceNumber: invoiceDraft.invoiceNumber.trim(),
        supplier: invoiceDraft.supplier.trim(),
        product: invoiceDraft.product.trim(),
        quantity: toHundredths(quantity) / 100,
        unitPrice,
        paidBy: invoiceDraft.paidBy.trim(),
        receivedBy: invoiceDraft.receivedBy.trim(),
        location: invoiceDraft.location.trim(),
        ...(uploadedFileId ? {
          archiveFileId: uploadedFileId,
          archiveFileName: invoiceFile?.name,
          archiveMimeType: invoiceFile?.type || "application/octet-stream",
          archiveDataUrl: undefined,
        } : original ? {
          archiveFileId: original.archiveFileId,
          archiveFileName: original.archiveFileName,
          archiveMimeType: original.archiveMimeType,
          archiveDataUrl: original.archiveDataUrl,
        } : archiveDataUrl ? {
          archiveFileName: invoiceFile?.name,
          archiveMimeType: invoiceFile?.type,
          archiveDataUrl,
        } : {}),
      }
      const next: FuelData = {
        ...dataRef.current,
        invoices: original
          ? dataRef.current.invoices.map(invoice => invoice.id === original.id ? savedInvoice : invoice)
          : [savedInvoice, ...dataRef.current.invoices],
      }
      const saved = await commitData(next, original ? "تم تحديث الفاتورة." : "تم تسجيل فاتورة الوارد.")
      if (!saved) {
        if (uploadedFileId) await deleteStoredFile(uploadedFileId)
        return
      }
      if (original?.archiveFileId && uploadedFileId) {
        try {
          await deleteStoredFile(original.archiveFileId)
        } catch (cause) {
          setError(`تم حفظ الفاتورة الجديدة، لكن تعذر حذف نسخة الأرشيف القديمة: ${cause instanceof Error ? cause.message : "خطأ غير معروف."}`)
        }
      }
      setInvoiceDraft(null)
      setInvoiceFile(null)
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر حفظ الفاتورة: ${cause.message}` : "تعذر حفظ الفاتورة.")
    } finally {
      endBusy()
    }
  }

  const submitWithdrawal = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!withdrawalDraft || !startBusy()) return
    const quantity = Number(withdrawalDraft.quantity)
    const invoice = dataRef.current.invoices.find(item => item.id === withdrawalDraft.invoiceId)
    if (!invoice || !withdrawalDraft.date || !withdrawalDraft.withdrawnBy.trim()
      || !withdrawalDraft.purpose.trim() || !withdrawalDraft.beneficiary.trim()
      || !Number.isFinite(quantity) || quantity <= 0
      || Math.abs(quantity * 100 - toHundredths(quantity)) > 0.000001) {
      setError("أكمل بيانات المسحوب وأدخل كمية موجبة حتى منزلتين عشريتين.")
      endBusy()
      return
    }
    const alreadyWithdrawn = dataRef.current.withdrawals
      .filter(item => item.invoiceId === invoice.id && item.id !== withdrawalDraft.id)
      .reduce((sum, item) => sum + toHundredths(item.quantity), 0)
    const available = toHundredths(invoice.quantity) - alreadyWithdrawn
    if (toHundredths(quantity) > available) {
      setError(`الكمية المتاحة لهذه الفاتورة ${formatQuantity(Math.max(0, available) / 100)} لتر فقط. لم يتم حفظ المسحوب.`)
      endBusy()
      return
    }
    const savedWithdrawal: FuelWithdrawal = {
      ...withdrawalDraft,
      id: withdrawalDraft.id ?? createId(),
      quantity: toHundredths(quantity) / 100,
      withdrawnBy: withdrawalDraft.withdrawnBy.trim(),
      purpose: withdrawalDraft.purpose.trim(),
      beneficiary: withdrawalDraft.beneficiary.trim(),
    }
    const next: FuelData = {
      ...dataRef.current,
      withdrawals: withdrawalDraft.id
        ? dataRef.current.withdrawals.map(item => item.id === withdrawalDraft.id ? savedWithdrawal : item)
        : [savedWithdrawal, ...dataRef.current.withdrawals],
    }
    try {
      if (await commitData(next, withdrawalDraft.id ? "تم تحديث المسحوب." : "تم تسجيل المسحوب وربطه بالفاتورة.")) setWithdrawalDraft(null)
    } finally {
      endBusy()
    }
  }

  const removeInvoice = async (invoice: FuelInvoice) => {
    if (dataRef.current.withdrawals.some(item => item.invoiceId === invoice.id)) {
      setError("لا يمكن حذف فاتورة عليها مسحوبات. احذف المسحوبات المرتبطة بها أولاً.")
      return
    }
    if (!window.confirm(`هل تريد حذف فاتورة ${invoice.invoiceNumber}؟`)) return
    if (!startBusy()) return
    try {
      const next = { ...dataRef.current, invoices: dataRef.current.invoices.filter(item => item.id !== invoice.id) }
      if (await commitData(next, "تم حذف الفاتورة.")) {
        if (invoice.archiveFileId) {
          try {
            await deleteStoredFile(invoice.archiveFileId)
          } catch (cause) {
            setError(`حُذفت الفاتورة، لكن تعذر حذف ملف الأرشيف: ${cause instanceof Error ? cause.message : "خطأ غير معروف."}`)
          }
        }
      }
    } finally {
      endBusy()
    }
  }

  const removeWithdrawal = async (withdrawal: FuelWithdrawal) => {
    if (!window.confirm("هل تريد حذف سجل المسحوب؟ ستعود كميته إلى رصيد الفاتورة.")) return
    if (!startBusy()) return
    try {
      const next = { ...dataRef.current, withdrawals: dataRef.current.withdrawals.filter(item => item.id !== withdrawal.id) }
      await commitData(next, "تم حذف المسحوب وإعادة كميته إلى الرصيد المتاح.")
    } finally {
      endBusy()
    }
  }

  const openArchive = async (invoice: FuelInvoice) => {
    if (invoice.archiveDataUrl) {
      setArchiveView({ name: invoice.archiveFileName ?? invoice.invoiceNumber, url: invoice.archiveDataUrl, revoke: false })
      return
    }
    if (!invoice.archiveFileId || !invoice.archiveFileName) {
      setError("لا توجد صورة أو وثيقة مؤرشفة لهذه الفاتورة.")
      return
    }
    try {
      const file = await fetchStoredFile(invoice.archiveFileId, invoice.archiveFileName, invoice.archiveMimeType ?? "application/octet-stream")
      const url = URL.createObjectURL(file)
      setArchiveView({ name: invoice.archiveFileName, url, revoke: true })
      setError("")
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر فتح الفاتورة المؤرشفة: ${cause.message}` : "تعذر فتح الفاتورة المؤرشفة.")
    }
  }

  const openWithdrawal = (invoiceId = "") => {
    setWithdrawalDraft(emptyWithdrawalDraft(invoiceId))
    setTab("withdrawals")
  }

  const exportReport = () => window.print()

  if (loadedKey !== storageKey) {
    return <main dir="rtl" className="min-h-full bg-[#f4f7fb] p-8 text-center text-sm text-slate-600">{error || "جارٍ تحميل بيانات المحروقات..."}</main>
  }

  const visibleInvoices = data.invoices.filter(invoice => {
    const query = search.trim().toLocaleLowerCase("ar")
    return !query || [invoice.invoiceNumber, invoice.supplier, invoice.product, invoice.location, invoice.receivedBy]
      .some(value => value.toLocaleLowerCase("ar").includes(query))
  })
  const visibleReportInvoices = data.invoices.filter(invoice => {
    const query = search.trim().toLocaleLowerCase("ar")
    return !query
      || [invoice.invoiceNumber, invoice.supplier, invoice.product, invoice.location, invoice.receivedBy]
        .some(value => value.toLocaleLowerCase("ar").includes(query))
      || filteredWithdrawals.some(item => item.invoiceId === invoice.id)
  })

  return (
    <main dir="rtl" className="min-h-full bg-[#f4f7fb] px-4 py-7 text-slate-900 sm:px-7 lg:px-10">
      <div className="mx-auto max-w-7xl space-y-6">
        <header className="flex flex-wrap items-center justify-between gap-4 rounded-3xl bg-[#10243c] px-6 py-7 text-white shadow-xl sm:px-9">
          <div>
            <div className="inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/5 px-3 py-1.5 text-xs text-cyan-100"><Fuel className="h-4 w-4" />إدارة الوارد والمسحوبات والأرشيف</div>
            <h1 className="mt-3 text-2xl font-extrabold sm:text-3xl">منصة المحروقات</h1>
            <p className="mt-2 text-sm text-slate-300">كل مسحوب مرتبط بفاتورة وارد، ولا يمكن تجاوز رصيدها المتبقي.</p>
          </div>
          <button type="button" onClick={openNewInvoice} className="inline-flex items-center gap-2 rounded-xl bg-cyan-500 px-4 py-3 text-sm font-bold text-white hover:bg-cyan-400"><Plus className="h-4 w-4" />فاتورة وارد جديدة</button>
        </header>

        {error && <div role="alert" className="flex items-start justify-between gap-3 rounded-xl border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-800"><span>{error}</span><button type="button" onClick={() => setError("")} aria-label="إغلاق" className="shrink-0"><X className="h-4 w-4" /></button></div>}
        {notice && <div role="status" className="rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-800">{notice}</div>}

        <section className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
          {totals.map(type => (
            <article key={type.value} className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
              <h2 className="text-sm font-bold text-slate-800">{type.label}</h2>
              <div className="mt-3 grid grid-cols-3 gap-2 text-center">
                <div className="rounded-xl bg-emerald-50 p-2"><div className="text-[10px] text-emerald-700">الوارد</div><div className="mt-1 text-sm font-extrabold text-emerald-900">{formatQuantity(type.received)}</div></div>
                <div className="rounded-xl bg-amber-50 p-2"><div className="text-[10px] text-amber-700">المسحوب</div><div className="mt-1 text-sm font-extrabold text-amber-900">{formatQuantity(type.withdrawn)}</div></div>
                <div className="rounded-xl bg-sky-50 p-2"><div className="text-[10px] text-sky-700">المتبقي</div><div className="mt-1 text-sm font-extrabold text-sky-900">{formatQuantity(type.remaining)}</div></div>
              </div>
              <p className="mt-2 text-center text-[10px] text-slate-500">لتر</p>
            </article>
          ))}
        </section>

        <nav className="flex flex-wrap gap-2 rounded-2xl border border-slate-200 bg-white p-2 shadow-sm">
          {([
            ["dashboard", "الملخص"],
            ["invoices", "فواتير الوارد"],
            ["withdrawals", "المسحوبات"],
            ["report", "التقرير التفصيلي"],
          ] as Array<[Tab, string]>).map(([id, label]) => (
            <button key={id} type="button" onClick={() => setTab(id)} className={`rounded-xl px-4 py-2.5 text-sm font-semibold transition ${tab === id ? "bg-[#10243c] text-white" : "text-slate-600 hover:bg-slate-100"}`}>{label}</button>
          ))}
        </nav>

        {tab === "dashboard" && (
          <section className="grid gap-4 lg:grid-cols-2">
            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="flex items-center justify-between gap-3"><h2 className="font-bold">آخر فواتير الوارد</h2><button type="button" onClick={() => setTab("invoices")} className="text-xs font-semibold text-cyan-800 hover:underline">عرض الكل</button></div>
              {data.invoices.length === 0 ? <p className="py-8 text-center text-sm text-slate-500">لم تسجل فواتير وارد بعد.</p> : (
                <ul className="mt-3 divide-y">{data.invoices.slice(0, 5).map(invoice => <li key={invoice.id} className="flex items-center justify-between gap-3 py-3"><div className="min-w-0"><div className="truncate text-sm font-semibold">{invoice.invoiceNumber} · {invoice.supplier}</div><div className="mt-1 text-xs text-slate-500">{invoice.date} · {fuelLabel(invoice.fuelType)} · {invoice.product}</div></div><div className="shrink-0 text-left text-xs"><div className="font-bold">{formatQuantity(invoice.quantity)} لتر</div><div className="mt-1 text-slate-500">متبقي {formatQuantity(remainingFor(invoice, data.withdrawals))}</div></div></li>)}</ul>
              )}
            </div>
            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="flex items-center justify-between gap-3"><h2 className="font-bold">آخر المسحوبات</h2><button type="button" onClick={() => setTab("withdrawals")} className="text-xs font-semibold text-cyan-800 hover:underline">عرض الكل</button></div>
              {data.withdrawals.length === 0 ? <p className="py-8 text-center text-sm text-slate-500">لم تسجل مسحوبات بعد.</p> : (
                <ul className="mt-3 divide-y">{data.withdrawals.slice(0, 5).map(item => { const invoice = invoicesById.get(item.invoiceId); return <li key={item.id} className="flex items-center justify-between gap-3 py-3"><div className="min-w-0"><div className="truncate text-sm font-semibold">{item.purpose} · {item.beneficiary}</div><div className="mt-1 text-xs text-slate-500">{item.date} · {invoice?.invoiceNumber ?? "فاتورة محذوفة"} · سحبه {item.withdrawnBy}</div></div><div className="shrink-0 text-left text-sm font-bold">{formatQuantity(item.quantity)} لتر</div></li> })}</ul>
              )}
            </div>
          </section>
        )}

        {tab !== "dashboard" && (
          <section className="overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-sm">
            <div className="flex flex-wrap items-end justify-between gap-3 border-b border-slate-100 p-4 sm:p-5">
              <div>
                <h2 className="font-bold">{tab === "invoices" ? "فواتير الوارد" : tab === "withdrawals" ? "سجل المسحوبات" : "تقرير الوارد والمسحوبات"}</h2>
                <p className="mt-1 text-xs text-slate-500">{tab === "report" ? "حدد الفترة لمراجعة تفاصيل كل مسحوب والفاتورة التي صُرف منها." : "ابحث برقم الفاتورة أو الشركة أو الصنف أو المستفيد."}</p>
              </div>
              <div className="flex flex-wrap items-end gap-2">
                {tab === "report" && <>
                  <label className="text-[10px] font-semibold text-slate-600">من<input type="date" value={dateFrom} onChange={event => setDateFrom(event.target.value)} className="mt-1 block rounded-lg border px-2 py-2 text-xs" /></label>
                  <label className="text-[10px] font-semibold text-slate-600">إلى<input type="date" value={dateTo} onChange={event => setDateTo(event.target.value)} className="mt-1 block rounded-lg border px-2 py-2 text-xs" /></label>
                  <button type="button" onClick={exportReport} className="inline-flex items-center gap-1.5 rounded-lg border px-3 py-2 text-xs font-semibold hover:bg-slate-50"><Printer className="h-4 w-4" />طباعة التقرير</button>
                </>}
                <label className="relative block min-w-48"><Search className="absolute right-3 top-2.5 h-4 w-4 text-slate-400" /><input value={search} onChange={event => setSearch(event.target.value)} placeholder="بحث..." className="w-full rounded-lg border py-2 pr-9 pl-3 text-xs outline-none focus:border-cyan-600" /></label>
                {tab === "invoices" && <button type="button" onClick={openNewInvoice} className="inline-flex items-center gap-1.5 rounded-lg bg-cyan-800 px-3 py-2 text-xs font-semibold text-white hover:bg-cyan-900"><Plus className="h-4 w-4" />فاتورة وارد</button>}
                {tab === "withdrawals" && <button type="button" onClick={() => openWithdrawal()} disabled={!data.invoices.some(invoice => remainingFor(invoice, data.withdrawals) > 0)} className="inline-flex items-center gap-1.5 rounded-lg bg-cyan-800 px-3 py-2 text-xs font-semibold text-white hover:bg-cyan-900 disabled:cursor-not-allowed disabled:opacity-50"><Plus className="h-4 w-4" />تسجيل مسحوب</button>}
              </div>
            </div>

            {tab === "invoices" ? (
              data.invoices.length === 0 ? <div className="p-12 text-center text-sm text-slate-500">لا توجد فواتير وارد. أضف أول فاتورة للبدء.</div> : <div className="overflow-x-auto"><table className="w-full min-w-[1000px] text-right text-xs"><thead className="bg-slate-50 text-slate-600"><tr>{["رقم الفاتورة", "الشركة والتاريخ", "الصنف", "الوارد", "سعر الوحدة", "دفعها / استلمها", "الموقع الحالي", "المتبقي", "الإجراءات"].map(label => <th key={label} className="px-3 py-3 font-semibold">{label}</th>)}</tr></thead><tbody className="divide-y divide-slate-100">{visibleInvoices.map(invoice => <tr key={invoice.id} className="align-top hover:bg-slate-50/70"><td className="px-3 py-3 font-bold">{invoice.invoiceNumber}</td><td className="px-3 py-3"><div className="font-semibold">{invoice.supplier}</div><div className="mt-1 text-slate-500">{invoice.date}</div></td><td className="px-3 py-3">{fuelLabel(invoice.fuelType)} · {invoice.product}</td><td className="px-3 py-3 font-semibold">{formatQuantity(invoice.quantity)} لتر</td><td className="px-3 py-3">{formatMoney(invoice.unitPrice)}</td><td className="px-3 py-3"><div>{invoice.paidBy}</div><div className="mt-1 text-slate-500">{invoice.receivedBy}</div></td><td className="px-3 py-3">{invoice.location}</td><td className="px-3 py-3 font-bold text-cyan-900">{formatQuantity(remainingFor(invoice, data.withdrawals))} لتر</td><td className="px-3 py-3"><div className="flex flex-wrap gap-1.5">{remainingFor(invoice, data.withdrawals) > 0 && <button type="button" onClick={() => openWithdrawal(invoice.id)} className="rounded-md bg-cyan-50 px-2 py-1.5 font-semibold text-cyan-900 hover:bg-cyan-100">صرف</button>}{(invoice.archiveFileId || invoice.archiveDataUrl) && <button type="button" onClick={() => void openArchive(invoice)} title="عرض الفاتورة المؤرشفة" className="rounded-md bg-violet-50 p-1.5 text-violet-800 hover:bg-violet-100"><Archive className="h-4 w-4" /></button>}<button type="button" onClick={() => { setInvoiceDraft({ ...invoice }); setInvoiceFile(null) }} title="تعديل" className="rounded-md border p-1.5 hover:bg-slate-100"><Pencil className="h-4 w-4" /></button><button type="button" onClick={() => void removeInvoice(invoice)} title="حذف" disabled={busy} className="rounded-md p-1.5 text-rose-700 hover:bg-rose-50 disabled:opacity-50"><Trash2 className="h-4 w-4" /></button></div>{invoice.archiveFileName && <div className="mt-1 text-[10px] text-slate-500">{invoice.archiveFileName}</div>}</td></tr>)}</tbody></table></div>
            ) : (
              <>
                {tab === "report" && <div className="report-print-logo"><img src={TASWIYAH_LOGO_PATH} alt="شعار تسوية هب" /></div>}
                {tab === "report" && <div className="grid gap-2 border-b bg-slate-50 p-4 text-xs sm:grid-cols-3"><div>المسحوب ضمن الفترة: <strong>{formatQuantity(filteredWithdrawals.reduce((sum, item) => sum + item.quantity, 0))} لتر</strong></div><div>عدد الحركات: <strong>{filteredWithdrawals.length}</strong></div><div>إجمالي الوارد: <strong>{formatQuantity(data.invoices.reduce((sum, invoice) => sum + invoice.quantity, 0))} لتر</strong></div></div>}
                {filteredWithdrawals.length === 0 ? <div className="p-12 text-center text-sm text-slate-500">لا توجد مسحوبات تطابق البحث أو الفترة المحددة.</div> : <div className="overflow-x-auto"><table className="w-full min-w-[900px] text-right text-xs"><thead className="bg-slate-50 text-slate-600"><tr>{["تاريخ السحب", "الفاتورة الواردة", "نوع / صنف المحروق", "الكمية", "من سحبه", "الاستخدام", "لصالح", "الإجراءات"].map(label => <th key={label} className="px-3 py-3 font-semibold">{label}</th>)}</tr></thead><tbody className="divide-y divide-slate-100">{filteredWithdrawals.map(item => { const invoice = invoicesById.get(item.invoiceId); return <tr key={item.id} className="hover:bg-slate-50/70"><td className="px-3 py-3">{item.date}</td><td className="px-3 py-3"><div className="font-semibold">{invoice?.invoiceNumber ?? "فاتورة غير متاحة"}</div><div className="mt-1 text-slate-500">{invoice?.supplier ?? "—"}</div></td><td className="px-3 py-3">{invoice ? `${fuelLabel(invoice.fuelType)} · ${invoice.product}` : "—"}</td><td className="px-3 py-3 font-bold">{formatQuantity(item.quantity)} لتر</td><td className="px-3 py-3">{item.withdrawnBy}</td><td className="px-3 py-3">{item.purpose}</td><td className="px-3 py-3">{item.beneficiary}</td><td className="px-3 py-3"><div className="flex gap-1.5"><button type="button" onClick={() => setWithdrawalDraft({ ...item })} title="تعديل" className="rounded-md border p-1.5 hover:bg-slate-100"><Pencil className="h-4 w-4" /></button><button type="button" onClick={() => void removeWithdrawal(item)} title="حذف" disabled={busy} className="rounded-md p-1.5 text-rose-700 hover:bg-rose-50 disabled:opacity-50"><Trash2 className="h-4 w-4" /></button></div></td></tr> })}</tbody></table></div>}
                {tab === "report" && <div className="overflow-x-auto border-t"><div className="px-4 pt-4"><h3 className="font-bold">الوارد حسب الفاتورة والتفصيل ضمن الفترة</h3><p className="mt-1 text-xs text-slate-500">يظهر الوارد المسجل لكل فاتورة، وتفاصيل وجهة استخدام المسحوبات المطابقة للفترة المحددة.</p></div><table className="mt-3 w-full min-w-[850px] text-right text-xs"><thead className="bg-slate-50 text-slate-600"><tr>{["الفاتورة / الشركة", "الصنف", "الوارد", "المسحوب ضمن الفترة وتفاصيله", "المتبقي الحالي"].map(label => <th key={label} className="px-3 py-3 font-semibold">{label}</th>)}</tr></thead><tbody className="divide-y divide-slate-100">{visibleReportInvoices.map(invoice => { const invoiceItems = filteredWithdrawals.filter(item => item.invoiceId === invoice.id); const periodQuantity = invoiceItems.reduce((sum, item) => sum + item.quantity, 0); return <tr key={invoice.id} className="align-top hover:bg-slate-50/70"><td className="px-3 py-3"><div className="font-bold">{invoice.invoiceNumber}</div><div className="mt-1">{invoice.supplier}</div><div className="mt-1 text-slate-500">{invoice.date}</div></td><td className="px-3 py-3">{fuelLabel(invoice.fuelType)} · {invoice.product}</td><td className="px-3 py-3 font-semibold">{formatQuantity(invoice.quantity)} لتر</td><td className="px-3 py-3"><div className="font-bold">{formatQuantity(periodQuantity)} لتر</div>{invoiceItems.length > 0 && <ul className="mt-1 space-y-1 text-slate-600">{invoiceItems.map(item => <li key={item.id}>{item.date}: {formatQuantity(item.quantity)} لتر إلى {item.purpose} لصالح {item.beneficiary} (سحبه {item.withdrawnBy})</li>)}</ul>}</td><td className="px-3 py-3 font-semibold">{formatQuantity(remainingFor(invoice, data.withdrawals))} لتر</td></tr> })}</tbody></table>{visibleReportInvoices.length === 0 && <p className="p-8 text-center text-sm text-slate-500">لا توجد فواتير وارد تطابق البحث.</p>}</div>}
              </>
            )}
          </section>
        )}
      </div>

      {invoiceDraft && <div role="dialog" aria-modal="true" aria-labelledby="fuel-invoice-title" className="fixed inset-0 z-[80] flex items-start justify-center overflow-y-auto bg-slate-950/50 p-4 sm:items-center"><form onSubmit={event => void submitInvoice(event)} className="my-4 w-full max-w-3xl space-y-4 rounded-2xl bg-white p-5 shadow-2xl sm:p-7">
        <div className="flex items-center justify-between gap-3"><div><h2 id="fuel-invoice-title" className="text-lg font-bold">{invoiceDraft.id ? "تعديل فاتورة وارد" : "تسجيل فاتورة وارد"}</h2><p className="mt-1 text-xs text-slate-500">رقم الفاتورة يُنشأ تلقائياً، ويمكن تعديله إذا لزم.</p></div><button type="button" onClick={() => setInvoiceDraft(null)} aria-label="إغلاق" className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          <label className={labelClass()}><span>رقم الفاتورة</span><input required value={invoiceDraft.invoiceNumber} onChange={event => setInvoiceDraft({ ...invoiceDraft, invoiceNumber: event.target.value })} className={inputClass()} /></label>
          <label className={labelClass()}><span>اسم الشركة المورّدة</span><input required value={invoiceDraft.supplier} onChange={event => setInvoiceDraft({ ...invoiceDraft, supplier: event.target.value })} className={inputClass()} /></label>
          <label className={labelClass()}><span>تاريخ الفاتورة</span><input required type="date" value={invoiceDraft.date} onChange={event => setInvoiceDraft({ ...invoiceDraft, date: event.target.value })} className={inputClass()} /></label>
          <label className={labelClass()}><span>نوع المحروق</span><select value={invoiceDraft.fuelType} onChange={event => setInvoiceDraft({ ...invoiceDraft, fuelType: event.target.value as FuelType })} className={inputClass()}>{FUEL_TYPES.map(type => <option key={type.value} value={type.value}>{type.label}</option>)}</select></label>
          <label className={labelClass()}><span>اسم الصنف / المحروق</span><input required value={invoiceDraft.product} onChange={event => setInvoiceDraft({ ...invoiceDraft, product: event.target.value })} placeholder="مثال: سولار للمولد" className={inputClass()} /></label>
          <label className={labelClass()}><span>الكمية الواردة (لتر)</span><input required type="number" min="0.01" step="0.01" value={invoiceDraft.quantity || ""} onChange={event => setInvoiceDraft({ ...invoiceDraft, quantity: Number(event.target.value) })} className={inputClass()} /></label>
          <label className={labelClass()}><span>سعر اللتر / الوحدة</span><input required type="number" min="0" step="0.01" value={invoiceDraft.unitPrice} onChange={event => setInvoiceDraft({ ...invoiceDraft, unitPrice: Number(event.target.value) })} className={inputClass()} /></label>
          <label className={labelClass()}><span>من قام بالدفع</span><input required value={invoiceDraft.paidBy} onChange={event => setInvoiceDraft({ ...invoiceDraft, paidBy: event.target.value })} className={inputClass()} /></label>
          <label className={labelClass()}><span>من استلم الكمية</span><input required value={invoiceDraft.receivedBy} onChange={event => setInvoiceDraft({ ...invoiceDraft, receivedBy: event.target.value })} className={inputClass()} /></label>
          <label className={labelClass()}><span>مكان وجود الكمية حالياً</span><input required value={invoiceDraft.location} onChange={event => setInvoiceDraft({ ...invoiceDraft, location: event.target.value })} className={inputClass()} /></label>
          <label className={`${labelClass()} sm:col-span-2`}><span>صورة الفاتورة للأرشيف (اختياري)</span><input type="file" accept="image/*,.pdf" onChange={event => setInvoiceFile(event.target.files?.[0] ?? null)} className={`${inputClass()} file:ml-3 file:rounded-md file:border-0 file:bg-cyan-50 file:px-3 file:py-1 file:text-xs file:font-semibold file:text-cyan-900`} /><span className="font-normal text-slate-500">{invoiceFile ? invoiceFile.name : invoiceDraft.id && data.invoices.find(item => item.id === invoiceDraft.id)?.archiveFileName ? `الملف الحالي: ${data.invoices.find(item => item.id === invoiceDraft.id)?.archiveFileName} (اختر ملفاً لاستبداله)` : "يمكنك رفع صورة أو PDF لعرضها لاحقاً من الأرشيف."}</span></label>
        </div>
        <div className="flex justify-end gap-2 border-t pt-4"><button type="button" disabled={busy} onClick={() => setInvoiceDraft(null)} className="rounded-xl border px-4 py-2.5 text-sm font-semibold hover:bg-slate-50">إلغاء</button><button type="submit" disabled={busy} className="rounded-xl bg-cyan-800 px-5 py-2.5 text-sm font-semibold text-white hover:bg-cyan-900 disabled:opacity-50">{busy ? "جارٍ الحفظ..." : "حفظ الفاتورة"}</button></div>
      </form></div>}

      {withdrawalDraft && <div role="dialog" aria-modal="true" aria-labelledby="fuel-withdrawal-title" className="fixed inset-0 z-[80] flex items-start justify-center overflow-y-auto bg-slate-950/50 p-4 sm:items-center"><form onSubmit={event => void submitWithdrawal(event)} className="my-4 w-full max-w-2xl space-y-4 rounded-2xl bg-white p-5 shadow-2xl sm:p-7">
        <div className="flex items-center justify-between gap-3"><div><h2 id="fuel-withdrawal-title" className="text-lg font-bold">{withdrawalDraft.id ? "تعديل سجل المسحوب" : "تسجيل مسحوب"}</h2><p className="mt-1 text-xs text-slate-500">اختر فاتورة الوارد التي صُرف منها المحروق؛ الكمية لن تتجاوز المتبقي.</p></div><button type="button" onClick={() => setWithdrawalDraft(null)} aria-label="إغلاق" className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <div className="grid gap-3 sm:grid-cols-2">
          <label className={`${labelClass()} sm:col-span-2`}><span>فاتورة الوارد المرتبطة</span><select required value={withdrawalDraft.invoiceId} onChange={event => setWithdrawalDraft({ ...withdrawalDraft, invoiceId: event.target.value })} className={inputClass()}><option value="">اختر الفاتورة</option>{data.invoices.filter(invoice => invoice.id === withdrawalDraft.invoiceId || remainingFor(invoice, data.withdrawals) > 0).map(invoice => <option key={invoice.id} value={invoice.id}>{invoice.invoiceNumber} · {invoice.supplier} · {invoice.product} · المتبقي {formatQuantity(remainingFor(invoice, data.withdrawals))} لتر</option>)}</select></label>
          {withdrawalDraft.invoiceId && <div className="sm:col-span-2 rounded-xl bg-cyan-50 px-3 py-2 text-xs text-cyan-950">الرصيد المتاح: <strong>{formatQuantity(Math.max(0, (data.invoices.find(invoice => invoice.id === withdrawalDraft.invoiceId)?.quantity ?? 0) - data.withdrawals.filter(item => item.invoiceId === withdrawalDraft.invoiceId && item.id !== withdrawalDraft.id).reduce((sum, item) => sum + item.quantity, 0)))} لتر</strong></div>}
          <label className={labelClass()}><span>تاريخ السحب</span><input required type="date" value={withdrawalDraft.date} onChange={event => setWithdrawalDraft({ ...withdrawalDraft, date: event.target.value })} className={inputClass()} /></label>
          <label className={labelClass()}><span>الكمية المسحوبة (لتر)</span><input required type="number" min="0.01" step="0.01" value={withdrawalDraft.quantity || ""} onChange={event => setWithdrawalDraft({ ...withdrawalDraft, quantity: Number(event.target.value) })} className={inputClass()} /></label>
          <label className={labelClass()}><span>من قام بالسحب</span><input required value={withdrawalDraft.withdrawnBy} onChange={event => setWithdrawalDraft({ ...withdrawalDraft, withdrawnBy: event.target.value })} className={inputClass()} /></label>
          <label className={labelClass()}><span>لصالح من</span><input required value={withdrawalDraft.beneficiary} onChange={event => setWithdrawalDraft({ ...withdrawalDraft, beneficiary: event.target.value })} className={inputClass()} /></label>
          <label className={`${labelClass()} sm:col-span-2`}><span>وجهة الاستخدام</span><input required value={withdrawalDraft.purpose} onChange={event => setWithdrawalDraft({ ...withdrawalDraft, purpose: event.target.value })} placeholder="مثال: موتور السوبر" className={inputClass()} /></label>
        </div>
        <div className="flex justify-end gap-2 border-t pt-4"><button type="button" disabled={busy} onClick={() => setWithdrawalDraft(null)} className="rounded-xl border px-4 py-2.5 text-sm font-semibold hover:bg-slate-50">إلغاء</button><button type="submit" disabled={busy} className="rounded-xl bg-cyan-800 px-5 py-2.5 text-sm font-semibold text-white hover:bg-cyan-900 disabled:opacity-50">{busy ? "جارٍ الحفظ..." : "حفظ المسحوب"}</button></div>
      </form></div>}

      {archiveView && <div role="dialog" aria-modal="true" aria-label="الفاتورة المؤرشفة" className="fixed inset-0 z-[90] flex items-center justify-center bg-slate-950/70 p-3 sm:p-8"><div className="flex max-h-full w-full max-w-5xl flex-col overflow-hidden rounded-2xl bg-white shadow-2xl"><div className="flex items-center justify-between gap-3 border-b px-4 py-3"><div className="flex min-w-0 items-center gap-2"><FileText className="h-4 w-4 shrink-0 text-violet-700" /><span className="truncate text-sm font-semibold">{archiveView.name}</span></div><button type="button" onClick={() => setArchiveView(null)} aria-label="إغلاق الفاتورة المؤرشفة" className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>{archiveView.url.startsWith("data:application/pdf") || archiveView.name.toLocaleLowerCase("en").endsWith(".pdf") ? <iframe title={archiveView.name} src={archiveView.url} className="min-h-[65vh] flex-1" /> : <div className="overflow-auto bg-slate-100 p-3 text-center"><img src={archiveView.url} alt={archiveView.name} className="mx-auto max-h-[75vh] max-w-full object-contain" /></div>}</div></div>}
    </main>
  )
}
