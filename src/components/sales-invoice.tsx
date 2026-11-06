"use client"

import { useEffect, useMemo, useState } from "react"
import { FolderOpen, Plus, Printer, RotateCcw, Save, Trash2 } from "lucide-react"
import { centralStorageEnabled, centralStorageGet, centralStorageSet } from "../lib/central-storage"

type InvoiceItem = {
  id: string
  code: string
  name: string
  unit: string
  quantity: string
  unitPrice: string
}

type InvoiceCurrency = "ILS" | "USD" | "EUR"
type InvoiceDraft = {
  invoiceNumber: string
  date: string
  licenseNumber: string
  customerName: string
  customerAddress: string
  items: InvoiceItem[]
  currency: InvoiceCurrency
  discountAmount: string
  discountRate: string
  taxRate: string
  cashPayment: string
  checkPayment: string
  receiptNumber: string
  notes: string
  customAmountInWords: string
}
type SavedInvoice = InvoiceDraft & { id: string; updatedAt: string }

const CURRENCIES: Record<InvoiceCurrency, { label: string; symbol: string; major: string; minor: string }> = {
  ILS: { label: "شيكل", symbol: "₪", major: "شيكلًا إسرائيليًا", minor: "أغورة" },
  USD: { label: "دولار", symbol: "$", major: "دولارًا أمريكيًا", minor: "سنتًا" },
  EUR: { label: "يورو", symbol: "€", major: "يورو", minor: "سنتًا" },
}

const INPUT_CLASS = "w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm outline-none focus:border-cyan-600 focus:ring-2 focus:ring-cyan-100"

function localDateValue() {
  const now = new Date()
  return new Date(now.getTime() - now.getTimezoneOffset() * 60_000).toISOString().slice(0, 10)
}

function formatAmount(value: number) {
  return new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value)
}

function belowThousandToArabic(value: number) {
  const ones = ["", "واحد", "اثنان", "ثلاثة", "أربعة", "خمسة", "ستة", "سبعة", "ثمانية", "تسعة"]
  const teens = ["عشرة", "أحد عشر", "اثنا عشر", "ثلاثة عشر", "أربعة عشر", "خمسة عشر", "ستة عشر", "سبعة عشر", "ثمانية عشر", "تسعة عشر"]
  const tens = ["", "", "عشرون", "ثلاثون", "أربعون", "خمسون", "ستون", "سبعون", "ثمانون", "تسعون"]
  const hundreds = ["", "مائة", "مائتان", "ثلاثمائة", "أربعمائة", "خمسمائة", "ستمائة", "سبعمائة", "ثمانمائة", "تسعمائة"]
  const parts: string[] = []
  const hundred = Math.floor(value / 100)
  const remainder = value % 100
  if (hundred) parts.push(hundreds[hundred])
  if (remainder >= 10 && remainder < 20) {
    parts.push(teens[remainder - 10])
  } else if (remainder >= 20) {
    const one = remainder % 10
    parts.push(one ? `${ones[one]} و${tens[Math.floor(remainder / 10)]}` : tens[Math.floor(remainder / 10)])
  } else if (remainder) {
    parts.push(ones[remainder])
  }
  return parts.join(" و")
}

function integerToArabicWords(value: number) {
  if (value === 0) return "صفر"
  const scales = [
    { one: "", two: "", few: "", many: "" },
    { one: "ألف", two: "ألفان", few: "آلاف", many: "ألفًا" },
    { one: "مليون", two: "مليونان", few: "ملايين", many: "مليونًا" },
    { one: "مليار", two: "ملياران", few: "مليارات", many: "مليارًا" },
  ]
  const groups: string[] = []
  let remainder = Math.floor(value)
  let scale = 0
  while (remainder > 0) {
    const group = remainder % 1000
    if (group) {
      if (scale === 0) {
        groups.unshift(belowThousandToArabic(group))
      } else {
        const names = scales[scale]
        const groupWords = belowThousandToArabic(group)
        const scaleWords = group === 1 ? names.one
          : group === 2 ? names.two
            : group >= 3 && group <= 10 ? `${groupWords} ${names.few}`
              : `${groupWords} ${names.many}`
        groups.unshift(scaleWords)
      }
    }
    remainder = Math.floor(remainder / 1000)
    scale += 1
  }
  return groups.join(" و")
}

function amountToArabicWords(value: number, currency: InvoiceCurrency) {
  const currencyInfo = CURRENCIES[currency]
  const minorUnits = Math.round(Math.max(0, value) * 100)
  const whole = Math.floor(minorUnits / 100)
  const fraction = minorUnits % 100
  const wholeWords = `${integerToArabicWords(whole)} ${currencyInfo.major}`
  return fraction ? `${wholeWords} و${integerToArabicWords(fraction)} ${currencyInfo.minor} فقط لا غير` : `${wholeWords} فقط لا غير`
}

function formatDate(value: string) {
  const [year, month, day] = value.split("-")
  return year && month && day ? `${day}/${month}/${year}` : value
}

function createItem(): InvoiceItem {
  return { id: globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random()}`, code: "", name: "", unit: "", quantity: "1", unitPrice: "0" }
}

export function SalesInvoice({ username }: { username: string }) {
  const normalizedUsername = username.normalize("NFKC").trim().toLocaleLowerCase("en-US")
  const numberStorageKey = `sales_invoice_next_number:${normalizedUsername}`
  const invoicesStorageKey = `sales_invoices:${normalizedUsername}`
  const [invoiceNumber, setInvoiceNumber] = useState("500")
  const [date, setDate] = useState(localDateValue)
  const [licenseNumber, setLicenseNumber] = useState("")
  const [customerName, setCustomerName] = useState("")
  const [customerAddress, setCustomerAddress] = useState("")
  const [items, setItems] = useState<InvoiceItem[]>([createItem()])
  const [currency, setCurrency] = useState<InvoiceCurrency>("ILS")
  const [discountAmount, setDiscountAmount] = useState("0")
  const [discountRate, setDiscountRate] = useState("0")
  const [taxRate, setTaxRate] = useState("0")
  const [cashPayment, setCashPayment] = useState("")
  const [checkPayment, setCheckPayment] = useState("")
  const [receiptNumber, setReceiptNumber] = useState("")
  const [notes, setNotes] = useState("")
  const [customAmountInWords, setCustomAmountInWords] = useState("")
  const [error, setError] = useState("")
  const [savedInvoices, setSavedInvoices] = useState<SavedInvoice[]>([])
  const [savedInvoicesReady, setSavedInvoicesReady] = useState(false)
  const [saving, setSaving] = useState(false)
  const [currentSavedId, setCurrentSavedId] = useState<string | null>(null)

  useEffect(() => {
    let active = true
    setSavedInvoicesReady(false)
    try {
      const savedNextNumber = localStorage.getItem(numberStorageKey)
      if (savedNextNumber && /^\d+$/.test(savedNextNumber)) setInvoiceNumber(savedNextNumber)
    } catch (cause) {
      console.error("تعذر تحميل رقم الفاتورة التالي:", cause)
      setError("تعذر تحميل رقم الفاتورة المحفوظ من هذا المتصفح.")
    }
    const loadSavedInvoices = async () => {
      try {
        const stored = centralStorageEnabled
          ? await centralStorageGet<SavedInvoice[]>(invoicesStorageKey) ?? []
          : JSON.parse(localStorage.getItem(invoicesStorageKey) ?? "[]") as SavedInvoice[]
        if (!Array.isArray(stored)) throw new Error("قائمة الفواتير المحفوظة غير صالحة.")
        if (active) {
          setSavedInvoices(stored)
          setSavedInvoicesReady(true)
        }
      } catch (cause) {
        if (active) {
          console.error("تعذر تحميل الفواتير المحفوظة:", cause)
          setError(cause instanceof Error ? `تعذر تحميل الفواتير المحفوظة: ${cause.message}` : "تعذر تحميل الفواتير المحفوظة.")
        }
      }
    }
    void loadSavedInvoices()
    return () => { active = false }
  }, [numberStorageKey])

  const subtotal = useMemo(() => items.reduce((sum, item) =>
    sum + (Number(item.quantity) || 0) * (Number(item.unitPrice) || 0), 0,
  ), [items])
  const discount = Math.min(subtotal, Math.max(0, Number(discountAmount) || 0) + subtotal * Math.max(0, Number(discountRate) || 0) / 100)
  const afterDiscount = Math.max(0, subtotal - discount)
  const tax = afterDiscount * Math.max(0, Number(taxRate) || 0) / 100
  const total = afterDiscount + tax
  const currencyInfo = CURRENCIES[currency]
  const generatedAmountInWords = amountToArabicWords(total, currency)
  const amountInWords = customAmountInWords.trim() || generatedAmountInWords
  const currentDraft: InvoiceDraft = {
    invoiceNumber, date, licenseNumber, customerName, customerAddress, items, currency,
    discountAmount, discountRate, taxRate, cashPayment, checkPayment, receiptNumber, notes, customAmountInWords,
  }

  const updateItem = (id: string, field: keyof Omit<InvoiceItem, "id">, value: string) => {
    setItems(current => current.map(item => item.id === id ? { ...item, [field]: value } : item))
  }

  const startNewInvoice = () => {
    const nextNumber = String((Number(invoiceNumber) || 499) + 1)
    try {
      localStorage.setItem(numberStorageKey, nextNumber)
    } catch (cause) {
      console.error("تعذر حفظ رقم الفاتورة التالي:", cause)
      setError("تعذر حفظ رقم الفاتورة التالي؛ تأكد من تفعيل مساحة التخزين في المتصفح.")
      return
    }
    setInvoiceNumber(nextNumber)
    setDate(localDateValue())
    setLicenseNumber("")
    setCustomerName("")
    setCustomerAddress("")
    setItems([createItem()])
    setCurrency("ILS")
    setDiscountAmount("0")
    setDiscountRate("0")
    setTaxRate("0")
    setCashPayment("")
    setCheckPayment("")
    setReceiptNumber("")
    setNotes("")
    setCustomAmountInWords("")
    setCurrentSavedId(null)
    setError("")
  }

  const saveInvoice = async () => {
    if (!savedInvoicesReady) {
      setError("انتظر تحميل الفواتير المحفوظة قبل حفظ الفاتورة.")
      return
    }
    if (!items.some(item => item.name.trim())) {
      setError("أدخل صنفاً واحداً على الأقل قبل حفظ الفاتورة.")
      return
    }
    const now = new Date().toISOString()
    const existing = currentSavedId ? savedInvoices.find(invoice => invoice.id === currentSavedId) : undefined
    const saved: SavedInvoice = {
      ...currentDraft,
      id: existing?.id ?? globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random()}`,
      updatedAt: now,
    }
    const next = [saved, ...savedInvoices.filter(invoice => invoice.id !== saved.id)]
    setSaving(true)
    setError("")
    try {
      if (centralStorageEnabled) await centralStorageSet(invoicesStorageKey, next)
      else localStorage.setItem(invoicesStorageKey, JSON.stringify(next))
      setSavedInvoices(next)
      setCurrentSavedId(saved.id)
    } catch (cause) {
      console.error("تعذر حفظ الفاتورة:", cause)
      setError(cause instanceof Error ? `تعذر حفظ الفاتورة: ${cause.message}` : "تعذر حفظ الفاتورة.")
    } finally {
      setSaving(false)
    }
  }

  const openSavedInvoice = (invoice: SavedInvoice) => {
    setInvoiceNumber(invoice.invoiceNumber)
    setDate(invoice.date)
    setLicenseNumber(invoice.licenseNumber)
    setCustomerName(invoice.customerName)
    setCustomerAddress(invoice.customerAddress)
    setItems(invoice.items)
    setCurrency(invoice.currency)
    setDiscountAmount(invoice.discountAmount)
    setDiscountRate(invoice.discountRate)
    setTaxRate(invoice.taxRate)
    setCashPayment(invoice.cashPayment)
    setCheckPayment(invoice.checkPayment)
    setReceiptNumber(invoice.receiptNumber)
    setNotes(invoice.notes)
    setCustomAmountInWords(invoice.customAmountInWords)
    setCurrentSavedId(invoice.id)
    setError("")
  }

  const deleteSavedInvoice = async (invoice: SavedInvoice) => {
    const next = savedInvoices.filter(item => item.id !== invoice.id)
    try {
      if (centralStorageEnabled) await centralStorageSet(invoicesStorageKey, next)
      else localStorage.setItem(invoicesStorageKey, JSON.stringify(next))
      setSavedInvoices(next)
      if (currentSavedId === invoice.id) setCurrentSavedId(null)
    } catch (cause) {
      console.error("تعذر حذف الفاتورة المحفوظة:", cause)
      setError(cause instanceof Error ? `تعذر حذف الفاتورة المحفوظة: ${cause.message}` : "تعذر حذف الفاتورة المحفوظة.")
    }
  }

  const printInvoice = () => {
    if (!items.some(item => item.name.trim())) {
      setError("أدخل صنفاً واحداً على الأقل قبل طباعة الفاتورة.")
      return
    }
    setError("")
    window.print()
  }

  return (
    <main className="sales-invoice-page min-h-full bg-slate-50 p-4 sm:p-6 lg:p-8">
      <style>{`
        @page { size: A4 portrait; margin: 0; }
        .invoice-print { color: #111827; font-weight: 600; }
        .invoice-print .border-slate-300, .invoice-print .border-slate-400, .invoice-print .border-slate-500 { border-color: #475569 !important; }
        .invoice-print thead th { background: #e2e8f0 !important; color: #0f172a !important; font-size: 11px; font-weight: 800 !important; }
        .invoice-print tbody td { color: #111827 !important; font-size: 11px; font-weight: 600 !important; }
        .invoice-print .text-slate-500, .invoice-print .text-slate-600 { color: #1f2937 !important; }
        .invoice-print .bg-slate-100 { background: #e2e8f0 !important; }
        .invoice-print .font-bold { font-weight: 800 !important; }
        @media print {
          html, body, #root { width: 210mm !important; min-width: 210mm !important; margin: 0 !important; padding: 0 !important; background: #fff !important; }
          body * { visibility: hidden !important; }
          .invoice-print, .invoice-print * { visibility: visible !important; }
          .app-shell { display: block !important; width: 210mm !important; height: auto !important; max-height: none !important; overflow: visible !important; }
          .recon-scope { display: block !important; height: auto !important; overflow: visible !important; }
          .invoice-builder { display: none !important; }
          .sales-invoice-page { display: block !important; width: 210mm !important; min-height: 297mm !important; margin: 0 !important; padding: 0 !important; overflow: visible !important; background: #fff !important; }
          .invoice-layout { display: block !important; }
          .invoice-print { position: static !important; width: 190mm !important; min-height: 277mm !important; margin: 0 auto !important; padding: 10mm !important; border: 0 !important; border-radius: 0 !important; box-shadow: none !important; background: #fff !important; }
          .invoice-paper { width: 190mm !important; min-height: 277mm !important; }
          .invoice-controls { display: none !important; }
          * { print-color-adjust: exact !important; -webkit-print-color-adjust: exact !important; }
        }
      `}</style>
      <div className="mx-auto max-w-[1500px]">
        <header className="invoice-builder mb-5 flex flex-wrap items-center justify-between gap-3">
          <div>
            <h1 className="text-2xl font-bold text-slate-900">منصة الفواتير</h1>
            <p className="mt-1 text-sm text-slate-600">أدخل بيانات الفاتورة وشاهد معاينة مطابقة للنموذج قبل الطباعة أو الحفظ PDF.</p>
          </div>
          <div className="flex gap-2">
            <button type="button" onClick={startNewInvoice} className="inline-flex items-center gap-2 rounded-xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-semibold text-slate-700 hover:bg-slate-100"><RotateCcw className="h-4 w-4" />فاتورة جديدة</button>
            <button type="button" onClick={() => void saveInvoice()} disabled={saving || !savedInvoicesReady} className="inline-flex items-center gap-2 rounded-xl bg-emerald-700 px-4 py-2.5 text-sm font-semibold text-white hover:bg-emerald-800 disabled:cursor-not-allowed disabled:opacity-50"><Save className="h-4 w-4" />{saving ? "جاري الحفظ..." : currentSavedId ? "تحديث الفاتورة" : "حفظ الفاتورة"}</button>
            <button type="button" onClick={printInvoice} className="inline-flex items-center gap-2 rounded-xl bg-cyan-700 px-4 py-2.5 text-sm font-semibold text-white hover:bg-cyan-800"><Printer className="h-4 w-4" />طباعة / PDF</button>
          </div>
        </header>
        {error && <div role="alert" className="invoice-builder mb-4 rounded-xl border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-800">{error}</div>}

        <div className="invoice-layout grid items-start gap-5 xl:grid-cols-[minmax(350px,0.8fr)_minmax(650px,1.2fr)]">
          <section className="invoice-builder space-y-4 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
            <section className="rounded-xl border border-slate-200 p-3">
              <h2 className="mb-2 flex items-center gap-2 text-sm font-bold text-slate-800"><FolderOpen className="h-4 w-4 text-cyan-700" />الفواتير المحفوظة ({savedInvoices.length})</h2>
              {savedInvoices.length ? <div className="max-h-56 space-y-1 overflow-y-auto">
                {savedInvoices.map(invoice => (
                  <div key={invoice.id} className={`flex items-center gap-2 rounded-lg px-2 py-1.5 ${currentSavedId === invoice.id ? "bg-cyan-50" : "hover:bg-slate-50"}`}>
                    <button type="button" onClick={() => openSavedInvoice(invoice)} className="min-w-0 flex-1 text-right">
                      <span className="block truncate text-xs font-semibold text-slate-800">فاتورة {invoice.invoiceNumber} {invoice.customerName && `· ${invoice.customerName}`}</span>
                      <span className="block text-[10px] text-slate-500">{invoice.date} · {CURRENCIES[invoice.currency].label}</span>
                    </button>
                    <button type="button" onClick={() => void deleteSavedInvoice(invoice)} aria-label={`حذف الفاتورة ${invoice.invoiceNumber}`} title="حذف الفاتورة المحفوظة" className="rounded-md p-1.5 text-rose-700 hover:bg-rose-50"><Trash2 className="h-3.5 w-3.5" /></button>
                  </div>
                ))}
              </div> : <p className="text-xs text-slate-500">{savedInvoicesReady ? "لا توجد فواتير محفوظة بعد." : "جاري تحميل الفواتير..."}</p>}
            </section>
            <section>
              <h2 className="mb-3 text-sm font-bold text-slate-800">بيانات الفاتورة والزبون</h2>
              <div className="grid gap-3 sm:grid-cols-2">
                <label className="text-xs font-semibold text-slate-600">رقم الفاتورة<input value={invoiceNumber} onChange={event => setInvoiceNumber(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} /></label>
                <label className="text-xs font-semibold text-slate-600">التاريخ<input type="date" value={date} onChange={event => setDate(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} /></label>
                <label className="text-xs font-semibold text-slate-600">العملة<select value={currency} onChange={event => setCurrency(event.target.value as InvoiceCurrency)} className={`${INPUT_CLASS} mt-1.5`}><option value="ILS">شيكل (₪)</option><option value="USD">دولار ($)</option><option value="EUR">يورو (€)</option></select></label>
                <label className="text-xs font-semibold text-slate-600">رقم المشغّل المرخّص<input value={licenseNumber} onChange={event => setLicenseNumber(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} /></label>
                <label className="text-xs font-semibold text-slate-600">اسم الزبون<input value={customerName} onChange={event => setCustomerName(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} /></label>
                <label className="text-xs font-semibold text-slate-600 sm:col-span-2">العنوان<input value={customerAddress} onChange={event => setCustomerAddress(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} /></label>
              </div>
            </section>

            <section>
              <div className="mb-3 flex items-center justify-between gap-2">
                <h2 className="text-sm font-bold text-slate-800">الأصناف</h2>
                <button type="button" onClick={() => setItems(current => [...current, createItem()])} className="inline-flex items-center gap-1 rounded-lg bg-cyan-50 px-3 py-2 text-xs font-semibold text-cyan-800 hover:bg-cyan-100"><Plus className="h-3.5 w-3.5" />إضافة صنف</button>
              </div>
              <div className="space-y-3">
                {items.map((item, index) => (
                  <fieldset key={item.id} className="grid grid-cols-2 gap-2 rounded-xl border border-slate-200 p-3 sm:grid-cols-3">
                    <legend className="px-1 text-xs font-semibold text-slate-500">الصنف {index + 1}</legend>
                    <label className="text-[11px] text-slate-500">رقم الصنف<input value={item.code} onChange={event => updateItem(item.id, "code", event.target.value)} className={`${INPUT_CLASS} mt-1`} /></label>
                    <label className="col-span-2 text-[11px] text-slate-500 sm:col-span-2">اسم الصنف<input value={item.name} onChange={event => updateItem(item.id, "name", event.target.value)} className={`${INPUT_CLASS} mt-1`} /></label>
                    <label className="text-[11px] text-slate-500">الوحدة<input value={item.unit} onChange={event => updateItem(item.id, "unit", event.target.value)} className={`${INPUT_CLASS} mt-1`} /></label>
                    <label className="text-[11px] text-slate-500">الكمية<input type="number" min="0" step="any" value={item.quantity} onChange={event => updateItem(item.id, "quantity", event.target.value)} className={`${INPUT_CLASS} mt-1`} dir="ltr" /></label>
                    <label className="text-[11px] text-slate-500">سعر الوحدة<input type="number" min="0" step="0.01" value={item.unitPrice} onChange={event => updateItem(item.id, "unitPrice", event.target.value)} className={`${INPUT_CLASS} mt-1`} dir="ltr" /></label>
                    {items.length > 1 && <button type="button" onClick={() => setItems(current => current.filter(row => row.id !== item.id))} aria-label={`حذف الصنف ${index + 1}`} className="justify-self-start self-end rounded-lg p-2 text-rose-700 hover:bg-rose-50"><Trash2 className="h-4 w-4" /></button>}
                  </fieldset>
                ))}
              </div>
            </section>

            <section>
              <h2 className="mb-3 text-sm font-bold text-slate-800">الخصم والضريبة والدفع</h2>
              <div className="grid gap-3 sm:grid-cols-2">
                <label className="text-xs font-semibold text-slate-600">مبلغ الخصم ({currencyInfo.symbol})<input type="number" min="0" step="0.01" value={discountAmount} onChange={event => setDiscountAmount(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} dir="ltr" /></label>
                <label className="text-xs font-semibold text-slate-600">نسبة الخصم %<input type="number" min="0" step="0.01" value={discountRate} onChange={event => setDiscountRate(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} dir="ltr" /></label>
                <label className="text-xs font-semibold text-slate-600">نسبة الضريبة %<input type="number" min="0" step="0.01" value={taxRate} onChange={event => setTaxRate(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} dir="ltr" /></label>
                <label className="text-xs font-semibold text-slate-600">نقداً ({currencyInfo.symbol})<input type="number" min="0" step="0.01" value={cashPayment} onChange={event => setCashPayment(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} dir="ltr" /></label>
                <label className="text-xs font-semibold text-slate-600">شيكات ({currencyInfo.symbol})<input type="number" min="0" step="0.01" value={checkPayment} onChange={event => setCheckPayment(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} dir="ltr" /></label>
                <label className="text-xs font-semibold text-slate-600">رقم الوصل<input value={receiptNumber} onChange={event => setReceiptNumber(event.target.value)} className={`${INPUT_CLASS} mt-1.5`} /></label>
                <label className="text-xs font-semibold text-slate-600 sm:col-span-2">المبلغ كتابةً
                  <input value={customAmountInWords} onChange={event => setCustomAmountInWords(event.target.value)} placeholder={generatedAmountInWords} className={`${INPUT_CLASS} mt-1.5`} />
                  <span className="mt-1 block font-normal text-slate-500">يُحسب تلقائياً، ويمكنك كتابة صياغتك الخاصة لتظهر على الفاتورة.</span>
                </label>
                <label className="text-xs font-semibold text-slate-600 sm:col-span-2">ملاحظات الفاتورة<textarea value={notes} onChange={event => setNotes(event.target.value)} rows={3} className={`${INPUT_CLASS} mt-1.5`} /></label>
              </div>
            </section>
          </section>

          <article className="invoice-print invoice-paper mx-auto min-h-[277mm] w-full max-w-[190mm] border border-slate-300 bg-white p-6 text-[12px] text-slate-900 shadow-lg sm:p-8" dir="rtl">
            <header dir="ltr" className="grid grid-cols-[1fr_1.7fr] items-start gap-4 border-b-2 border-slate-500 pb-3">
              <div className="text-center">
                <h2 className="text-sm font-bold">فاتورة ضريبية - مبيعات</h2>
                <p className="font-semibold">الأصل</p>
                <div className="mt-2 rounded-lg border border-slate-400 px-2 py-1 text-[10px]">رقم المشغّل المرخّص: {licenseNumber}</div>
              </div>
              <div className="text-center">
                <h1 className="text-xl font-bold">سوبر ماركت بيت الجملة</h1>
                <p className="mt-1 text-xs">غزة - دير البلح - شارع أبو عريف</p>
                <div className="mt-1 flex justify-center gap-8 text-[10px]"><span>هاتف:</span><span>فاكس:</span></div>
              </div>
              <div className="col-span-2 flex items-center justify-center gap-5 text-sm font-bold">
                <span className="h-px flex-1 bg-slate-500" />
                <span>رقم {invoiceNumber}</span>
                <span className="h-px flex-1 bg-slate-500" />
              </div>
            </header>

            <section dir="ltr" className="mt-3 grid grid-cols-[0.8fr_1.2fr] gap-2">
              <div dir="rtl" className="flex min-h-[18mm] flex-col justify-center rounded-lg border border-slate-400 px-3 py-2 text-[11px]">
                <p>تاريخ الفاتورة: {formatDate(date)}</p>
                <p className="mt-2">سعر الوحدة لا يشمل الضريبة المضافة</p>
              </div>
              <div dir="rtl" className="min-h-[18mm] rounded-lg border border-slate-400 px-3 py-2 text-[11px]">
                <p>الاسم: {customerName}</p>
                <p className="mt-2">العنوان: {customerAddress}</p>
              </div>
            </section>

            <table className="mt-3 w-full table-fixed border-collapse text-[11px]">
              <thead><tr className="bg-slate-100">
                <th className="w-[12%] border border-slate-500 p-1.5">رقم الصنف</th>
                <th className="w-[36%] border border-slate-500 p-1.5">اسم الصنف</th>
                <th className="w-[9%] border border-slate-500 p-1.5">الوحدة</th>
                <th className="w-[12%] border border-slate-500 p-1.5">الكمية</th>
                <th className="w-[14%] border border-slate-500 p-1.5">سعر الوحدة</th>
                <th className="w-[17%] border border-slate-500 p-1.5">الإجمالي</th>
              </tr></thead>
              <tbody>
                {items.map(item => <tr key={item.id} className="h-7">
                  <td className="border border-slate-400 px-1.5 py-1">{item.code}</td>
                  <td className="border border-slate-400 px-1.5 py-1">{item.name}</td>
                  <td className="border border-slate-400 px-1.5 py-1 text-center">{item.unit}</td>
                  <td className="border border-slate-400 px-1.5 py-1 text-center">{item.quantity}</td>
                  <td className="border border-slate-400 px-1.5 py-1 text-center">{formatAmount(Number(item.unitPrice) || 0)} {currencyInfo.symbol}</td>
                  <td className="border border-slate-400 px-1.5 py-1 text-center">{formatAmount((Number(item.quantity) || 0) * (Number(item.unitPrice) || 0))} {currencyInfo.symbol}</td>
                </tr>)}
              </tbody>
            </table>

            <section dir="ltr" className="mt-1 grid grid-cols-[1fr_1.2fr] gap-3">
              <div dir="rtl" className="grid grid-cols-[1fr_auto] self-start text-[11px]">
                <span className="border border-slate-400 px-2 py-1.5">المجموع</span><span className="min-w-[24mm] border border-slate-400 px-2 py-1.5 text-left" dir="ltr">{formatAmount(subtotal)} {currencyInfo.symbol}</span>
                <span className="border border-slate-400 px-2 py-1.5">مبلغ الخصم</span><span className="border border-slate-400 px-2 py-1.5 text-left" dir="ltr">{formatAmount(discount)} {currencyInfo.symbol}</span>
                <span className="border border-slate-400 px-2 py-1.5">نسبة الخصم %</span><span className="border border-slate-400 px-2 py-1.5 text-left" dir="ltr">{Number(discountRate) || 0}</span>
                <span className="border border-slate-400 px-2 py-1.5">المبلغ بعد الخصم</span><span className="border border-slate-400 px-2 py-1.5 text-left" dir="ltr">{formatAmount(afterDiscount)} {currencyInfo.symbol}</span>
                <span className="border border-slate-400 px-2 py-1.5">نسبة الضريبة %</span><span className="border border-slate-400 px-2 py-1.5 text-left" dir="ltr">{Number(taxRate) || 0}</span>
                <span className="border border-slate-500 bg-slate-100 px-2 py-1.5 font-bold">مبلغ الفاتورة</span><span className="border border-slate-500 bg-slate-100 px-2 py-1.5 text-left font-bold" dir="ltr">{formatAmount(total)} {currencyInfo.symbol}</span>
              </div>
              <div dir="rtl" className="min-h-[27mm] rounded-lg border border-slate-400 p-3 text-[11px]">
                <h3 className="mb-2 text-center font-bold">تفاصيل الدفع</h3>
                <div className="grid grid-cols-2 gap-y-2">
                  <span>نقداً: {cashPayment ? `${formatAmount(Number(cashPayment) || 0)} ${currencyInfo.symbol}` : ""}</span><span>رقم الوصل: {receiptNumber}</span>
                  <span>شيكات: {checkPayment ? `${formatAmount(Number(checkPayment) || 0)} ${currencyInfo.symbol}` : ""}</span><span>رقم الحساب:</span>
                  <span className="col-span-2 mt-1 border-t border-slate-200 pt-2 font-bold">على الحساب: {formatAmount(Math.max(0, total - (Number(cashPayment) || 0) - (Number(checkPayment) || 0)))} {currencyInfo.symbol}</span>
                </div>
              </div>
            </section>
            <p className="mt-1 min-h-5 text-[11px] font-bold">المبلغ كتابة: {amountInWords}</p>

            <section dir="ltr" className="mt-3 grid grid-cols-2 gap-5">
              <div dir="rtl" className="flex min-h-[31mm] items-center justify-around text-[11px]">
                <div className="text-center">
                  <p className="mb-3 font-semibold">توقيع المشغّل المرخّص:</p>
                </div>
                <p>اسم وتوقيع المستلم:</p>
              </div>
              <div dir="rtl" className="min-h-[31mm] overflow-hidden rounded-lg border border-slate-400 text-[11px]">
                <h3 className="bg-slate-700 px-3 py-1 text-right font-bold text-white">ملاحظات</h3>
                <p className="whitespace-pre-wrap p-3">{notes}</p>
              </div>
            </section>
          </article>
        </div>
      </div>
    </main>
  )
}
