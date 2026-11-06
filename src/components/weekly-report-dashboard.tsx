"use client"

import { useEffect, useMemo, useState } from "react"
import { AlertTriangle, CalendarDays, Download, Printer, RefreshCw } from "lucide-react"
import { centralStorageGet, centralStorageSet } from "../lib/central-storage"
import {
  INGESTION_METADATA_KEY,
  INGESTION_PARSED_DATA_KEY_PREFIX,
  type IngestionFileMetadata,
} from "../lib/data-ingestion"
import { upgradeParsedData, type ParsedFileData } from "../lib/data-parser"
import {
  calculateWeeklyReport,
  type WeeklyReportFile,
  type WeeklyReportSettings,
} from "../lib/weekly-report"
import { TASWIYAH_LOGO_PATH } from "../lib/report-branding"
import { useGlobalLoading } from "./global-loading"

const emptySettings: WeeklyReportSettings = {
  weeklyPayroll: null,
  monthlyRent: null,
  shrinkageCost: null,
}

function isoWeekValue(date: Date) {
  const thursday = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()))
  thursday.setUTCDate(thursday.getUTCDate() + 4 - (thursday.getUTCDay() || 7))
  const firstThursday = new Date(Date.UTC(thursday.getUTCFullYear(), 0, 4))
  const week = 1 + Math.round((thursday.getTime() - firstThursday.getTime()) / 604_800_000)
  return `${thursday.getUTCFullYear()}-W${String(week).padStart(2, "0")}`
}

function isoWeekRange(week: string) {
  const match = week.match(/^(\d{4})-W(\d{2})$/)
  if (!match) return null
  const year = Number(match[1])
  const weekNumber = Number(match[2])
  if (weekNumber < 1 || weekNumber > 53) return null
  const jan4 = new Date(Date.UTC(year, 0, 4))
  const monday = new Date(jan4)
  monday.setUTCDate(jan4.getUTCDate() - ((jan4.getUTCDay() + 6) % 7) + (weekNumber - 1) * 7)
  const nextMonday = new Date(monday)
  nextMonday.setUTCDate(monday.getUTCDate() + 7)
  return { start: monday, end: nextMonday }
}

function settingsKey(weekStart: Date) {
  return `weekly_report_settings:${weekStart.toISOString().slice(0, 10)}`
}

function amount(value: number | null, currency = "ILS") {
  if (value === null) return "غير مكتمل"
  return `${new Intl.NumberFormat("ar", { maximumFractionDigits: 2 }).format(value)} ${currency}`
}

function formatDate(date: Date) {
  return new Intl.DateTimeFormat("ar", { dateStyle: "medium", timeZone: "UTC" }).format(date)
}

function currencyForFiles(files: WeeklyReportFile[]) {
  const currencies = Array.from(new Set(files.map(file => file.parsedData.summary.currency).filter(Boolean)))
  return currencies.length === 1 ? currencies[0]! : "ILS"
}

function parseOptionalNumber(value: string) {
  if (!value.trim()) return null
  const parsed = Number(value)
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : Number.NaN
}

function csvCell(value: string | number | null) {
  let safe = value === null ? "" : String(value)
  if (/^[\s]*[=+\-@]/.test(safe)) safe = `'${safe}`
  return `"${safe.replace(/"/g, '""')}"`
}

function exportReport(
  week: string,
  report: ReturnType<typeof calculateWeeklyReport>,
  currency: string,
) {
  const rows: Array<Array<string | number | null>> = [
    ["تقرير أداء تسوية هب الأسبوعي", week],
    ["إجمالي المبيعات", report.revenue],
    ["تكلفة البضاعة المباعة", report.cogs],
    ["مجمل الربح", report.grossProfit],
    ["الرواتب الأسبوعية", report.payroll],
    ["حصة الإيجار الأسبوعية", report.rent],
    ["السولار والمحروقات", report.fuel],
    ["الكهرباء والمرافق", report.utilities],
    ["المصاريف المفرقة", report.otherExpenses],
    ["التوالف والعجز", report.shrinkage],
    ["إجمالي المصاريف التشغيلية", report.totalOperatingExpenses],
    ["صافي الربح", report.netProfit],
    ["هامش الربح %", report.profitMargin],
    [],
    ["التطبيق", "مبيعات POS", "الكشف الفعلي", "الفرق (الفعلي - المتوقع)"],
    ...["Jawwal Pay", "Bank App", "App 3"].map((name, index) => [
      name,
      report.expectedApps[index],
      report.actualApps[index],
      report.appDifferences[index],
    ]),
    [],
    ["التدفق والديون"],
    ["ديون زبائن جديدة", report.newDebts],
    ["تحصيل ديون سابقة", report.collectedDebts],
    ["مدفوعات الموردين", report.vendorPayments],
    ["العملة", currency],
  ]
  const content = `\uFEFF${rows.map(row => row.map(value => csvCell(value ?? null)).join(",")).join("\r\n")}`
  const url = URL.createObjectURL(new Blob([content], { type: "text/csv;charset=utf-8" }))
  const anchor = document.createElement("a")
  anchor.href = url
  anchor.download = `taswiyah-weekly-report-${week}.csv`
  anchor.click()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}

export function WeeklyReportDashboard() {
  const [week, setWeek] = useState(() => isoWeekValue(new Date()))
  const [allFiles, setAllFiles] = useState<IngestionFileMetadata[]>([])
  const [reportFiles, setReportFiles] = useState<WeeklyReportFile[]>([])
  const [settings, setSettings] = useState<WeeklyReportSettings>(emptySettings)
  const [payrollInput, setPayrollInput] = useState("")
  const [rentInput, setRentInput] = useState("")
  const [shrinkageInput, setShrinkageInput] = useState("")
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState("")
  const [message, setMessage] = useState("")

  useGlobalLoading(loading && !error, "جاري تحميل بيانات التقرير الأسبوعي...")
  useGlobalLoading(saving, "جاري حفظ إعدادات التقرير...")

  const range = useMemo(() => isoWeekRange(week), [week])
  const key = useMemo(() => range ? settingsKey(range.start) : "", [range])
  const report = useMemo(() => calculateWeeklyReport(reportFiles, settings), [reportFiles, settings])
  const currency = useMemo(() => currencyForFiles(reportFiles), [reportFiles])

  useEffect(() => {
    if (!range || !key) return
    let active = true
    setLoading(true)
    setError("")
    setMessage("")
    void Promise.all([
      centralStorageGet<unknown>(INGESTION_METADATA_KEY),
      centralStorageGet<Partial<WeeklyReportSettings>>(key),
    ]).then(async ([savedMetadata, savedSettings]) => {
      const metadata = Array.isArray(savedMetadata) ? savedMetadata as IngestionFileMetadata[] : []
      const settingsValue: WeeklyReportSettings = {
        weeklyPayroll: typeof savedSettings?.weeklyPayroll === "number" && savedSettings.weeklyPayroll >= 0 ? savedSettings.weeklyPayroll : null,
        monthlyRent: typeof savedSettings?.monthlyRent === "number" && savedSettings.monthlyRent >= 0 ? savedSettings.monthlyRent : null,
        shrinkageCost: typeof savedSettings?.shrinkageCost === "number" && savedSettings.shrinkageCost >= 0 ? savedSettings.shrinkageCost : null,
      }
      const parsed = await Promise.all(metadata.map(async file => {
        const inline = upgradeParsedData(file.parsedData)
        const parsedData = inline ?? upgradeParsedData(await centralStorageGet<unknown>(`${INGESTION_PARSED_DATA_KEY_PREFIX}${file.fileId}`))
        return parsedData ? { ...file, parsedData } as WeeklyReportFile : null
      }))
      if (!active) return
      setAllFiles(metadata)
      setReportFiles(parsed.filter((file): file is WeeklyReportFile => file !== null)
        .filter(file => {
          const uploadedAt = Date.parse(file.uploadTimestamp)
          return Number.isFinite(uploadedAt)
            && uploadedAt >= range.start.getTime()
            && uploadedAt < range.end.getTime()
        }))
      setSettings(settingsValue)
      setPayrollInput(settingsValue.weeklyPayroll === null ? "" : String(settingsValue.weeklyPayroll))
      setRentInput(settingsValue.monthlyRent === null ? "" : String(settingsValue.monthlyRent))
      setShrinkageInput(settingsValue.shrinkageCost === null ? "" : String(settingsValue.shrinkageCost))
    }).catch(cause => {
      if (active) setError(cause instanceof Error ? cause.message : "تعذر تحميل بيانات التقرير الأسبوعي.")
    }).finally(() => {
      if (active) setLoading(false)
    })
    return () => { active = false }
  }, [key, range])

  const saveSettings = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (!key) {
      setError("تعذر تحديد أسبوع صالح لحفظ الإعدادات.")
      return
    }
    const next: WeeklyReportSettings = {
      weeklyPayroll: parseOptionalNumber(payrollInput),
      monthlyRent: parseOptionalNumber(rentInput),
      shrinkageCost: parseOptionalNumber(shrinkageInput),
    }
    if ([next.weeklyPayroll, next.monthlyRent, next.shrinkageCost].some(value => Number.isNaN(value))) {
      setError("أدخل مبالغ رقمية موجبة أو صفراً، أو اترك الحقل فارغاً إذا كانت البيانات غير متاحة.")
      return
    }
    setSaving(true)
    setError("")
    setMessage("")
    try {
      await centralStorageSet(key, next)
      setSettings(next)
      setMessage("تم حفظ مدخلات الأسبوع في التخزين المركزي.")
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذر حفظ مدخلات التقرير.")
    } finally {
      setSaving(false)
    }
  }

  const selectedRange = range
    ? `${formatDate(range.start)} – ${formatDate(new Date(range.end.getTime() - 86_400_000))}`
    : ""
  const totalKnownFiles = allFiles.filter(file => {
    const uploadedAt = Date.parse(file.uploadTimestamp)
    return !!range && Number.isFinite(uploadedAt)
      && uploadedAt >= range.start.getTime()
      && uploadedAt < range.end.getTime()
  }).length
  const salesCostMissing = report.revenue === null || report.cogs === null
  const apps = [
    { name: "Jawwal Pay", color: "bg-violet-500" },
    { name: "Bank App", color: "bg-emerald-500" },
    { name: "App 3", color: "bg-blue-500" },
  ]
  const expenses = [
    { label: "الرواتب الأسبوعية", method: "إدخال أسبوعي مباشر", value: report.payroll },
    { label: "حصة الإيجار الأسبوعية", method: "الإيجار الشهري × 12 ÷ 52", value: report.rent },
    { label: "السولار والمحروقات", method: "لترات الملفات × تكلفة اللتر (أو إجمالي ملف الوقود)", value: report.fuel },
    { label: "الكهرباء والمرافق", method: "ملفات المصاريف الموسومة كمرافق", value: report.utilities },
    { label: "المصاريف التشغيلية المفرقة", method: "مجموع ملفات المصاريف غير الوقود والمرافق", value: report.otherExpenses },
    { label: "التوالف والعجز", method: "إدخال أسبوعي بسعر التكلفة", value: report.shrinkage },
  ]

  return (
    <section className="mx-auto max-w-7xl space-y-6 p-5 sm:p-8" dir="rtl">
      <div className="report-print-logo"><img src={TASWIYAH_LOGO_PATH} alt="شعار تسوية هب" /></div>
      <header className="flex flex-wrap items-start justify-between gap-4">
        <div>
          <p className="text-sm font-semibold text-primary">Taswiyah Hub · Weekly Financial &amp; Operations Report</p>
          <h1 className="mt-1 text-2xl font-bold sm:text-3xl">تقرير الأداء الأسبوعي الشامل</h1>
          <p className="mt-2 text-sm text-muted-foreground">ملخص المبيعات والتكاليف والتشغيل ومطابقة المحافظ الإلكترونية.</p>
        </div>
        <div className="weekly-report-no-print flex flex-wrap items-center gap-2">
          <label className="inline-flex items-center gap-2 rounded-xl border bg-card px-3 py-2">
            <CalendarDays className="h-4 w-4 text-primary" />
            <span className="text-xs font-semibold">الأسبوع</span>
            <input
              type="week"
              value={week}
              onChange={event => setWeek(event.target.value)}
              className="max-w-36 bg-transparent text-sm outline-none"
              aria-label="اختر أسبوع التقرير"
            />
          </label>
          <button type="button" onClick={() => window.print()} className="inline-flex items-center gap-2 rounded-xl border bg-card px-3 py-2 text-sm font-semibold hover:bg-muted">
            <Printer className="h-4 w-4" /> طباعة / PDF
          </button>
          <button type="button" onClick={() => exportReport(week, report, currency)} disabled={loading} className="inline-flex items-center gap-2 rounded-xl bg-primary px-3 py-2 text-sm font-semibold text-primary-foreground disabled:opacity-50">
            <Download className="h-4 w-4" /> تصدير CSV
          </button>
        </div>
      </header>

      <div className="rounded-xl border border-blue-200 bg-blue-50 px-4 py-3 text-sm text-blue-900">
        الفترة: <strong>{selectedRange}</strong>. تُدرج الملفات بحسب <strong>تاريخ رفعها</strong>، ويُفترض أن كل ملف يخص هذا الأسبوع؛ الصفوف داخل الملف لا تُقسّم حسب تاريخ الحركة.
      </div>
      {message && <p role="status" className="rounded-xl bg-emerald-50 px-4 py-3 text-sm text-emerald-800">{message}</p>}
      {error && <p role="alert" className="rounded-xl bg-destructive/10 px-4 py-3 text-sm text-destructive">{error}</p>}

      <section className="space-y-3" aria-labelledby="weekly-kpi-heading">
        <div>
          <h2 id="weekly-kpi-heading" className="text-lg font-bold">المؤشرات التنفيذية</h2>
          <p className="text-xs text-muted-foreground">{loading ? "جارٍ تحميل البيانات..." : `${reportFiles.length} ملفاً محللاً من أصل ${totalKnownFiles} ملف ضمن الأسبوع`}</p>
        </div>
        <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
          {[
            { label: "إجمالي المبيعات", value: amount(report.revenue, currency), hint: "من ملفات POS / المبيعات", color: "border-blue-200 bg-blue-50" },
            { label: "تكلفة البضاعة (COGS)", value: amount(report.cogs, currency), hint: "من ملفات ربحية الأقسام", color: "border-violet-200 bg-violet-50" },
            { label: "المصاريف والنزيف", value: amount(report.totalOperatingExpenses, currency), hint: "يشمل مدخلات المصاريف اليدوية والملفات", color: "border-amber-200 bg-amber-50" },
            { label: "صافي الربح التشغيلي", value: amount(report.netProfit, currency), hint: report.profitMargin === null ? "أكمل البيانات المطلوبة لحساب الربح والهامش" : `هامش الربح ${report.profitMargin.toFixed(2)}%`, color: "border-emerald-200 bg-emerald-50" },
          ].map(card => (
            <article key={card.label} className={`rounded-2xl border p-5 ${card.color}`}>
              <p className="text-sm font-semibold text-muted-foreground">{card.label}</p>
              <p className="mt-3 break-words text-2xl font-bold tabular-nums">{card.value}</p>
              <p className="mt-2 text-xs text-muted-foreground">{card.hint}</p>
            </article>
          ))}
        </div>
        {report.grossProfit !== null && <p className="rounded-xl border bg-card px-4 py-3 text-sm">مجمل الربح: <strong>{amount(report.grossProfit, currency)}</strong> (المبيعات − تكلفة البضاعة)</p>}
      </section>

      <section className="rounded-2xl border bg-card p-4 sm:p-5">
        <h2 className="mb-4 text-lg font-bold">مطابقة المحافظ والتطبيقات الإلكترونية</h2>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[650px] text-right text-sm">
            <thead><tr className="border-b text-xs text-muted-foreground"><th className="px-3 py-3">التطبيق / المحفظة</th><th className="px-3 py-3">مبيعات POS المتوقعة</th><th className="px-3 py-3">المقبوض في الكشف</th><th className="px-3 py-3">الفارق (الفعلي − المتوقع)</th><th className="px-3 py-3">الحالة</th></tr></thead>
            <tbody>
              {apps.map((app, index) => {
                const difference = report.appDifferences[index]
                const status = difference === null ? "بيانات غير مكتملة" : Math.abs(difference) < 0.01 ? "مطابق" : difference < 0 ? "عجز" : "زيادة"
                return (
                  <tr key={app.name} className="border-b last:border-0">
                    <td className="px-3 py-3 font-semibold"><span className={`me-2 inline-block h-2.5 w-2.5 rounded-full ${app.color}`} />{app.name}</td>
                    <td className="px-3 py-3 tabular-nums">{amount(report.expectedApps[index], currency)}</td>
                    <td className="px-3 py-3 tabular-nums">{amount(report.actualApps[index], currency)}</td>
                    <td className="px-3 py-3 tabular-nums">{amount(difference, currency)}</td>
                    <td className={`px-3 py-3 font-semibold ${difference === null ? "text-muted-foreground" : Math.abs(difference) < 0.01 ? "text-emerald-700" : difference < 0 ? "text-red-700" : "text-amber-700"}`}>
                      {difference !== null && Math.abs(difference) >= 0.01 ? <AlertTriangle className="me-1 inline h-4 w-4" /> : null}{status}
                    </td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
        {report.unmatchedWalletTotal > 0 && <p className="mt-3 text-xs text-amber-700">مبالغ كشوف غير مرتبطة تلقائياً بتطبيق: {amount(report.unmatchedWalletTotal, currency)}. راجع أسماء الملفات.</p>}
        {report.appDifferences.some(value => value === null) && <p className="mt-3 text-xs text-muted-foreground">أضف ملف POS وملف كشف واضح الاسم لكل تطبيق حتى تظهر نتيجة المطابقة.</p>}
      </section>

      <section className="rounded-2xl border bg-card p-4 sm:p-5">
        <h2 className="mb-4 text-lg font-bold">المصاريف التشغيلية والمحروقات</h2>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[600px] text-right text-sm">
            <thead><tr className="border-b text-xs text-muted-foreground"><th className="px-3 py-3">بند المصروف</th><th className="px-3 py-3">طريقة الحساب والأصل</th><th className="px-3 py-3">التكلفة</th></tr></thead>
            <tbody>{expenses.map(item => <tr key={item.label} className="border-b last:border-0"><td className="px-3 py-3 font-semibold">{item.label}</td><td className="px-3 py-3 text-muted-foreground">{item.method}</td><td className="px-3 py-3 font-semibold tabular-nums">{amount(item.value, currency)}</td></tr>)}</tbody>
            <tfoot><tr className="bg-muted/50 font-bold"><td colSpan={2} className="px-3 py-3">الإجمالي التشغيلي</td><td className="px-3 py-3 tabular-nums">{amount(report.totalOperatingExpenses, currency)}</td></tr></tfoot>
          </table>
        </div>
      </section>

      <section className="rounded-2xl border bg-card p-4 sm:p-5">
        <h2 className="mb-4 text-lg font-bold">حركة التدفق النقدي والديون</h2>
        <div className="grid gap-3 sm:grid-cols-2">
          {[
            ["ديون الزبائن الجديدة (الآجل)", report.newDebts],
            ["تحصيلات ديون سابقة", report.collectedDebts],
            ["مشتريات آجلة من الموردين", null],
            ["مدفوعات وسداد للموردين", report.vendorPayments],
          ].map(([label, value]) => (
            <div key={label as string} className="flex items-center justify-between gap-3 rounded-xl border bg-muted/20 p-4">
              <span className="text-sm font-medium">{label}</span>
              <strong className="shrink-0 tabular-nums">{amount(value as number | null, currency)}</strong>
            </div>
          ))}
        </div>
        <p className="mt-3 text-xs text-muted-foreground">مشتريات الموردين الآجلة لا تُستنتج من إجمالي الفواتير؛ يتطلب احتسابها ملفاً يميز صراحةً المشتريات الآجلة.</p>
      </section>

      <section className="weekly-report-no-print rounded-2xl border bg-card p-4 sm:p-5">
        <h2 className="text-lg font-bold">مدخلات التكاليف الأسبوعية</h2>
        <p className="mt-1 text-sm text-muted-foreground">أدخل القيم التي لا تأتي من ملفات الرفع. اترك الحقل فارغاً إن لم تتوفر البيانات؛ لا تُعامل كصفر.</p>
        <form onSubmit={event => void saveSettings(event)} className="mt-4 grid gap-4 sm:grid-cols-3">
          <label className="space-y-2 text-sm font-semibold">
            الرواتب الأسبوعية
            <input type="number" min="0" step="0.01" value={payrollInput} onChange={event => setPayrollInput(event.target.value)} placeholder="أدخل المبلغ" className="h-11 w-full rounded-xl border bg-background px-3 font-normal" />
          </label>
          <label className="space-y-2 text-sm font-semibold">
            الإيجار الشهري
            <input type="number" min="0" step="0.01" value={rentInput} onChange={event => setRentInput(event.target.value)} placeholder="أدخل المبلغ" className="h-11 w-full rounded-xl border bg-background px-3 font-normal" />
          </label>
          <label className="space-y-2 text-sm font-semibold">
            التوالف والعجز بسعر التكلفة
            <input type="number" min="0" step="0.01" value={shrinkageInput} onChange={event => setShrinkageInput(event.target.value)} placeholder="أدخل المبلغ" className="h-11 w-full rounded-xl border bg-background px-3 font-normal" />
          </label>
          <div className="flex flex-wrap items-center gap-3 sm:col-span-3">
            <button type="submit" disabled={saving || loading} className="inline-flex items-center gap-2 rounded-xl bg-primary px-4 py-2.5 text-sm font-semibold text-primary-foreground disabled:opacity-50">
              {saving ? <RefreshCw className="h-4 w-4 animate-spin" /> : null}{saving ? "جارٍ الحفظ..." : "حفظ مدخلات الأسبوع"}
            </button>
            <span className="text-xs text-muted-foreground">حصة الإيجار الأسبوعية = الإيجار الشهري × 12 ÷ 52</span>
          </div>
        </form>
      </section>

      {report.warnings.length > 0 && (
        <section className="rounded-2xl border border-amber-200 bg-amber-50 p-4" aria-label="ملاحظات اكتمال التقرير">
          <h2 className="font-bold text-amber-900">ملاحظات واكتمال البيانات</h2>
          <ul className="mt-2 list-inside list-disc space-y-1 text-sm text-amber-900">{report.warnings.map(warning => <li key={warning}>{warning}</li>)}</ul>
          {salesCostMissing && <p className="mt-3 text-xs font-semibold text-amber-900">لا يعرض التقرير صافي ربح نهائياً حتى تتوفر المبيعات وتكلفة البضاعة.</p>}
        </section>
      )}
    </section>
  )
}
