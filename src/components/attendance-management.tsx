"use client"

import { useEffect, useMemo, useState } from "react"
import QRCode from "qrcode"
import { ArrowRightLeft, CalendarDays, Check, Copy, Download, Link, Printer, RefreshCw, Shield, ShieldCheck, Users } from "lucide-react"
import { centralStorageEnabled, centralStorageGet, centralStorageSet } from "../lib/central-storage"

type Employee = { id: string; employeeNumber: string; name: string; dailyRate: number; customerNumber: string; phone: string }
type PayrollLine = { employeeId: string; employeeNumber: string; name: string; dailyRate: number; attendanceDays?: number; gross: number; advance: number; debtBalance: number | null; debtDeduction: number; percentageDeduction: number; percentageBase: number }
type PayrollSheet = { id: string; name: string; periodFrom: string; periodTo: string; days: number; percentageEnabled: boolean; locked?: boolean; lines: PayrollLine[] }
type PayrollData = { employees: Employee[]; sheets: PayrollSheet[] }
type StaffPin = { employeeId: string; employeeNumber: string; pinConfigured: boolean }
type Punch = { id: string; employeeId: string; employeeNumber: string; employeeName: string; action: "in" | "out"; timestamp: string; date: string; time: string }
type Summary = { employeeId: string; employeeNumber: string; name: string; hours: number; workedDays: number; calculatedDays: number }
type DailyAttendance = { employeeId: string; employeeNumber: string; name: string; date: string; arrival: string; departure: string; hours: number }
type DayPolicy = "hours" | "full_day"

const EMPTY_PAYROLL: PayrollData = { employees: [], sheets: [] }
const INPUT_CLASS = "w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm outline-none focus:border-cyan-700 focus:ring-2 focus:ring-cyan-100"

function localDateValue() {
  const now = new Date()
  return new Date(now.getTime() - now.getTimezoneOffset() * 60_000).toISOString().slice(0, 10)
}

function firstDayOfMonth(value: string) {
  return `${value.slice(0, 7)}-01`
}

function roundDays(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100
}

function formatHours(value: number) {
  return new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value)
}

function normalizeUsername(value: string) {
  return value.normalize("NFKC").trim().toLocaleLowerCase("en-US")
}

async function responseJson<T>(response: Response): Promise<T> {
  const body = await response.json().catch(() => ({})) as T & { error?: string }
  if (!response.ok) throw new Error(body.error || `فشل الطلب (HTTP ${response.status}).`)
  return body
}

export function AttendanceManagement({ username, onNavigateToPayroll }: { username: string; onNavigateToPayroll: () => void }) {
  const normalized = normalizeUsername(username)
  const payrollKey = `employee_payroll:${normalized}`
  const [payroll, setPayroll] = useState<PayrollData>(EMPTY_PAYROLL)
  const [staffPins, setStaffPins] = useState<StaffPin[]>([])
  const [pinDrafts, setPinDrafts] = useState<Record<string, string>>({})
  const [clearPins, setClearPins] = useState<Record<string, boolean>>({})
  const [linkToken, setLinkToken] = useState("")
  const [qrDataUrl, setQrDataUrl] = useState("")
  const [fromDate, setFromDate] = useState(firstDayOfMonth(localDateValue()))
  const [toDate, setToDate] = useState(localDateValue())
  const [punches, setPunches] = useState<Punch[]>([])
  const [policy, setPolicy] = useState<DayPolicy>("hours")
  const [dayOverrides, setDayOverrides] = useState<Record<string, string>>({})
  const [selectedSheetId, setSelectedSheetId] = useState("")
  const [ready, setReady] = useState(false)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState("")
  const [notice, setNotice] = useState("")
  const [copied, setCopied] = useState(false)

  useEffect(() => {
    let active = true
    const load = async () => {
      setReady(false)
      setError("")
      if (!centralStorageEnabled) {
        setError("تسجيل الحضور عبر الرابط يتطلب تفعيل التخزين المركزي والسحابي في إعدادات المنصة.")
        return
      }
      try {
        const [payrollData, kioskResponse, staffResponse] = await Promise.all([
          centralStorageGet<PayrollData>(payrollKey),
          fetch("/api/attendance/admin/kiosk", { credentials: "include" }).then(response => responseJson<{ token: string | null }>(response)),
          fetch("/api/attendance/admin/staff", { credentials: "include" }).then(response => responseJson<{ employees: StaffPin[] }>(response)),
        ])
        const validPayroll = payrollData && Array.isArray(payrollData.employees) && Array.isArray(payrollData.sheets) ? payrollData : EMPTY_PAYROLL
        if (active) {
          setPayroll(validPayroll)
          setLinkToken(kioskResponse.token ?? "")
          setStaffPins(staffResponse.employees)
          setReady(true)
        }
      } catch (cause) {
        if (active) setError(cause instanceof Error ? `تعذر تحميل منصة الحضور: ${cause.message}` : "تعذر تحميل منصة الحضور.")
      }
    }
    void load()
    return () => { active = false }
  }, [payrollKey])

  const attendanceUrl = useMemo(() => {
    if (!linkToken || typeof window === "undefined") return ""
    const url = new URL(window.location.href)
    url.search = ""
    url.hash = ""
    url.searchParams.set("attendance", linkToken)
    return url.toString()
  }, [linkToken])

  useEffect(() => {
    let active = true
    if (!attendanceUrl) {
      setQrDataUrl("")
      return
    }
    QRCode.toDataURL(attendanceUrl, { width: 420, margin: 2, errorCorrectionLevel: "H" })
      .then(dataUrl => { if (active) setQrDataUrl(dataUrl) })
      .catch(cause => {
        console.error("تعذر إنشاء رمز QR للحضور:", cause)
        if (active) setError("تعذر إنشاء صورة QR للرابط.")
      })
    return () => { active = false }
  }, [attendanceUrl])

  const configuredStaff = useMemo(() => new Map(staffPins.map(employee => [employee.employeeId, employee])), [staffPins])
  const currentStaff = useMemo(() => payroll.employees.map(employee => ({
    ...employee,
    pinConfigured: configuredStaff.get(employee.id)?.pinConfigured ?? false,
  })), [configuredStaff, payroll.employees])
  const availableSheets = useMemo(() => payroll.sheets.filter(sheet => !sheet.locked), [payroll.sheets])

  const summaries = useMemo<Summary[]>(() => {
    const employeesById = new Map(payroll.employees.map(employee => [employee.id, {
      id: employee.id,
      employeeNumber: employee.employeeNumber,
      name: employee.name,
    }]))
    for (const punch of punches) {
      if (!employeesById.has(punch.employeeId)) employeesById.set(punch.employeeId, {
        id: punch.employeeId,
        employeeNumber: punch.employeeNumber,
        name: punch.employeeName,
      })
    }
    return [...employeesById.values()].map(employee => {
      const employeePunches = punches.filter(punch => punch.employeeId === employee.id)
        .sort((a, b) => a.timestamp.localeCompare(b.timestamp))
      const byDay = new Map<string, Punch[]>()
      for (const punch of employeePunches) byDay.set(punch.date, [...(byDay.get(punch.date) ?? []), punch])
      let totalHours = 0
      let workedDays = 0
      for (const dayPunches of byDay.values()) {
        const start = dayPunches.find(punch => punch.action === "in")
        const end = dayPunches.find(punch => punch.action === "out")
        if (!start) continue
        workedDays += 1
        if (end) {
          totalHours += Math.max(0, (Date.parse(end.timestamp) - Date.parse(start.timestamp)) / 3_600_000)
        }
      }
      return {
        employeeId: employee.id,
        employeeNumber: employee.employeeNumber,
        name: employee.name,
        hours: roundDays(totalHours),
        workedDays,
        calculatedDays: policy === "full_day" ? workedDays : roundDays(totalHours / 8),
      }
    })
  }, [payroll.employees, punches, policy])

  const dailyAttendance = useMemo<DailyAttendance[]>(() => {
    const employeeById = new Map(payroll.employees.map(employee => [employee.id, {
      employeeNumber: employee.employeeNumber,
      name: employee.name,
    }]))
    for (const punch of punches) {
      if (!employeeById.has(punch.employeeId)) employeeById.set(punch.employeeId, {
        employeeNumber: punch.employeeNumber,
        name: punch.employeeName,
      })
    }
    const grouped = new Map<string, Punch[]>()
    for (const punch of punches) {
      const key = `${punch.employeeId}:${punch.date}`
      grouped.set(key, [...(grouped.get(key) ?? []), punch])
    }
    return [...grouped.entries()].map(([key, dayPunches]) => {
      const [employeeId, date] = key.split(":")
      const sorted = dayPunches.sort((a, b) => a.timestamp.localeCompare(b.timestamp))
      const start = sorted.find(punch => punch.action === "in")
      const end = sorted.find(punch => punch.action === "out")
      const employee = employeeById.get(employeeId)
      return {
        employeeId,
        employeeNumber: employee?.employeeNumber ?? "",
        name: employee?.name ?? "",
        date,
        arrival: start?.time ?? "—",
        departure: end?.time ?? "—",
        hours: start && end ? roundDays(Math.max(0, (Date.parse(end.timestamp) - Date.parse(start.timestamp)) / 3_600_000)) : 0,
      }
    }).sort((a, b) => a.date.localeCompare(b.date) || a.name.localeCompare(b.name))
  }, [payroll.employees, punches])

  const ensureEmployeePins = async () => {
    setBusy(true)
    setError("")
    setNotice("")
    try {
      const response = await fetch("/api/attendance/admin/staff", {
        method: "PUT",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          employees: payroll.employees.map(employee => ({
            employeeId: employee.id,
            employeeNumber: employee.employeeNumber,
            name: employee.name,
            pin: pinDrafts[employee.id] || undefined,
            clearPin: clearPins[employee.id] === true,
          })),
        }),
      })
      await responseJson<{ saved: boolean; configuredCount: number }>(response)
      const staffResponse = await fetch("/api/attendance/admin/staff", { credentials: "include" })
      const result = await responseJson<{ employees: StaffPin[] }>(staffResponse)
      setStaffPins(result.employees)
      setPinDrafts({})
      setClearPins({})
      setNotice("تم حفظ إعدادات الموظفين السرية. الأرقام السرية لا تُعرض بعد حفظها.")
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر حفظ أرقام الموظفين السرية: ${cause.message}` : "تعذر حفظ أرقام الموظفين السرية.")
    } finally {
      setBusy(false)
    }
  }

  const createKioskLink = async () => {
    setBusy(true)
    setError("")
    setNotice("")
    try {
      const response = await fetch("/api/attendance/admin/kiosk", {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: "{}",
      })
      const result = await responseJson<{ token: string }>(response)
      setLinkToken(result.token)
      setNotice("تم إنشاء رابط ورمز QR جديدين. أي رابط سابق أصبح غير صالح.")
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر إنشاء رابط الحضور: ${cause.message}` : "تعذر إنشاء رابط الحضور.")
    } finally {
      setBusy(false)
    }
  }

  const loadReport = async () => {
    setBusy(true)
    setError("")
    setNotice("")
    try {
      const query = new URLSearchParams({ from: fromDate, to: toDate })
      const response = await fetch(`/api/attendance/admin/records?${query}`, { credentials: "include" })
      const result = await responseJson<{ records: Punch[] }>(response)
      setPunches(result.records)
      setDayOverrides({})
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر تحميل تقرير الحضور: ${cause.message}` : "تعذر تحميل تقرير الحضور.")
    } finally {
      setBusy(false)
    }
  }

  const importToPayroll = async () => {
    const sheet = availableSheets.find(item => item.id === selectedSheetId)
    if (!sheet) {
      setError("اختر كشف رواتب مفتوحاً لترحيل الحضور إليه.")
      return
    }
    if (sheet.percentageEnabled) {
      setError("أوقف احتساب النسبة في كشف الرواتب قبل ترحيل أيام الحضور إليه.")
      return
    }
    const imported = summaries.map(summary => ({
      ...summary,
      days: Number(dayOverrides[summary.employeeId] ?? summary.calculatedDays),
    }))
    const invalid = imported.find(item => !Number.isFinite(item.days) || item.days < 0 || item.days > sheet.days)
    if (invalid) {
      setError(`عدد الأيام المعتمد للموظف ${invalid.name} يجب أن يكون بين صفر و${sheet.days} حسب إعداد الكشف.`)
      return
    }
    const advanceIssue = sheet.lines.find(line => {
      const attendance = imported.find(item => item.employeeId === line.employeeId)
      return attendance !== undefined && line.advance > line.dailyRate * attendance.days
    })
    if (advanceIssue) {
      setError(`لا يمكن الترحيل: المقبوض من ${advanceIssue.name} أكبر من راتبه بعد احتساب أيام الحضور. عدّل المقبوض أو الأيام أولاً.`)
      return
    }
    const updatedSheet: PayrollSheet = {
      ...sheet,
      lines: sheet.lines.map(line => {
        const record = imported.find(item => item.employeeId === line.employeeId)
        if (!record) return line
        const dailyRate = line.dailyRate
        const gross = Math.round((dailyRate * record.days + Number.EPSILON) * 100) / 100
        const debtDeduction = Math.round((Math.min(Math.max(0, gross - line.advance), line.debtBalance ?? 0) + Number.EPSILON) * 100) / 100
        return { ...line, dailyRate, attendanceDays: record.days, gross, debtDeduction }
      }),
    }
    const next = { ...payroll, sheets: payroll.sheets.map(item => item.id === sheet.id ? updatedSheet : item) }
    setBusy(true)
    setError("")
    setNotice("")
    try {
      if (centralStorageEnabled) await centralStorageSet(payrollKey, next)
      else localStorage.setItem(payrollKey, JSON.stringify(next))
      setPayroll(next)
      setNotice(`تم ترحيل أيام الحضور إلى كشف «${sheet.name}».`)
    } catch (cause) {
      setError(cause instanceof Error ? `تعذر ترحيل الحضور ولم يتم اعتماد التعديل: ${cause.message}` : "تعذر ترحيل الحضور.")
    } finally {
      setBusy(false)
    }
  }

  const copyLink = async () => {
    try {
      await navigator.clipboard.writeText(attendanceUrl)
      setCopied(true)
      window.setTimeout(() => setCopied(false), 2000)
    } catch (cause) {
      console.error("تعذر نسخ رابط الحضور:", cause)
      setError("تعذر نسخ الرابط تلقائياً؛ انسخه يدوياً من الحقل.")
    }
  }

  const downloadQr = () => {
    if (!qrDataUrl) return
    const anchor = document.createElement("a")
    anchor.href = qrDataUrl
    anchor.download = "attendance-qr.png"
    anchor.click()
  }

  if (!ready && centralStorageEnabled && !error) {
    return <main dir="rtl" className="min-h-full bg-slate-50 p-8 text-center text-sm text-slate-600">جارٍ تحميل إعدادات الحضور...</main>
  }

  return (
    <main dir="rtl" className="min-h-full bg-slate-50 p-4 text-slate-900 sm:p-7">
      <div className="mx-auto max-w-7xl space-y-6">
        <header className="rounded-3xl bg-slate-900 p-6 text-white shadow-lg sm:p-8">
          <div className="flex flex-wrap items-center justify-between gap-4">
            <div><div className="inline-flex items-center gap-2 rounded-full bg-white/10 px-3 py-1.5 text-xs"><CalendarDays className="h-4 w-4" />موظفو {username}</div><h1 className="mt-3 text-2xl font-extrabold">الحضور والانصراف</h1><p className="mt-2 max-w-3xl text-sm leading-6 text-slate-300">رابط وبيانات الحضور مرتبطة بحسابك وحده. لا يتم عرض قائمة موظفيك للعامة، ويحتاج كل موظف رقمه الوظيفي ورقمه السري.</p></div>
            <Users className="h-10 w-10 text-cyan-300" />
          </div>
        </header>

        {error && <div role="alert" className="rounded-xl border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-800">{error}</div>}
        {notice && <div role="status" className="rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-800">{notice}</div>}

        {centralStorageEnabled && (
          <>
            <section className="grid gap-5 lg:grid-cols-[minmax(0,1.3fr)_minmax(260px,0.7fr)]">
              <article className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
                <h2 className="flex items-center gap-2 font-bold"><Link className="h-5 w-5 text-cyan-800" />رابط خاص بحسابك</h2>
                <p className="mt-2 text-sm leading-6 text-slate-600">الباركود يحتوي رابطاً عشوائياً مرتبطاً بحسابك فقط؛ لا يحتوي بيانات الموظفين. يمكن مشاركة الرابط أو تنزيل صورة QR وطباعتها. رقم الموظف وحده لا يكفي للتسجيل.</p>
                {linkToken ? <>
                  <div className="mt-4 flex flex-wrap gap-2">
                    <button type="button" onClick={() => void copyLink()} className="inline-flex items-center gap-2 rounded-lg border px-3 py-2 text-sm font-semibold hover:bg-slate-50"><Copy className="h-4 w-4" />{copied ? "تم النسخ" : "نسخ الرابط"}</button>
                    <button type="button" onClick={downloadQr} disabled={!qrDataUrl} className="inline-flex items-center gap-2 rounded-lg bg-cyan-800 px-3 py-2 text-sm font-semibold text-white hover:bg-cyan-900 disabled:opacity-50"><Download className="h-4 w-4" />تنزيل صورة QR</button>
                    <button type="button" onClick={() => void createKioskLink()} disabled={busy} className="inline-flex items-center gap-2 rounded-lg border border-rose-200 px-3 py-2 text-sm font-semibold text-rose-800 hover:bg-rose-50 disabled:opacity-50"><RefreshCw className="h-4 w-4" />إبطال الرابط وإنشاء جديد</button>
                  </div>
                  <input aria-label="رابط تسجيل الحضور" readOnly value={attendanceUrl} onFocus={event => event.currentTarget.select()} className={`${INPUT_CLASS} mt-3`} dir="ltr" />
                  {qrDataUrl && <div className="mt-4 flex justify-center rounded-xl bg-white p-3"><img src={qrDataUrl} alt="رمز QR لرابط الحضور الخاص بالحساب" className="h-48 w-48" /></div>}
                </> : <button type="button" onClick={() => void createKioskLink()} disabled={busy} className="mt-4 inline-flex items-center gap-2 rounded-xl bg-cyan-800 px-4 py-2.5 text-sm font-bold text-white hover:bg-cyan-900 disabled:opacity-50"><Link className="h-4 w-4" />إنشاء رابط ورمز QR</button>}
              </article>
              <article className="rounded-2xl border border-amber-200 bg-amber-50 p-5">
                <h2 className="flex items-center gap-2 font-bold text-amber-950"><Shield className="h-5 w-5" />نصائح حماية</h2>
                <ul className="mt-3 list-disc space-y-2 pr-5 text-sm leading-6 text-amber-950">
                  <li>عيّن PIN مختلفاً من 6 إلى 12 رقماً لكل موظف ولا تكتبه على الباركود.</li>
                  <li>عند تسرب الرابط، أبطله وأنشئ رابطاً جديداً.</li>
                  <li>استخدم الموقع على اتصال HTTPS فقط؛ يرفض المتصفح بعض ميزات المشاركة والنسخ على اتصال غير آمن.</li>
                  <li>الرابط يحدد الحساب فقط، ولا يمنع مشاركة QR؛ PIN الموظف هو التحقق الشخصي.</li>
                </ul>
              </article>
            </section>

            <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div><h2 className="flex items-center gap-2 font-bold"><ShieldCheck className="h-5 w-5 text-emerald-700" />إعداد أرقام الموظفين السرية</h2><p className="mt-1 text-xs text-slate-500">الموظفون مأخوذون من «كشف رواتب الموظفين». أدخل PIN جديداً فقط عند تعيينه أو تغييره.</p></div>
                <button type="button" onClick={() => void ensureEmployeePins()} disabled={busy || !payroll.employees.length} className="rounded-xl bg-emerald-700 px-4 py-2.5 text-sm font-bold text-white hover:bg-emerald-800 disabled:opacity-50">حفظ إعدادات PIN</button>
              </div>
              {currentStaff.length ?               <div className="mt-4 overflow-x-auto"><table className="w-full min-w-[620px] text-sm"><thead className="bg-slate-50"><tr><th className="px-3 py-2 text-right">رقم الموظف</th><th className="px-3 py-2 text-right">الاسم</th><th className="px-3 py-2 text-right">حالة PIN</th><th className="px-3 py-2 text-right">PIN جديد</th><th className="px-3 py-2 text-right">إجراء</th></tr></thead><tbody className="divide-y">{currentStaff.map(employee => <tr key={employee.id}><td className="px-3 py-2 font-bold">{employee.employeeNumber}</td><td className="px-3 py-2">{employee.name}</td><td className="px-3 py-2">{employee.pinConfigured ? <span className="text-emerald-800">مهيأ</span> : <span className="text-rose-700">غير مهيأ</span>}</td><td className="px-3 py-2"><input type="password" inputMode="numeric" minLength={6} maxLength={12} placeholder="6–12 رقماً" value={pinDrafts[employee.id] ?? ""} onChange={event => { const pin = event.target.value.replace(/\D/g, "").slice(0, 12); setPinDrafts(current => ({ ...current, [employee.id]: pin })); if (pin) setClearPins(current => ({ ...current, [employee.id]: false })) }} className="w-36 rounded-lg border px-2 py-1.5" dir="ltr" /></td><td className="px-3 py-2"><label className="flex items-center gap-1 text-xs"><input type="checkbox" checked={clearPins[employee.id] ?? false} onChange={event => { setClearPins(current => ({ ...current, [employee.id]: event.target.checked })); if (event.target.checked) setPinDrafts(current => ({ ...current, [employee.id]: "" })) }} />مسح PIN</label></td></tr>)}</tbody></table></div> : <p className="mt-4 rounded-xl bg-slate-50 p-4 text-sm text-slate-600">لا يوجد موظفون في منصة الرواتب بعد. أضف الموظفين هناك ثم حدّث هذه الصفحة.</p>}
            </section>

            <section className="space-y-4 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <div>
                <h2 className="flex items-center gap-2 font-bold"><CalendarDays className="h-5 w-5 text-cyan-800" />تقرير الحضور وترحيله للرواتب</h2>
                <p className="mt-1 text-xs leading-5 text-slate-500">الساعات تحسب من وقت الخادم. الوضع النسبي يحسب كل 8 ساعات يوماً (وقد ينتج أجزاء يوم للرواتب)، أو اختر احتساب كل يوم فيه حركة يوماً كاملاً. يمكنك تعديل الأيام المعتمدة يدوياً قبل الترحيل.</p>
              </div>
              <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-5">
                <label className="text-xs font-semibold">من تاريخ<input type="date" value={fromDate} onChange={event => setFromDate(event.target.value)} className={`${INPUT_CLASS} mt-1`} /></label>
                <label className="text-xs font-semibold">إلى تاريخ<input type="date" value={toDate} onChange={event => setToDate(event.target.value)} className={`${INPUT_CLASS} mt-1`} /></label>
                <label className="text-xs font-semibold lg:col-span-2">طريقة احتساب الأيام<select value={policy} onChange={event => setPolicy(event.target.value as DayPolicy)} className={`${INPUT_CLASS} mt-1`}><option value="hours">حسب الساعات: كل 8 ساعات = يوم</option><option value="full_day">كل يوم فيه حضور = يوم كامل مهما كانت الساعات</option></select></label>
                <button type="button" onClick={() => void loadReport()} disabled={busy} className="self-end rounded-xl bg-slate-900 px-4 py-2.5 text-sm font-bold text-white hover:bg-slate-800 disabled:opacity-50">عرض التقرير</button>
              </div>
              <div className="overflow-x-auto"><table className="w-full min-w-[680px] text-sm"><thead className="bg-slate-50"><tr>{["رقم الموظف", "الاسم", "أيام بحضور", "مجموع الساعات", "أيام للراتب"].map(label => <th key={label} className="px-3 py-2 text-right">{label}</th>)}</tr></thead><tbody className="divide-y">{summaries.map(summary => <tr key={summary.employeeId}><td className="px-3 py-2 font-bold">{summary.employeeNumber}</td><td className="px-3 py-2">{summary.name}</td><td className="px-3 py-2">{summary.workedDays}</td><td className="px-3 py-2" dir="ltr">{formatHours(summary.hours)}</td><td className="px-3 py-2"><input type="number" min="0" max={availableSheets.find(sheet => sheet.id === selectedSheetId)?.days} step="0.01" value={dayOverrides[summary.employeeId] ?? String(summary.calculatedDays)} onChange={event => setDayOverrides(current => ({ ...current, [summary.employeeId]: event.target.value }))} className="w-24 rounded-lg border px-2 py-1.5" /></td></tr>)}</tbody></table></div>
              <details className="rounded-xl border border-slate-200">
                <summary className="cursor-pointer px-4 py-3 text-sm font-bold text-slate-700">تفاصيل يومية للحضور والانصراف ({dailyAttendance.length})</summary>
                <div className="overflow-x-auto border-t"><table className="w-full min-w-[620px] text-sm"><thead className="bg-slate-50"><tr>{["التاريخ", "رقم الموظف", "اسم الموظف", "وقت الحضور", "وقت الانصراف", "الساعات"].map(label => <th key={label} className="px-3 py-2 text-right">{label}</th>)}</tr></thead><tbody className="divide-y">{dailyAttendance.map(day => <tr key={`${day.employeeId}:${day.date}`}><td className="px-3 py-2">{day.date}</td><td className="px-3 py-2">{day.employeeNumber}</td><td className="px-3 py-2">{day.name}</td><td className="px-3 py-2" dir="ltr">{day.arrival}</td><td className="px-3 py-2" dir="ltr">{day.departure}</td><td className="px-3 py-2" dir="ltr">{formatHours(day.hours)}</td></tr>)}</tbody></table>{!dailyAttendance.length && <p className="p-4 text-center text-xs text-slate-500">لا توجد حركات في الفترة المحددة.</p>}</div>
              </details>
              <div className="grid gap-3 sm:grid-cols-[minmax(0,1fr)_auto]">
                <label className="text-xs font-semibold">كشف الرواتب المفتوح<select value={selectedSheetId} onChange={event => setSelectedSheetId(event.target.value)} className={`${INPUT_CLASS} mt-1`}><option value="">اختر كشفاً...</option>{availableSheets.map(sheet => <option key={sheet.id} value={sheet.id}>{sheet.name} · حتى {sheet.days} يوم · {sheet.periodFrom} – {sheet.periodTo}</option>)}</select></label>
                <button type="button" onClick={() => void importToPayroll()} disabled={busy || !summaries.length} className="inline-flex items-center justify-center gap-2 self-end rounded-xl bg-cyan-800 px-4 py-2.5 text-sm font-bold text-white hover:bg-cyan-900 disabled:opacity-50"><ArrowRightLeft className="h-4 w-4" />ترحيل أيام الحضور</button>
              </div>
              {availableSheets.length === 0 && <p className="rounded-lg bg-amber-50 p-3 text-xs text-amber-900">أنشئ كشف رواتب غير مثبت أولاً. يجب أن يسمح عدد أيام الكشف بأيام الحضور المعتمدة.</p>}
              {notice.includes("ترحيل") && <button type="button" onClick={onNavigateToPayroll} className="inline-flex items-center gap-2 text-sm font-bold text-cyan-900 underline"><ArrowRightLeft className="h-4 w-4" />فتح كشف الرواتب المحدث</button>}
            </section>
          </>
        )}
      </div>
    </main>
  )
}
