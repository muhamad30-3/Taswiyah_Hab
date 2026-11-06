"use client"

import { useState, type FormEvent } from "react"
import { ArrowDownToLine, ArrowUpFromLine, CheckCircle2, LoaderCircle, ShieldCheck } from "lucide-react"

type PunchAction = "in" | "out"

function normalizeDigits(value: string) {
  return value.replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, digit => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)))
}

export function AttendanceKiosk({ token }: { token: string }) {
  const [employeeNumber, setEmployeeNumber] = useState("")
  const [pin, setPin] = useState("")
  const [action, setAction] = useState<PunchAction>("in")
  const [busy, setBusy] = useState(false)
  const [message, setMessage] = useState("")
  const [error, setError] = useState("")

  const submitPunch = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    setBusy(true)
    setMessage("")
    setError("")
    try {
      const response = await fetch(`/api/attendance/public/${encodeURIComponent(token)}/punch`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ employeeNumber: employeeNumber.trim(), pin, action }),
      })
      const result = await response.json() as { error?: string; employeeName?: string; date?: string; time?: string }
      if (!response.ok) throw new Error(result.error || "تعذر تسجيل الحركة.")
      setMessage(`تم تسجيل ${action === "in" ? "الحضور" : "الانصراف"} للموظف ${result.employeeName} في ${result.time} بتاريخ ${result.date}.`)
      setPin("")
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذر الاتصال بالخادم. تحقق من الإنترنت وحاول مجدداً.")
    } finally {
      setBusy(false)
    }
  }

  return (
    <main dir="rtl" className="flex min-h-[100dvh] items-center justify-center bg-slate-100 p-4">
      <section className="w-full max-w-md rounded-3xl border border-slate-200 bg-white p-6 shadow-xl sm:p-8">
        <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-2xl bg-cyan-50 text-cyan-800"><ShieldCheck className="h-7 w-7" /></div>
        <h1 className="mt-4 text-center text-2xl font-extrabold text-slate-900">تسجيل حضور الموظفين</h1>
        <p className="mt-2 text-center text-sm text-slate-600">أدخل رقمك الوظيفي ورقمك السري الشخصي لتسجيل الحركة.</p>
        {message ? (
          <div role="status" className="mt-6 rounded-2xl border border-emerald-200 bg-emerald-50 p-4 text-center text-sm font-semibold leading-7 text-emerald-900">
            <CheckCircle2 className="mx-auto mb-2 h-7 w-7 text-emerald-700" />{message}
            <button type="button" onClick={() => { setMessage(""); setEmployeeNumber(""); setAction("in") }} className="mt-4 block w-full rounded-xl border border-emerald-300 bg-white px-4 py-2.5 text-sm font-bold text-emerald-900 hover:bg-emerald-100">تسجيل موظف آخر</button>
          </div>
        ) : (
          <form onSubmit={event => void submitPunch(event)} className="mt-6 space-y-4">
            <label className="block text-sm font-semibold text-slate-700">رقم الموظف
              <input required autoFocus inputMode="numeric" autoComplete="username" value={employeeNumber} onChange={event => setEmployeeNumber(event.target.value)} className="mt-1.5 w-full rounded-xl border border-slate-300 px-3 py-3 text-lg font-bold tracking-wide outline-none focus:border-cyan-700 focus:ring-2 focus:ring-cyan-100" dir="ltr" />
            </label>
            <label className="block text-sm font-semibold text-slate-700">الرقم السري الشخصي
              <input required type="password" inputMode="numeric" autoComplete="current-password" minLength={6} maxLength={12} value={pin} onChange={event => setPin(normalizeDigits(event.target.value).replace(/\D/g, "").slice(0, 12))} className="mt-1.5 w-full rounded-xl border border-slate-300 px-3 py-3 text-lg tracking-[0.3em] outline-none focus:border-cyan-700 focus:ring-2 focus:ring-cyan-100" dir="ltr" />
            </label>
            <div className="grid grid-cols-2 gap-2 rounded-xl bg-slate-100 p-1.5">
              <button type="button" onClick={() => setAction("in")} className={`flex items-center justify-center gap-2 rounded-lg px-3 py-2.5 text-sm font-bold ${action === "in" ? "bg-emerald-700 text-white shadow" : "text-slate-600 hover:bg-white"}`}><ArrowDownToLine className="h-4 w-4" />تسجيل حضور</button>
              <button type="button" onClick={() => setAction("out")} className={`flex items-center justify-center gap-2 rounded-lg px-3 py-2.5 text-sm font-bold ${action === "out" ? "bg-rose-700 text-white shadow" : "text-slate-600 hover:bg-white"}`}><ArrowUpFromLine className="h-4 w-4" />تسجيل انصراف</button>
            </div>
            {error && <div role="alert" className="rounded-xl border border-rose-200 bg-rose-50 px-3 py-2.5 text-sm text-rose-800">{error}</div>}
            <button type="submit" disabled={busy} className="flex w-full items-center justify-center gap-2 rounded-xl bg-cyan-800 px-4 py-3 font-bold text-white hover:bg-cyan-900 disabled:cursor-wait disabled:opacity-60">
              {busy && <LoaderCircle className="h-4 w-4 animate-spin" />}
              {busy ? "جارٍ التحقق والتسجيل..." : "تأكيد التسجيل"}
            </button>
            <p className="text-center text-xs leading-5 text-slate-500">وقت الجهاز لا يُستخدم؛ وقت التسجيل يؤخذ من الخادم. لا تشارك رقمك السري مع الآخرين.</p>
          </form>
        )}
      </section>
    </main>
  )
}
