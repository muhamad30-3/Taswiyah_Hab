"use client"

import { useEffect, useState } from "react"
import { Lock, User, Eye, EyeOff, ArrowLeft } from "lucide-react"
import { centralLogin, centralRegister, centralStorageEnabled, type CentralUser } from "../lib/central-storage"
import { passwordPolicyError } from "../lib/password-policy"

export function LoginScreen({ onLogin, initialError }: { onLogin: (user: CentralUser) => void; initialError?: string }) {
  const [username, setUsername] = useState("")
  const [email, setEmail] = useState("")
  const [password, setPassword] = useState("")
  const [showPassword, setShowPassword] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [notice, setNotice] = useState<string | null>(null)
  const [showDefaultCredentials, setShowDefaultCredentials] = useState(false)
  const [registrationMode, setRegistrationMode] = useState(false)
  const [busy, setBusy] = useState(false)

  useEffect(() => {
    if (initialError) setError(initialError)
    if (!centralStorageEnabled) setShowDefaultCredentials(!localStorage.getItem("recon_credentials"))
  }, [initialError])

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setError(null)
    setNotice(null)
    if (centralStorageEnabled) {
      if (registrationMode) {
        const passwordError = passwordPolicyError(password)
        if (passwordError) {
          setError(passwordError)
          return
        }
      }
      setBusy(true)
      try {
        if (registrationMode) {
          await centralRegister(username.trim(), email.trim(), password)
          setRegistrationMode(false)
          setPassword("")
          setNotice("تم إنشاء الحساب. سيصبح الدخول متاحاً بعد موافقة المسؤول.")
        } else {
          onLogin(await centralLogin(username.trim(), password))
        }
      } catch (cause) {
        setError(cause instanceof Error ? cause.message : "تعذر تسجيل الدخول إلى الخادم.")
      } finally {
        setBusy(false)
      }
      return
    }

    const stored = (() => {
      try {
        const raw = localStorage.getItem("recon_credentials")
        return raw ? (JSON.parse(raw) as { username: string; password: string }) : null
      } catch {
        return null
      }
    })()

    const creds = stored ?? { username: "admin", password: "admin123" }

    if (username.trim() === creds.username && password === creds.password) {
      localStorage.setItem("recon_auth", JSON.stringify({ username: creds.username, at: Date.now() }))
      onLogin({ username: creds.username, role: "admin" })
    } else {
      setError("اسم المستخدم أو كلمة المرور غير صحيحة")
    }
  }

  return (
    <main dir="rtl" className="relative flex min-h-[100dvh] items-center justify-center overflow-hidden bg-[hsl(216_48%_15%)] p-4">
      <div className="pointer-events-none absolute -left-24 -top-24 h-72 w-72 rounded-full bg-[hsl(43_96%_52%/.13)] blur-3xl" />
      <div className="pointer-events-none absolute -bottom-32 -right-20 h-96 w-96 rounded-full bg-[hsl(215_86%_35%/.25)] blur-3xl" />
      <div className="relative grid w-full max-w-4xl overflow-hidden rounded-[2rem] border border-white/10 bg-[hsl(216_40%_18%/.88)] shadow-2xl md:grid-cols-[1.05fr_.95fr]">
        <div className="hidden flex-col justify-between bg-[linear-gradient(145deg,hsl(215_86%_35%),hsl(216_48%_15%))] p-10 text-white md:flex">
          <div>
            <img src="/taswiyah-logo.png" alt="Taswiyah Hub" className="h-28 w-28 rounded-3xl bg-white/95 object-contain p-2" />
            <p className="mt-10 max-w-xs text-3xl font-bold leading-[1.45]">قرار مالي واضح، من أول كشف.</p>
            <p className="mt-4 max-w-sm text-sm leading-7 text-white/70">مساحة عمل يومية لمراجعة الفواتير والحوالات، مع أثر واضح لكل مطابقة.</p>
          </div>
          <div className="flex items-center gap-2 text-xs text-white/55"><span className="h-2 w-2 rounded-full bg-[hsl(43_96%_52%)]" /> {centralStorageEnabled ? "بياناتك تحفظ في قاعدة البيانات المركزية" : "بياناتك محفوظة محلياً في هذا المتصفح"}</div>
        </div>
        <div className="p-6 sm:p-10">
          <div className="mb-8 flex items-center gap-3 md:hidden">
            <img src="/taswiyah-logo.png" alt="Taswiyah Hub" className="h-14 w-14 rounded-2xl bg-white object-contain p-1" />
            <div><div className="text-lg font-bold text-white">Taswiyah Hub</div><div className="text-xs text-white/60">محطتك الأولى لتسوية الحسابات بدقة.</div></div>
          </div>
          <div className="mb-8">
            <div className="mb-2 text-xs font-semibold tracking-[.15em] text-[hsl(43_96%_52%)]">TASWIYAH HUB</div>
            <h1 className="text-2xl font-bold text-white">{registrationMode ? "إنشاء حساب جديد" : "أهلاً بك في مكتبك المالي"}</h1>
            <p className="mt-2 text-sm text-white/55">{registrationMode ? "سيراجع المسؤول طلبك قبل تفعيل الحساب" : "سجّل الدخول للمتابعة إلى مساحة التسوية"}</p>
          </div>

        <form
          onSubmit={handleSubmit}
          className="space-y-4"
        >
          <div className="space-y-1.5">
            <label htmlFor="username" className="text-xs font-medium text-muted-foreground">
              اسم المستخدم
            </label>
            <div className="relative">
              <User className="pointer-events-none absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
              <input
                id="username"
                type="text"
                autoComplete="username"
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                required
                minLength={3}
                maxLength={32}
                className="w-full rounded-xl border border-white/10 bg-white/[.07] py-3 pr-9 pl-3 text-sm text-white outline-none placeholder:text-white/30 focus:border-[hsl(43_96%_52%)] focus:ring-1 focus:ring-[hsl(43_96%_52%/.25)]"
                placeholder={registrationMode ? "اسم المستخدم" : "اسم المستخدم"}
              />
            </div>
          </div>

          {registrationMode && (
            <div className="space-y-1.5">
              <label htmlFor="email" className="text-xs font-medium text-muted-foreground">البريد الإلكتروني</label>
              <input
                id="email"
                type="email"
                autoComplete="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                className="w-full rounded-xl border border-white/10 bg-white/[.07] px-3 py-3 text-sm text-white outline-none placeholder:text-white/30 focus:border-[hsl(43_96%_52%)] focus:ring-1 focus:ring-[hsl(43_96%_52%/.25)]"
                placeholder="name@example.com"
                required
              />
            </div>
          )}

          <div className="space-y-1.5">
            <label htmlFor="password" className="text-xs font-medium text-muted-foreground">
              كلمة المرور
            </label>
            <div className="relative">
              <Lock className="pointer-events-none absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
              <input
                id="password"
                type={showPassword ? "text" : "password"}
                autoComplete={registrationMode ? "new-password" : "current-password"}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                minLength={registrationMode ? 6 : undefined}
                maxLength={256}
                required
              className="w-full rounded-xl border border-white/10 bg-white/[.07] py-3 pr-9 pl-9 text-sm text-white outline-none placeholder:text-white/30 focus:border-[hsl(43_96%_52%)] focus:ring-1 focus:ring-[hsl(43_96%_52%/.25)]"
                placeholder="••••••••"
              />
              <button
                type="button"
                onClick={() => setShowPassword((s) => !s)}
                className="absolute left-2 top-1/2 -translate-y-1/2 rounded p-1 text-muted-foreground hover:bg-muted"
                aria-label={showPassword ? "إخفاء كلمة المرور" : "إظهار كلمة المرور"}
              >
                {showPassword ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
              </button>
            </div>
            {registrationMode && <p className="text-xs text-white/60">6 محارف على الأقل؛ يمكنك استخدام 6 أرقام.</p>}
          </div>

          {error && (
            <p className="rounded-lg bg-destructive/10 px-3 py-2 text-xs font-medium text-destructive">
              {error}
            </p>
          )}
          {notice && <p role="status" className="rounded-lg bg-emerald-500/10 px-3 py-2 text-xs font-medium text-emerald-300">{notice}</p>}

          <button
            type="submit"
            disabled={busy}
            className="group flex w-full items-center justify-center gap-2 rounded-xl bg-[hsl(43_96%_52%)] py-3 text-sm font-bold text-[hsl(216_48%_15%)] transition-colors hover:bg-[hsl(43_96%_60%)]"
          >
            {busy ? "جارٍ التحقق..." : registrationMode ? "إرسال طلب إنشاء الحساب" : "تسجيل الدخول"} <ArrowLeft className="h-4 w-4 transition-transform group-hover:-translate-x-1" />
          </button>

          {centralStorageEnabled && (
            <button
              type="button"
              onClick={() => { setRegistrationMode(value => !value); setError(null); setNotice(null) }}
              className="w-full rounded-xl border border-white/15 py-3 text-sm font-semibold text-white/80 transition-colors hover:bg-white/5"
            >
              {registrationMode ? "لديك حساب؟ تسجيل الدخول" : "إنشاء حساب جديد"}
            </button>
          )}

          {showDefaultCredentials && !centralStorageEnabled && (
            <p className="text-center text-xs text-muted-foreground">
              الحساب الافتراضي: admin / admin123
            </p>
          )}
        </form>
      </div>
      </div>
    </main>
  )
}
