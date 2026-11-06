"use client"

import { createContext, useCallback, useContext, useEffect, useId, useState, type ReactNode } from "react"
import { TASWIYAH_LOGO_PATH } from "../lib/report-branding"

type LoadingIndicator = {
  message: string
  detail?: string
  progress?: number
}

type LoadingContextValue = (id: string, indicator: LoadingIndicator | null) => void

const LoadingContext = createContext<LoadingContextValue | null>(null)

export function GlobalLoadingProvider({ children }: { children: ReactNode }) {
  const [indicators, setIndicators] = useState<Map<string, LoadingIndicator>>(() => new Map())
  const updateIndicator = useCallback<LoadingContextValue>((id, indicator) => {
    setIndicators(current => {
      if (!indicator && !current.has(id)) return current
      const next = new Map(current)
      if (indicator) next.set(id, indicator)
      else next.delete(id)
      return next
    })
  }, [])
  const activeIndicator = [...indicators.values()].pop()

  return (
    <LoadingContext.Provider value={updateIndicator}>
      {children}
      {activeIndicator && <GlobalLoadingOverlay indicator={activeIndicator} />}
    </LoadingContext.Provider>
  )
}

export function useGlobalLoading(
  active: boolean,
  message: string,
  detail?: string,
  progress?: number,
) {
  const updateIndicator = useContext(LoadingContext)
  const id = useId()

  useEffect(() => {
    if (!updateIndicator) return
    updateIndicator(id, active ? { message, detail, progress } : null)
    return () => updateIndicator(id, null)
  }, [active, detail, id, message, progress, updateIndicator])
}

function GlobalLoadingOverlay({ indicator }: { indicator: LoadingIndicator }) {
  const progress = indicator.progress === undefined
    ? null
    : Math.max(0, Math.min(100, indicator.progress))

  return (
    <div
      className="taswiyah-loading-overlay fixed inset-0 z-[200] flex items-center justify-center bg-[#f4f7fb]/95 px-5 py-10 backdrop-blur-[3px]"
      dir="rtl"
      role="status"
      aria-live="polite"
    >
      <section className="w-full max-w-md rounded-3xl border border-slate-200/80 bg-white/95 px-7 py-9 text-center shadow-2xl shadow-slate-900/15 sm:px-10">
        <img
          src={TASWIYAH_LOGO_PATH}
          alt="Taswiyah Hub"
          className="mx-auto h-28 w-28 rounded-3xl object-contain p-2"
        />
        <h2 className="mt-6 text-lg font-bold text-slate-800">{indicator.message}</h2>
        {indicator.detail && <p className="mt-2 text-sm text-slate-500">{indicator.detail}</p>}
        <div className="relative mt-9">
          <div className="absolute inset-x-0 -top-5 flex items-center justify-around px-7" aria-hidden="true">
            <span className="taswiyah-loading-dot h-3 w-3 rounded-full bg-blue-600" />
            <span className="taswiyah-loading-dot h-3 w-3 rounded-full bg-emerald-500" />
            <span className="taswiyah-loading-dot h-3 w-3 rounded-full bg-amber-400" />
          </div>
          <div
            className="relative h-2.5 overflow-hidden rounded-full bg-slate-100"
            {...(progress === null ? {} : {
              role: "progressbar",
              "aria-label": "تقدّم التحميل",
              "aria-valuemin": 0,
              "aria-valuemax": 100,
              "aria-valuenow": progress,
            })}
          >
            {progress === null ? (
              <div className="taswiyah-loading-indeterminate h-full rounded-full bg-gradient-to-l from-blue-600 via-emerald-500 to-amber-400" />
            ) : (
              <div
                className="h-full rounded-full bg-gradient-to-l from-blue-600 via-emerald-500 to-amber-400 transition-[width] duration-300"
                style={{ width: `${progress}%` }}
              />
            )}
          </div>
        </div>
        {!indicator.detail && <p className="mt-4 text-xs font-medium text-slate-400">يُرجى الانتظار...</p>}
      </section>
    </div>
  )
}
