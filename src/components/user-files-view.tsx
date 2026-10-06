"use client"

import { useCallback, useEffect, useState } from "react"
import { Download, File, RefreshCw, Trash2, Upload } from "lucide-react"
import { deleteStoredFile, downloadStoredFile, listStoredFiles, uploadStoredFile, type StoredFile } from "../lib/central-storage"

function formatBytes(size: number) {
  if (size < 1024) return `${size} بايت`
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} كيلوبايت`
  return `${(size / (1024 * 1024)).toFixed(1)} ميغابايت`
}

export function UserFilesView() {
  const [files, setFiles] = useState<StoredFile[]>([])
  const [busy, setBusy] = useState(false)
  const [message, setMessage] = useState("")
  const [error, setError] = useState("")

  const refresh = useCallback(async () => {
    setBusy(true)
    setError("")
    try {
      setFiles(await listStoredFiles())
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذر تحميل الملفات.")
    } finally {
      setBusy(false)
    }
  }, [])

  useEffect(() => { void refresh() }, [refresh])

  const upload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const selected = Array.from(event.target.files ?? [])
    event.target.value = ""
    if (!selected.length) return
    setBusy(true)
    setMessage("")
    setError("")
    let uploadedCount = 0
    try {
      for (const file of selected) {
        await uploadStoredFile(file)
        uploadedCount++
      }
      setMessage(`تم حفظ ${selected.length} ملف${selected.length === 1 ? "" : "ات"} في مكتبتك.`)
      await refresh()
    } catch (cause) {
      if (uploadedCount > 0) {
        await refresh()
        setMessage(`تم حفظ ${uploadedCount} من أصل ${selected.length} ملفات قبل تعذر إكمال الرفع.`)
      }
      setError(cause instanceof Error ? cause.message : "تعذر رفع الملف.")
    } finally {
      setBusy(false)
    }
  }

  const remove = async (file: StoredFile) => {
    if (!window.confirm(`هل تريد حذف «${file.fileName}» نهائياً من مكتبتك؟`)) return
    setError("")
    try {
      await deleteStoredFile(file.id)
      setFiles(current => current.filter(item => item.id !== file.id))
      setMessage("تم حذف الملف.")
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذر حذف الملف.")
    }
  }

  const download = async (file: StoredFile) => {
    setError("")
    try {
      await downloadStoredFile(file)
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "تعذر تنزيل الملف.")
    }
  }

  return (
    <section className="mx-auto max-w-5xl p-5 sm:p-8" dir="rtl">
      <header className="mb-6 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold">ملفاتي</h1>
          <p className="mt-2 text-sm text-muted-foreground">ملفاتك محفوظة في حسابك وحده، ويمكنك تنزيلها وإعادة فتحها في أدوات المنصة. ملفات المشروع تضاف هنا عند حفظ المشروع.</p>
        </div>
        <label className="inline-flex cursor-pointer items-center gap-2 rounded-xl bg-primary px-4 py-2.5 text-sm font-semibold text-primary-foreground hover:opacity-90">
          <Upload className="h-4 w-4" />
          {busy ? "جارٍ المعالجة..." : "رفع ملفات"}
          <input
            type="file"
            multiple
            disabled={busy}
            onChange={event => void upload(event)}
            accept=".xlsx,.xls,.xlsm,.xlsb,.csv,.pdf,.doc,.docx,.docm,.dotx,.odt,.rtf,.ppt,.pptx,.pptm,.ppsx,.odp,.txt,.md,.html,.htm,.png,.jpg,.jpeg,.webp,.gif,.tif,.tiff,.bmp"
            className="sr-only"
          />
        </label>
      </header>

      <p className="mb-4 rounded-xl border bg-card px-4 py-3 text-xs text-muted-foreground">الصيغ المدعومة تشمل Excel وCSV وPDF وWord وPowerPoint والصور. الحد الأقصى للملف الواحد 25 ميغابايت.</p>
      {message && <p role="status" className="mb-4 rounded-lg bg-emerald-50 px-4 py-3 text-sm text-emerald-800">{message}</p>}
      {error && <p role="alert" className="mb-4 rounded-lg bg-destructive/10 px-4 py-3 text-sm text-destructive">{error}</p>}

      <div className="overflow-hidden rounded-2xl border bg-card">
        <div className="flex items-center justify-between border-b px-4 py-3">
          <h2 className="font-semibold">المحفوظات ({files.length})</h2>
          <button type="button" onClick={() => void refresh()} disabled={busy} className="rounded-lg p-2 text-muted-foreground hover:bg-muted" aria-label="تحديث قائمة الملفات">
            <RefreshCw className={`h-4 w-4 ${busy ? "animate-spin" : ""}`} />
          </button>
        </div>
        {files.length === 0 ? (
          <div className="p-10 text-center text-sm text-muted-foreground">{busy ? "جارٍ تحميل الملفات..." : "لم ترفع ملفات بعد."}</div>
        ) : (
          <ul className="divide-y">
            {files.map(file => (
              <li key={file.id} className="flex flex-wrap items-center gap-3 px-4 py-4">
                <File className="h-5 w-5 shrink-0 text-primary" />
                <div className="min-w-0 flex-1">
                  <div className="truncate text-sm font-semibold">{file.fileName}</div>
                  <div className="mt-1 text-xs text-muted-foreground">
                    {formatBytes(file.size)} · {new Date(file.uploadedAt).toLocaleString("ar")}
                  </div>
                </div>
                <button type="button" onClick={() => void download(file)} className="inline-flex items-center gap-1.5 rounded-lg border px-3 py-2 text-xs font-semibold hover:bg-muted">
                  <Download className="h-4 w-4" /> تنزيل للعمل عليه
                </button>
                <button type="button" onClick={() => void remove(file)} className="rounded-lg p-2 text-destructive hover:bg-destructive/10" aria-label={`حذف ${file.fileName}`}>
                  <Trash2 className="h-4 w-4" />
                </button>
              </li>
            ))}
          </ul>
        )}
      </div>
    </section>
  )
}
