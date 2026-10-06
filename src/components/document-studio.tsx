"use client"

import { useEffect, useMemo, useRef, useState } from "react"
import * as XLSX from "xlsx-js-style"
import { BorderStyle, Document, HeadingLevel, Packer, Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, AlignmentType, WidthType } from "docx"
import PptxGenJS from "pptxgenjs"
import { createWorker, PSM } from "tesseract.js"
import {
  AlertCircle, BookOpen, Check, CheckCircle2, ClipboardPaste, Download, FileImage, FileSpreadsheet,
  FileText, FileUp, ImagePlus, Languages, LoaderCircle, Presentation, ScanLine, ScanText,
  Sparkles, Upload, WandSparkles, X, CalendarClock,
} from "lucide-react"
import { extractDocument, type ExtractedDocument } from "../lib/document-extractor"
import { readUploadedDocument } from "../lib/document-reader"
import { appendExtractedDataToWorkbook, CONTACT_WORKSHEET_NAME, FINANCIAL_WORKSHEET_NAME, styleWorksheetTable } from "../lib/workbook-append"

type InputMode = "text" | "image" | "document"
type ExportFormat = "excel" | "word" | "pdf" | "powerpoint"
type AiAction = "translate" | "summarize" | "glossary" | "format" | "ask"

const formatOptions: Array<{ id: ExportFormat; label: string; description: string; icon: typeof FileText }> = [
  { id: "excel", label: "Excel", description: "ملف جديد أو إلحاق بملفك الحالي", icon: FileSpreadsheet },
  { id: "word", label: "Word", description: "تقرير منسق وقابل للتعديل", icon: FileText },
  { id: "pdf", label: "PDF", description: "نسخة جاهزة للطباعة والمشاركة", icon: Download },
  { id: "powerpoint", label: "PowerPoint", description: "عرض تقديمي موجز", icon: Presentation },
]

const contactColumns = ["اسم الشخص", "البريد الإلكتروني", "رقم الهاتف", "اسم المرسل"]
const contactDateTimeColumns = [...contactColumns, "تاريخ الرسالة", "وقت الرسالة"]

function safeFileName(value: string) {
  const cleaned = value.replace(/[<>:"/\\|?*\u0000-\u001f]/g, "").trim().slice(0, 50)
  return cleaned || "مستند-مستخرج"
}

function downloadBlob(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement("a")
  anchor.href = url
  anchor.download = name
  anchor.click()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}

function htmlEscape(value: string) {
  return value.replace(/[&<>"']/g, character => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  })[character] ?? character)
}

function numericValue(value: string) {
  const match = value.replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit))).match(/[-+]?\d[\d,]*(?:\.\d+)?/)
  if (!match) return null
  const number = Number(match[0].replace(/,/g, ""))
  return Number.isFinite(number) ? number : null
}

function formatOcrText(
  lines: Array<{
    text: string
    bbox: { y0: number; y1: number }
    words: Array<{ text: string; bbox: { x0: number; x1: number } }>
  }>,
  fallback: string,
) {
  const rows = lines.map(line => {
    const isRightToLeft = /[\u0600-\u06ff]/.test(line.text)
    const words = line.words.filter(word => word.text.trim()).sort((a, b) =>
      isRightToLeft ? b.bbox.x0 - a.bbox.x0 : a.bbox.x0 - b.bbox.x0)
    if (words.length < 2) return [line.text.trim()]
    const gapThreshold = Math.max(14, (line.bbox.y1 - line.bbox.y0) * 1.25)
    const cells: string[] = []
    let cell = words[0].text
    let previousBoundary = isRightToLeft ? words[0].bbox.x0 : words[0].bbox.x1
    for (const word of words.slice(1)) {
      const gap = isRightToLeft ? previousBoundary - word.bbox.x1 : word.bbox.x0 - previousBoundary
      if (gap > gapThreshold) {
        cells.push(cell.trim())
        cell = word.text
      } else {
        cell += ` ${word.text}`
      }
      previousBoundary = isRightToLeft ? word.bbox.x0 : word.bbox.x1
    }
    cells.push(cell.trim())
    return cells.length >= 3 ? cells : [line.text.trim()]
  }).filter(row => row.some(cell => cell.trim()))

  const tableRows = rows.filter(row => row.length >= 3)
  const columnCount = Math.max(0, ...tableRows.map(row => row.length))
  if (tableRows.length < 3 || columnCount < 3) return fallback

  const commonHeaders = columnCount === 4
    ? ["التاريخ", "الوقت", "نوع الحركة", "المبلغ"]
    : columnCount === 5
      ? ["التاريخ", "الوقت", "نوع الحركة", "التفاصيل", "المبلغ"]
      : ["التاريخ", "الوقت", "نوع الحركة", "اسم المحصل", "اسم العميل", "المبلغ"]
  const headers = Array.from({ length: columnCount }, (_, index) => commonHeaders[index] ?? `حقل ${index + 1}`)
  return [headers, ...tableRows.map(row => Array.from({ length: columnCount }, (_, index) => row[index] ?? ""))]
    .map(row => row.join("\t"))
    .join("\n")
}

function formatAllowed(format: ExportFormat, result: ExtractedDocument, text: string) {
  if (format === "excel") return result.rows.length > 0 || result.facts.length > 0 || result.contacts.length > 0
  return text.trim().length > 0
}

function removeMessageTimestamps(text: string) {
  return text.split("\n").map(line => line.replace(/\[([^\]]{5,80})\]\s*/g, (match, timestamp: string) => {
    const isTimestamp = /\d{1,2}:\d{2}/.test(timestamp) ||
      /\d{4}[/-]\d{1,2}[/-]\d{1,2}/.test(timestamp) ||
      /\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/.test(timestamp)
    return isTimestamp ? "" : match
  }).trimEnd()).join("\n").trim()
}

function contactsAsText(result: ExtractedDocument, keepTimestamps: boolean) {
  const headers = [...contactColumns, ...(keepTimestamps ? ["تاريخ الرسالة", "وقت الرسالة"] : [])]
  const rows = result.contacts.map(contact => [
    contact.name,
    contact.email,
    contact.phone,
    contact.sender,
    ...(keepTimestamps ? [contact.date, contact.time] : []),
  ])
  return [headers, ...rows].map(row => row.join("\t")).join("\n")
}

async function cleanPageImage(file: File): Promise<Blob> {
  const image = await createImageBitmap(file)
  const source = document.createElement("canvas")
  source.width = image.width
  source.height = image.height
  const sourceContext = source.getContext("2d", { willReadFrequently: true })
  if (!sourceContext) throw new Error("تعذر تجهيز الصورة في هذا المتصفح.")
  sourceContext.drawImage(image, 0, 0)
  image.close()
  const analysisScale = Math.min(1, 640 / Math.max(source.width, source.height))
  const analysis = document.createElement("canvas")
  analysis.width = Math.max(1, Math.round(source.width * analysisScale))
  analysis.height = Math.max(1, Math.round(source.height * analysisScale))
  const analysisContext = analysis.getContext("2d", { willReadFrequently: true })
  if (!analysisContext) throw new Error("تعذر تحليل حدود الصفحة.")
  analysisContext.drawImage(source, 0, 0, analysis.width, analysis.height)
  const analysisPixels = analysisContext.getImageData(0, 0, analysis.width, analysis.height).data
  const pixelCount = analysis.width * analysis.height
  const paperMask = new Uint8Array(pixelCount)
  const visited = new Uint8Array(pixelCount)
  const queue = new Uint32Array(pixelCount)
  for (let index = 0; index < pixelCount; index++) {
    const offset = index * 4
    const red = analysisPixels[offset]
    const green = analysisPixels[offset + 1]
    const blue = analysisPixels[offset + 2]
    const luminance = red * 0.299 + green * 0.587 + blue * 0.114
    if (luminance > 170 && Math.max(red, green, blue) - Math.min(red, green, blue) < 40) paperMask[index] = 1
  }
  let pageSize = 0
  let pageCorners: Array<{ score: number; x: number; y: number }> = []
  for (let index = 0; index < pixelCount; index++) {
    if (!paperMask[index] || visited[index]) continue
    let head = 0
    let tail = 0
    queue[tail++] = index
    visited[index] = 1
    const componentCorners = [
      { score: Number.POSITIVE_INFINITY, x: 0, y: 0 },
      { score: Number.NEGATIVE_INFINITY, x: 0, y: 0 },
      { score: Number.NEGATIVE_INFINITY, x: 0, y: 0 },
      { score: Number.POSITIVE_INFINITY, x: 0, y: 0 },
    ]
    while (head < tail) {
      const current = queue[head++]
      const x = current % analysis.width
      const y = Math.floor(current / analysis.width)
      const scores = [x + y, x - y, x + y, x - y]
      for (let corner = 0; corner < componentCorners.length; corner++) {
        const isMinimum = corner === 0 || corner === 3
        if (isMinimum ? scores[corner] < componentCorners[corner].score : scores[corner] > componentCorners[corner].score) {
          componentCorners[corner] = { score: scores[corner], x, y }
        }
      }
      const neighbors = [
        x > 0 ? current - 1 : -1,
        x + 1 < analysis.width ? current + 1 : -1,
        y > 0 ? current - analysis.width : -1,
        y + 1 < analysis.height ? current + analysis.width : -1,
      ]
      for (const neighbor of neighbors) {
        if (neighbor < 0 || !paperMask[neighbor] || visited[neighbor]) continue
        visited[neighbor] = 1
        queue[tail++] = neighbor
      }
    }
    if (tail > pageSize) {
      pageSize = tail
      pageCorners = componentCorners
    }
  }
  if (pageSize < pixelCount * 0.08) {
    throw new Error("لم أتمكن من تمييز حدود الورقة بوضوح. صوّر الصفحة كاملة على خلفية داكنة ومتباينة.")
  }
  const corners = pageCorners.map(({ x, y }) => ({ x: x / analysisScale, y: y / analysisScale }))
  const [topLeft, topRight, bottomRight, bottomLeft] = corners.map(({ x, y }) => ({
    x: x / analysisScale,
    y: y / analysisScale,
  }))
  const pageWidth = (Math.hypot(topRight.x - topLeft.x, topRight.y - topLeft.y)
    + Math.hypot(bottomRight.x - bottomLeft.x, bottomRight.y - bottomLeft.y)) / 2
  const pageHeight = (Math.hypot(bottomLeft.x - topLeft.x, bottomLeft.y - topLeft.y)
    + Math.hypot(bottomRight.x - topRight.x, bottomRight.y - topRight.y)) / 2
  if (pageWidth < 100 || pageHeight < 100) throw new Error("حدود الصفحة المكتشفة صغيرة جداً. أعد تصوير الورقة.")

  const portrait = pageHeight >= pageWidth
  const outputWidth = portrait ? 1240 : 1754
  const outputHeight = portrait ? 1754 : 1240
  const margin = 48
  const destinationCorners = [
    { x: margin, y: margin },
    { x: outputWidth - margin, y: margin },
    { x: outputWidth - margin, y: outputHeight - margin },
    { x: margin, y: outputHeight - margin },
  ]
  const equations: number[][] = []
  for (let index = 0; index < 4; index++) {
    const { x, y } = destinationCorners[index]
    const { x: sourceX, y: sourceY } = corners[index]
    equations.push([x, y, 1, 0, 0, 0, -sourceX * x, -sourceX * y, sourceX])
    equations.push([0, 0, 0, x, y, 1, -sourceY * x, -sourceY * y, sourceY])
  }
  for (let column = 0; column < 8; column++) {
    let pivot = column
    for (let row = column + 1; row < 8; row++) {
      if (Math.abs(equations[row][column]) > Math.abs(equations[pivot][column])) pivot = row
    }
    ;[equations[column], equations[pivot]] = [equations[pivot], equations[column]]
    const divisor = equations[column][column]
    if (Math.abs(divisor) < 1e-10) throw new Error("تعذر تصحيح ميل الصفحة. جرّب صورة أوضح.")
    for (let entry = column; entry <= 8; entry++) equations[column][entry] /= divisor
    for (let row = 0; row < 8; row++) {
      if (row === column) continue
      const factor = equations[row][column]
      for (let entry = column; entry <= 8; entry++) equations[row][entry] -= factor * equations[column][entry]
    }
  }
  const transform = equations.map(row => row[8])
  const output = document.createElement("canvas")
  output.width = outputWidth
  output.height = outputHeight
  const context = output.getContext("2d")
  if (!context) throw new Error("تعذر تجهيز الصفحة البيضاء.")
  const sourcePixels = sourceContext.getImageData(0, 0, source.width, source.height).data
  const cleaned = context.createImageData(outputWidth, outputHeight)
  cleaned.data.fill(255)
  for (let y = margin; y < outputHeight - margin; y++) {
    for (let x = margin; x < outputWidth - margin; x++) {
      const denominator = transform[6] * x + transform[7] * y + 1
      const sourceX = (transform[0] * x + transform[1] * y + transform[2]) / denominator
      const sourceY = (transform[3] * x + transform[4] * y + transform[5]) / denominator
      if (sourceX < 0 || sourceY < 0 || sourceX >= source.width - 1 || sourceY >= source.height - 1) continue
      const x0 = Math.floor(sourceX)
      const y0 = Math.floor(sourceY)
      const dx = sourceX - x0
      const dy = sourceY - y0
      const offset00 = (y0 * source.width + x0) * 4
      const offset10 = offset00 + 4
      const offset01 = offset00 + source.width * 4
      const offset11 = offset01 + 4
      const red = (sourcePixels[offset00] * (1 - dx) + sourcePixels[offset10] * dx) * (1 - dy)
        + (sourcePixels[offset01] * (1 - dx) + sourcePixels[offset11] * dx) * dy
      const green = (sourcePixels[offset00 + 1] * (1 - dx) + sourcePixels[offset10 + 1] * dx) * (1 - dy)
        + (sourcePixels[offset01 + 1] * (1 - dx) + sourcePixels[offset11 + 1] * dx) * dy
      const blue = (sourcePixels[offset00 + 2] * (1 - dx) + sourcePixels[offset10 + 2] * dx) * (1 - dy)
        + (sourcePixels[offset01 + 2] * (1 - dx) + sourcePixels[offset11 + 2] * dx) * dy
      const luminance = Math.round(red * 0.299 + green * 0.587 + blue * 0.114)
      const enhanced = luminance > 205 ? 255 : Math.max(0, Math.round((luminance - 20) * 1.3))
      const offset = (y * outputWidth + x) * 4
      cleaned.data[offset] = enhanced
      cleaned.data[offset + 1] = enhanced
      cleaned.data[offset + 2] = enhanced
      cleaned.data[offset + 3] = 255
    }
  }
  context.putImageData(cleaned, 0, 0)
  return new Promise((resolve, reject) => output.toBlob(blob => blob ? resolve(blob) : reject(new Error("تعذر إنشاء صورة الصفحة.")), "image/png"))
}

async function prepareVisionImage(file: File): Promise<{ imageData: string; scanned: boolean }> {
  let source: Blob = file
  let scanned = false
  try {
    source = await cleanPageImage(file)
    scanned = true
  } catch (cause) {
    console.info(`تعذر تصحيح حدود "${file.name}" قبل OCR السحابي؛ ستُرسل الصورة الأصلية بعد ضغطها.`, cause)
  }

  const image = await createImageBitmap(source)
  const scale = Math.min(1, 1400 / Math.max(image.width, image.height))
  const canvas = document.createElement("canvas")
  canvas.width = Math.max(1, Math.round(image.width * scale))
  canvas.height = Math.max(1, Math.round(image.height * scale))
  const context = canvas.getContext("2d")
  if (!context) {
    image.close()
    throw new Error("تعذر ضغط الصورة لإرسالها بأمان.")
  }
  context.fillStyle = "#fff"
  context.fillRect(0, 0, canvas.width, canvas.height)
  context.drawImage(image, 0, 0, canvas.width, canvas.height)
  image.close()

  let compressed = await new Promise<Blob | null>(resolve => canvas.toBlob(resolve, "image/jpeg", 0.78))
  if (!compressed) throw new Error("تعذر ضغط الصورة لإرسالها.")
  if (compressed.size > 4 * 1024 * 1024) {
    compressed = await new Promise<Blob | null>(resolve => canvas.toBlob(resolve, "image/jpeg", 0.58))
  }
  if (!compressed || compressed.size > 4 * 1024 * 1024) {
    throw new Error("تعذر ضغط الصورة إلى أقل من 4 ميغابايت لإرسالها.")
  }

  const dataUrl = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => typeof reader.result === "string" ? resolve(reader.result) : reject(new Error("تعذر تجهيز بيانات الصورة."))
    reader.onerror = () => reject(new Error("تعذر قراءة الصورة لإرسالها."))
    reader.readAsDataURL(compressed)
  })
  return { imageData: dataUrl.slice(dataUrl.indexOf(",") + 1), scanned }
}

export function DocumentStudio() {
  const [mode, setMode] = useState<InputMode>("text")
  const [text, setText] = useState("")
  const [files, setFiles] = useState<File[]>([])
  const [visionConsents, setVisionConsents] = useState<Set<number>>(() => new Set())
  const [recognizedImageTexts, setRecognizedImageTexts] = useState<string[]>([])
  const [cleanedScanUrl, setCleanedScanUrl] = useState("")
  const [documentFile, setDocumentFile] = useState<File | null>(null)
  const [documentText, setDocumentText] = useState("")
  const [imageUrls, setImageUrls] = useState<string[]>([])
  const [result, setResult] = useState<ExtractedDocument | null>(null)
  const [recognizedText, setRecognizedText] = useState("")
  const [aiOutput, setAiOutput] = useState("")
  const [aiQuestion, setAiQuestion] = useState("")
  const [translationDirection, setTranslationDirection] = useState("auto")
  const [allowAiUpload, setAllowAiUpload] = useState(false)
  const [busy, setBusy] = useState(false)
  const [progress, setProgress] = useState(0)
  const [status, setStatus] = useState("")
  const [error, setError] = useState("")
  const [notice, setNotice] = useState("")
  const [selectedFormat, setSelectedFormat] = useState<ExportFormat | null>(null)
  const [keepMessageTimestamps, setKeepMessageTimestamps] = useState(false)
  const [templateFile, setTemplateFile] = useState<File | null>(null)
  const [templateSheets, setTemplateSheets] = useState<string[]>([])
  const [templateSheetName, setTemplateSheetName] = useState("")
  const inputRef = useRef<HTMLInputElement>(null)
  const documentInputRef = useRef<HTMLInputElement>(null)
  const templateInputRef = useRef<HTMLInputElement>(null)
  const imageUrlsRef = useRef<string[]>([])

  const sourceText = mode === "image" ? recognizedText : mode === "document" ? documentText : text
  const currentResult = result
  const exportText = aiOutput || (currentResult?.contacts.length
    ? contactsAsText(currentResult, keepMessageTimestamps)
    : !keepMessageTimestamps ? removeMessageTimestamps(sourceText) : sourceText)
  const compatibleFormats = useMemo(
    () => currentResult
      ? formatOptions.filter(option => aiOutput.trim() ? option.id !== "excel" : formatAllowed(option.id, currentResult, sourceText))
      : [],
    [aiOutput, currentResult, sourceText],
  )

  useEffect(() => () => {
    imageUrlsRef.current.forEach(url => URL.revokeObjectURL(url))
    if (cleanedScanUrl) URL.revokeObjectURL(cleanedScanUrl)
  }, [cleanedScanUrl])

  const selectFiles = (fileList: FileList | null) => {
    if (!fileList) return
    const selected = Array.from(fileList)
    const invalid = selected.find(file => !file.type.startsWith("image/"))
    if (invalid) {
      setError(`الملف "${invalid.name}" ليس صورة مدعومة.`)
      return
    }
    const oversized = selected.find(file => file.size > 15 * 1024 * 1024)
    if (oversized) {
      setError(`حجم الصورة "${oversized.name}" أكبر من الحد المسموح (15 ميغابايت).`)
      return
    }
    setError("")
    setNotice("")
    imageUrlsRef.current.forEach(url => URL.revokeObjectURL(url))
    const nextUrls = selected.map(file => URL.createObjectURL(file))
    imageUrlsRef.current = nextUrls
    setFiles(selected)
    setVisionConsents(new Set())
    setRecognizedImageTexts([])
    setRecognizedText("")
    setResult(null)
    setAiOutput("")
    setCleanedScanUrl("")
    setImageUrls(nextUrls)
  }

  const selectDocument = (file: File | undefined) => {
    if (!file) return
    const extension = file.name.toLowerCase().split(".").pop()
    if (!["pdf", "docx", "docm", "dotx", "doc", "odt", "pptx", "pptm", "ppsx", "ppt", "xlsx", "xlsm", "xlsb", "xls", "ods", "csv", "txt", "md", "markdown", "rtf", "html", "htm"].includes(extension ?? "")) {
      setError("اختر ملفاً مدعوماً مثل PDF أو DOCX أو PPTX أو Excel أو CSV أو TXT.")
      return
    }
    if (file.size > 40 * 1024 * 1024) {
      setError("حجم الملف أكبر من الحد المسموح (40 ميغابايت).")
      return
    }
    setDocumentFile(file)
    setDocumentText("")
    setResult(null)
    setAiOutput("")
    setSelectedFormat(null)
    setCleanedScanUrl("")
    setError("")
    setNotice("")
  }

  const selectTemplate = async (file: File | undefined) => {
    if (!file) return
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      setError("ارفع ملف Excel بصيغة XLSX أو XLS.")
      return
    }
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", bookSheets: true })
      if (!workbook.SheetNames.length) throw new Error("لا يحتوي ملف Excel على أوراق عمل.")
      setTemplateFile(file)
      setTemplateSheets(workbook.SheetNames)
      setTemplateSheetName(workbook.SheetNames[0])
      setError("")
      setNotice("")
    } catch (cause) {
      console.error("تعذر فتح ملف Excel المرفق:", cause)
      setError("تعذر فتح ملف Excel. تحقق من سلامته ثم حاول رفعه مرة أخرى.")
    }
  }

  const analyze = async () => {
    setError("")
    setNotice("")
    setAiOutput("")
    setResult(null)
    setSelectedFormat(null)
    if (mode === "text") {
      const trimmed = text.trim()
      if (!trimmed) {
        setError("ألصق النص الذي تريد تحليله أولاً.")
        return
      }
      setRecognizedText(trimmed)
      setResult(extractDocument(trimmed))
      return
    }
    if (mode === "document") {
      if (!documentFile) {
        setError("ارفع مستنداً مدعوماً أولاً.")
        return
      }
      setBusy(true)
      setProgress(0)
      setStatus("قراءة الملف داخل المتصفح...")
      try {
        const extracted = await readUploadedDocument(documentFile, (message, percentage) => {
          setStatus(message)
          setProgress(percentage)
        })
        setDocumentText(extracted)
        setResult(extractDocument(extracted))
      } catch (cause) {
        console.error("تعذرت قراءة الملف:", cause)
        setError(cause instanceof Error ? cause.message : "تعذرت قراءة الملف. تحقق من سلامته وحاول مجدداً.")
      } finally {
        setBusy(false)
        setStatus("")
      }
      return
    }
    if (!files.length) {
      setError("ارفع صورة واحدة على الأقل لتحليلها.")
      return
    }

    setBusy(true)
    setProgress(0)
    setStatus("تهيئة قارئ الصور...")
    let worker: Awaited<ReturnType<typeof createWorker>> | undefined
    try {
      worker = await createWorker("ara+eng", 1, {
        logger: message => {
          if (message.status) setStatus(message.status === "recognizing text" ? "قراءة النص من الصورة..." : "تحميل محرك قراءة الصور...")
          if (typeof message.progress === "number") setProgress(Math.round(message.progress * 100))
        },
      })
      await worker.setParameters({ tessedit_pageseg_mode: PSM.SINGLE_BLOCK, preserve_interword_spaces: "1" })
      const parts: string[] = []
      const imageTexts = files.map(() => "")
      const failedFiles: string[] = []
      const unscannedFiles: string[] = []
      const uncertainFiles: Array<{ name: string; confidence: number }> = []
      for (let index = 0; index < files.length; index++) {
        setStatus(`تحليل الصورة ${index + 1} من ${files.length}...`)
        try {
          let imageForOcr: File | Blob = files[index]
          try {
            imageForOcr = await cleanPageImage(files[index])
          } catch (cause) {
            console.info(`لم تُكتشف حدود ورقة واضحة في "${files[index].name}"؛ ستتم قراءة الصورة كما هي.`, cause)
            unscannedFiles.push(files[index].name)
          }
          const { data } = await worker.recognize(imageForOcr)
          if (data.confidence < 55) uncertainFiles.push({ name: files[index].name, confidence: Math.round(data.confidence) })
          if (data.text.trim()) {
            const lines = data.blocks?.flatMap(block => block.paragraphs.flatMap(paragraph => paragraph.lines)) ?? []
            imageTexts[index] = formatOcrText(lines, data.text.trim())
            parts.push(imageTexts[index])
          }
        } catch (cause) {
          console.error(`تعذر تحليل الصورة "${files[index].name}":`, cause)
          failedFiles.push(files[index].name)
        }
        setProgress(Math.round(((index + 1) / files.length) * 100))
      }
      const recognized = parts.join("\n\n")
      if (!recognized) {
        setError(failedFiles.length === files.length
          ? `تعذرت قراءة الصور: ${failedFiles.join("، ")}. تحقق من صيغتها أو جرّب صوراً أوضح.`
          : "لم يتم العثور على نص واضح في الصور. جرّب صورة أوضح أو الصق النص مباشرة.")
        return
      }
      setRecognizedText(recognized)
      setRecognizedImageTexts(imageTexts)
      setResult(extractDocument(recognized))
      if (failedFiles.length || unscannedFiles.length || uncertainFiles.length) {
        const issues = [
          ...(failedFiles.length ? [`تعذرت قراءة ${failedFiles.join("، ")}`] : []),
          ...(unscannedFiles.length ? [`لم تُكتشف حدود ورقة A4 في ${unscannedFiles.join("، ")}`] : []),
          ...(uncertainFiles.length ? [`دقة التعرف منخفضة (${uncertainFiles.map(file => `${file.name}: ${file.confidence}%`).join("، ")})؛ راجع النص يدوياً، خصوصاً الكتابة اليدوية`] : []),
        ]
        setNotice(`تمت قراءة المحتوى المتاح؛ ${issues.join("؛ ")}.`)
      }
    } catch (cause) {
      console.error("تعذر تحليل الصورة:", cause)
      setError(cause instanceof Error
        ? `تعذر تهيئة قارئ الصور: ${cause.message}`
        : "تعذر قراءة الصور. تحقق من اتصال الإنترنت لتحميل محرك OCR، أو الصق النص مباشرة.")
    } finally {
      await worker?.terminate()
      setBusy(false)
      setStatus("")
    }
  }

  const runVisionOcr = async () => {
    const selectedIndices = [...visionConsents].filter(index => index >= 0 && index < files.length)
    if (!selectedIndices.length) {
      setError("وافق على إرسال صورة واحدة على الأقل قبل تشغيل OCR السحابي.")
      return
    }
    setBusy(true)
    setError("")
    setNotice("")
    setProgress(0)
    const nextImageTexts = [...recognizedImageTexts]
    while (nextImageTexts.length < files.length) nextImageTexts.push("")
    const failures: string[] = []
    const scanFallbacks: string[] = []
    let completed = 0
    try {
      for (const [position, index] of selectedIndices.entries()) {
        const file = files[index]
        setStatus(`قراءة الصورة بموافقتك عبر AI ${position + 1} من ${selectedIndices.length}...`)
        try {
          const prepared = await prepareVisionImage(file)
          if (!prepared.scanned) scanFallbacks.push(file.name)
          let response: Response
          try {
            response = await fetch("/api/documents/ocr", {
              method: "POST",
              headers: { "content-type": "application/json" },
              body: JSON.stringify({ imageData: prepared.imageData, mimeType: "image/jpeg" }),
            })
          } catch (cause) {
            throw new Error("تعذر الاتصال بخدمة OCR السحابية. تأكد من تشغيل خادم API.", { cause })
          }
          const responseText = await response.text()
          let payload: { result?: string; error?: string }
          try {
            payload = JSON.parse(responseText) as { result?: string; error?: string }
          } catch {
            throw new Error(response.ok ? "أعاد خادم OCR استجابة غير مفهومة." : "تعذر الاتصال بخادم OCR السحابي.")
          }
          if (!response.ok) throw new Error(payload.error || `تعذر قراءة الصورة (HTTP ${response.status}).`)
          if (!payload.result?.trim()) throw new Error("لم يتمكن مزود AI من استخراج نص واضح من الصورة.")
          nextImageTexts[index] = payload.result.trim()
          completed++
        } catch (cause) {
          console.error(`تعذر تنفيذ OCR السحابي للصورة "${file.name}":`, cause)
          failures.push(`${file.name}: ${cause instanceof Error ? cause.message : "تعذر قراءتها"}`)
        }
        setProgress(Math.round(((position + 1) / selectedIndices.length) * 100))
      }
      const combined = nextImageTexts.filter(Boolean).join("\n\n")
      if (completed) {
        setRecognizedImageTexts(nextImageTexts)
        setRecognizedText(combined)
        setResult(extractDocument(combined))
        setAiOutput("")
        setSelectedFormat(null)
      }
      if (!completed) {
        setError(failures.join("؛ ") || "تعذر تنفيذ OCR السحابي للصور المحددة.")
      } else {
        const messages = [
          `تم تحسين قراءة ${completed} صورة باستخدام OCR بصري.`,
          ...(scanFallbacks.length ? [`لم تُكتشف حدود الورقة في ${scanFallbacks.join("، ")}؛ أُرسلت الصورة كما هي بعد ضغطها.`] : []),
          ...(failures.length ? [`تعذرت قراءة بعض الصور: ${failures.join("؛ ")}`] : []),
        ]
        setNotice(messages.join(" "))
      }
    } finally {
      setVisionConsents(previous => new Set([...previous].filter(index => !selectedIndices.includes(index))))
      setBusy(false)
      setStatus("")
    }
  }

  const runAiAction = async (action: AiAction) => {
    const source = sourceText.trim()
    if (!source) {
      setError("أدخل النص أو اقرأ الملف أولاً قبل استخدام أدوات AI.")
      return
    }
    if (!allowAiUpload) {
      setError("وافق على إرسال النص المستخرج إلى مزود AI قبل المتابعة.")
      return
    }
    if (source.length > 50000) {
      setError("النص طويل جداً لطلب واحد. قسّمه إلى أجزاء لا تتجاوز 50,000 حرف.")
      return
    }
    if (action === "ask" && !aiQuestion.trim()) {
      setError("اكتب سؤالك عن المستند أولاً.")
      return
    }
    setBusy(true)
    setError("")
    setNotice("")
    setStatus("يرسل النص إلى مزود AI لمعالجته...")
    try {
      let response: Response
      try {
        response = await fetch("/api/documents/ai", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({
            action,
            text: source,
            language: translationDirection,
            question: aiQuestion.trim(),
          }),
        })
      } catch (cause) {
        throw new Error("تعذر الاتصال بخادم AI. تأكد من تشغيل خادم API واتصاله بالشبكة ثم أعد المحاولة.", { cause })
      }
      const responseText = await response.text()
      let payload: { result?: string; error?: string } = {}
      try {
        payload = JSON.parse(responseText) as { result?: string; error?: string }
      } catch {
        if (!response.ok) throw new Error("تعذر الاتصال بخادم AI. تحقق من تشغيل خادم API وربط المسار /api.")
        throw new Error("أعاد خادم AI استجابة غير مفهومة.")
      }
      if (!response.ok) throw new Error(payload.error || `تعذر تنفيذ الطلب (HTTP ${response.status}).`)
      if (!payload.result?.trim()) throw new Error("أعاد مزود AI نتيجة فارغة.")
      if (action === "translate") {
        const json = payload.result.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/, "")
        const parsed = JSON.parse(json) as { segments?: Array<{ source?: string; translation?: string }> }
        if (!Array.isArray(parsed.segments) || !parsed.segments.length || parsed.segments.some(segment =>
          !segment || typeof segment.source !== "string" || !segment.source.trim() ||
          typeof segment.translation !== "string" || !segment.translation.trim()
        )) {
          throw new Error("تعذر تنسيق الترجمة في أزواج أصل/ترجمة. أعد المحاولة.")
        }
        setAiOutput(parsed.segments.map(segment => `${segment.source}\n${segment.translation}`).join("\n\n"))
      } else {
        setAiOutput(payload.result.trim())
      }
      setSelectedFormat(null)
      setNotice("اكتملت المعالجة. راجع النتيجة ثم اختر Word أو PDF أو PowerPoint للتنزيل.")
    } catch (cause) {
      console.error("تعذر تنفيذ أداة AI:", cause)
      setError(cause instanceof Error ? cause.message : "تعذر تنفيذ طلب AI. تحقق من اتصال خادم الخدمة وإعداداته.")
    } finally {
      setBusy(false)
      setStatus("")
    }
  }

  const downloadCleanScan = async () => {
    if (!files[0]) {
      setError("ارفع صورة للصفحة أولاً.")
      return
    }
    try {
      setError("")
      const cleaned = await cleanPageImage(files[0])
      const previewUrl = URL.createObjectURL(cleaned)
      setCleanedScanUrl(previous => {
        if (previous) URL.revokeObjectURL(previous)
        return previewUrl
      })
      downloadBlob(cleaned, `${safeFileName(files[0].name.replace(/\.[^.]+$/, ""))}-مسح.png`)
      setNotice("اكتمل المسح: أزيلت الخلفية وصُححت الصفحة إلى مقاس A4 أبيض. راجع المعاينة أدناه.")
    } catch (cause) {
      console.error("تعذرت معالجة مسح الصورة:", cause)
      setError(cause instanceof Error ? cause.message : "تعذرت معالجة صورة الصفحة.")
    }
  }

  const makeWorkbook = () => {
    if (!result) throw new Error("لا توجد بيانات مستخرجة لإنشاء ملف Excel.")
    const workbook = XLSX.utils.book_new()
    if (result.rows.length) {
      const worksheet = XLSX.utils.aoa_to_sheet([result.columns, ...result.rows])
      result.columns.forEach((column, columnIndex) => {
        if (!/(amount|المبلغ|قيمة|الرصيد|السعر|total|balance)/i.test(column)) return
        result.rows.forEach((row, rowIndex) => {
          const value = numericValue(row[columnIndex] ?? "")
          if (value === null) return
          const address = XLSX.utils.encode_cell({ r: rowIndex + 1, c: columnIndex })
          worksheet[address] = { ...worksheet[address], t: "n", v: value, z: "#,##0.00" }
        })
      })
      styleWorksheetTable(worksheet, 0, 1, result.rows.length, result.columns.map((_, index) => index))
      worksheet["!cols"] = result.columns.map((column, index) => ({
        wch: Math.min(30, Math.max(column.length + 4, ...result.rows.slice(0, 100).map(row => (row[index] ?? "").length + 2))),
      }))
      XLSX.utils.book_append_sheet(workbook, worksheet, "البيانات")
    }
    if (result.facts.length) {
      const worksheet = XLSX.utils.aoa_to_sheet([["البند", "القيمة"], ...result.facts.map(fact => [fact.label, numericValue(fact.value) ?? fact.value])])
      worksheet["!cols"] = [{ wch: 32 }, { wch: 22 }]
      styleWorksheetTable(worksheet, 0, 1, result.facts.length, [0, 1])
      XLSX.utils.book_append_sheet(workbook, worksheet, FINANCIAL_WORKSHEET_NAME)
    }
    if (result.contacts.length) {
      const columns = keepMessageTimestamps ? contactDateTimeColumns : contactColumns
      const worksheet = XLSX.utils.aoa_to_sheet([
        columns,
        ...result.contacts.map(contact => [
          contact.name, contact.email, contact.phone, contact.sender,
          ...(keepMessageTimestamps ? [contact.date, contact.time] : []),
        ]),
      ])
      worksheet["!cols"] = columns.map((column, columnIndex) => ({
        wch: Math.max(column.length + 4, ...result.contacts.map(contact => [
          contact.name, contact.email, contact.phone, contact.sender, contact.date, contact.time,
        ][columnIndex].length + 2)),
      }))
      styleWorksheetTable(worksheet, 0, 1, result.contacts.length, columns.map((_, index) => index))
      XLSX.utils.book_append_sheet(workbook, worksheet, CONTACT_WORKSHEET_NAME)
    }
    return workbook
  }

  const exportWord = async (name: string) => {
    const paragraphs = exportText.split(/\n+/).filter(line => line.trim()).map(line => {
      const heading = /^#{1,3}\s|^\d+[.)]\s|^الملخص التنفيذي|^التوصيات/i.test(line.trim())
      const content = line.replace(/^#{1,3}\s*/, "").replace(/\*\*/g, "")
      const isArabic = /[\u0600-\u06ff]/.test(content)
      return new Paragraph({
        heading: heading ? HeadingLevel.HEADING_2 : undefined,
        alignment: isArabic ? AlignmentType.RIGHT : AlignmentType.LEFT,
        bidirectional: isArabic,
        spacing: { after: 140 },
        children: [new TextRun({ text: content, rightToLeft: isArabic, bold: heading })],
      })
    })
    const children: Array<Paragraph | Table> = [...paragraphs]
    const borders = {
      top: { style: BorderStyle.SINGLE, size: 1, color: "D6DFEA" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "D6DFEA" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "D6DFEA" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "D6DFEA" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "D6DFEA" },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "D6DFEA" },
    }
    if (!aiOutput.trim() && result?.rows.length) {
      const rows = [result.columns, ...result.rows].map((row, rowIndex) => new TableRow({
        children: row.map(cell => new TableCell({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            bidirectional: true,
            children: [new TextRun({ text: cell, bold: rowIndex === 0, rightToLeft: true })],
          })],
          shading: rowIndex === 0 ? { type: ShadingType.CLEAR, fill: "EDF2F7" } : undefined,
          borders,
        })),
      }))
      children.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE }, borders }))
    }
    if (!aiOutput.trim() && result?.contacts.length) {
      const columns = keepMessageTimestamps ? contactDateTimeColumns : contactColumns
      const rows = [columns, ...result.contacts.map(contact => [
        contact.name, contact.email, contact.phone, contact.sender,
        ...(keepMessageTimestamps ? [contact.date, contact.time] : []),
      ])].map((row, rowIndex) => new TableRow({
        children: row.map(cell => new TableCell({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            bidirectional: true,
            children: [new TextRun({ text: cell, bold: rowIndex === 0, rightToLeft: true })],
          })],
          shading: rowIndex === 0 ? { type: ShadingType.CLEAR, fill: "EDF2F7" } : undefined,
          borders,
        })),
      }))
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, text: "جهات الاتصال" }))
      children.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE }, borders }))
    }
    const document = new Document({
      styles: { default: { document: { run: { font: "Arial", size: 22 } } } },
      sections: [{ properties: {}, children }],
    })
    downloadBlob(await Packer.toBlob(document), `${name}.docx`)
  }

  const exportPowerPoint = async (name: string) => {
    const deck = new PptxGenJS()
    deck.layout = "LAYOUT_WIDE"
    deck.author = "المستندات الذكية"
    deck.subject = "مستند مستخرج من نص أو صورة"
    deck.theme = {
      headFontFace: "Arial",
      bodyFontFace: "Arial",
    }
    const paragraphs = exportText.split(/\n+/).map(line => line.trim()).filter(Boolean)
    const title = paragraphs[0]?.replace(/^#{1,3}\s*/, "").replace(/\*\*/g, "") ?? "ملخص المستند"
    const titleIsArabic = /[\u0600-\u06ff]/.test(title)
    const addTitleSlide = () => {
      const slide = deck.addSlide()
      slide.background = { color: "F4F7FB" }
      slide.addText(title.slice(0, 130), {
        x: 0.8, y: 1.55, w: 11.7, h: 1.4, fontFace: "Arial", fontSize: 28,
        bold: true, color: "183153", align: titleIsArabic ? "right" : "left", rtlMode: titleIsArabic, breakLine: false,
      })
      slide.addText(`تم إنشاؤه بواسطة المستندات الذكية · ${new Date().toLocaleDateString("ar")}`, {
        x: 0.8, y: 3.25, w: 11.7, h: 0.5, fontFace: "Arial", fontSize: 13,
        color: "60738A", align: "right", rtlMode: true,
      })
    }
    addTitleSlide()

    if (!aiOutput.trim() && result?.facts.length) {
      const slide = deck.addSlide()
      slide.addText("الأرقام الرئيسية", { x: 0.8, y: 0.45, w: 11.7, h: 0.55, fontSize: 24, bold: true, color: "183153", align: "right", rtlMode: true })
      slide.addText(result.facts.map(fact => `${fact.label}: ${fact.value}`).join("\n"), {
        x: 1, y: 1.3, w: 11.2, h: 5.2, fontSize: 20, color: "283B51",
        breakLine: false, valign: "middle", align: "right", rtlMode: true,
      })
    }

    const body = paragraphs.slice(1)
    for (let index = 0; index < body.length; index += 5) {
      const slide = deck.addSlide()
      const bodyChunk = body.slice(index, index + 5)
      const bodyIsArabic = /[\u0600-\u06ff]/.test(bodyChunk.join(" "))
      slide.addText("ملخص التقرير", { x: 0.8, y: 0.45, w: 11.7, h: 0.55, fontSize: 24, bold: true, color: "183153", align: "right", rtlMode: true })
      slide.addText(bodyChunk.join("\n"), {
        x: 1, y: 1.25, w: 11.2, h: 5.4, fontSize: 17, color: "283B51",
        breakLine: false, valign: "top", align: bodyIsArabic ? "right" : "left", rtlMode: bodyIsArabic,
      })
    }
    if (!aiOutput.trim() && result?.rows.length) {
      for (let index = 0; index < result.rows.length; index += 12) {
        const slide = deck.addSlide()
        slide.addText("جدول البيانات", { x: 0.8, y: 0.35, w: 11.7, h: 0.5, fontSize: 22, bold: true, color: "183153", align: "right", rtlMode: true })
        slide.addTable([result.columns, ...result.rows.slice(index, index + 12)].map(row => row.map(text => ({ text }))), {
          x: 0.45, y: 1, w: 12.4, h: 5.7, fontFace: "Arial", fontSize: 9,
          color: "26384C", border: { type: "solid", color: "D6DFEA", pt: 0.5 },
          fill: { color: "FFFFFF" }, rowH: 0.38,
        })
      }
    }
    if (!aiOutput.trim() && result?.contacts.length) {
      const columns = keepMessageTimestamps ? contactDateTimeColumns : contactColumns
      const slide = deck.addSlide()
      slide.addText("جهات الاتصال", { x: 0.8, y: 0.35, w: 11.7, h: 0.5, fontSize: 22, bold: true, color: "183153", align: "right", rtlMode: true })
      slide.addTable([
        columns,
        ...result.contacts.map(contact => [
          contact.name, contact.email, contact.phone, contact.sender,
          ...(keepMessageTimestamps ? [contact.date, contact.time] : []),
        ]),
      ].map(row => row.map(text => ({ text }))), {
        x: 0.45, y: 1, w: 12.4, h: 5.7, fontFace: "Arial", fontSize: 12,
        color: "26384C", border: { type: "solid", color: "D6DFEA", pt: 0.5 },
        fill: { color: "FFFFFF" }, rowH: 0.45,
      })
    }
    const blob = await deck.write({ outputType: "blob" })
    if (!(blob instanceof Blob)) throw new Error("تعذر إنشاء ملف PowerPoint.")
    downloadBlob(blob, `${name}.pptx`)
  }

  const exportPdf = () => {
    const frame = document.createElement("iframe")
    frame.setAttribute("title", "معاينة الطباعة إلى PDF")
    frame.style.position = "fixed"
    frame.style.right = "0"
    frame.style.bottom = "0"
    frame.style.width = "0"
    frame.style.height = "0"
    frame.style.border = "0"
    document.body.appendChild(frame)
    const printDocument = frame.contentDocument
    if (!printDocument) throw new Error("تعذر تجهيز معاينة PDF في هذا المتصفح.")
    const paragraphs = exportText.split(/\n+/).filter(line => line.trim()).map(line => {
      const heading = /^#{1,3}\s|^\d+[.)]\s/.test(line.trim())
      return `<p dir="auto" class="${heading ? "heading" : ""}">${htmlEscape(line.replace(/^#{1,3}\s*/, "").replace(/\*\*/g, ""))}</p>`
    }).join("")
    const table = !aiOutput.trim() && result?.rows.length
      ? `<table><thead><tr>${result.columns.map(cell => `<th>${htmlEscape(cell)}</th>`).join("")}</tr></thead><tbody>${result.rows.map(row => `<tr>${row.map(cell => `<td>${htmlEscape(cell)}</td>`).join("")}</tr>`).join("")}</tbody></table>`
      : ""
    const contacts = !aiOutput.trim() && result?.contacts.length
      ? `<h2>جهات الاتصال</h2><table><thead><tr>${(keepMessageTimestamps ? contactDateTimeColumns : contactColumns).map(cell => `<th>${htmlEscape(cell)}</th>`).join("")}</tr></thead><tbody>${result.contacts.map(contact => `<tr>${[contact.name, contact.email, contact.phone, contact.sender, ...(keepMessageTimestamps ? [contact.date, contact.time] : [])].map(cell => `<td>${htmlEscape(cell)}</td>`).join("")}</tr>`).join("")}</tbody></table>`
      : ""
    printDocument.open()
    printDocument.write(`<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>مستند مستخرج</title><style>
      body{font-family:Arial,sans-serif;color:#1c2b3b;margin:24mm 18mm;line-height:1.8;direction:rtl}
      h1{font-size:22px;color:#183153;border-bottom:2px solid #d8e0e8;padding-bottom:10px}
      p{white-space:pre-wrap;margin:8px 0;font-size:12pt}.heading{font-weight:bold;font-size:15pt;margin-top:18px}
      table{width:100%;border-collapse:collapse;margin-top:20px;font-size:9pt;direction:rtl}
      th,td{border:1px solid #c9d2dc;padding:6px;text-align:right}th{background:#edf2f7}
      @page{size:A4;margin:12mm}
    </style></head><body><h1>${htmlEscape(exportText.split(/\n/)[0]?.slice(0, 100) || "مستند مستخرج")}</h1>${paragraphs}${table}${contacts}</body></html>`)
    printDocument.close()
    const cleanup = () => window.setTimeout(() => frame.remove(), 1000)
    frame.contentWindow?.addEventListener("afterprint", cleanup, { once: true })
    window.setTimeout(() => frame.contentWindow?.print(), 250)
  }

  const exportSelected = async () => {
    if (!selectedFormat || !result || (aiOutput.trim() ? selectedFormat === "excel" : !formatAllowed(selectedFormat, result, sourceText))) return
    setError("")
    setNotice("")
    const name = safeFileName(exportText.split("\n")[0] || "مستند-مستخرج")
    try {
      if (selectedFormat === "excel" && templateFile) {
        const workbook = XLSX.read(await templateFile.arrayBuffer(), { type: "array", cellStyles: true })
        const appended = appendExtractedDataToWorkbook(workbook, templateSheetName, result, keepMessageTimestamps)
        const fileName = safeFileName(templateFile.name.replace(/\.[^.]+$/, ""))
        XLSX.writeFile(workbook, `${fileName}-محدّث.xlsx`)
        setNotice(`تمت إضافة ${appended.rowCount} صف إلى ورقة «${appended.sheetName}» دون حذف البيانات السابقة. تم تنزيل نسخة محدّثة وترك الملف الأصلي كما هو.`)
      } else if (selectedFormat === "excel") XLSX.writeFile(makeWorkbook(), `${name}.xlsx`)
      else if (selectedFormat === "word") await exportWord(name)
      else if (selectedFormat === "powerpoint") await exportPowerPoint(name)
      else exportPdf()
    } catch (cause) {
      console.error("تعذر تصدير المستند:", cause)
      setError("تعذر إنشاء الملف. أعد المحاولة أو اختر صيغة أخرى.")
    }
  }

  const updateRowCell = (rowIndex: number, columnIndex: number, value: string) => {
    setResult(previous => previous ? {
      ...previous,
      rows: previous.rows.map((row, currentRow) =>
        currentRow === rowIndex ? row.map((cell, currentColumn) => currentColumn === columnIndex ? value : cell) : row
      ),
    } : previous)
  }

  const updateFact = (factIndex: number, value: string) => {
    setResult(previous => previous ? {
      ...previous,
      facts: previous.facts.map((fact, index) => index === factIndex ? { ...fact, value } : fact),
    } : previous)
  }

  const updateContact = (rowIndex: number, columnIndex: number, value: string) => {
    const keys = ["name", "email", "phone", "sender", "date", "time"] as const
    const key = keys[columnIndex]
    if (!key) return
    setResult(previous => previous ? {
      ...previous,
      contacts: previous.contacts.map((contact, index) => index === rowIndex ? { ...contact, [key]: value } : contact),
    } : previous)
  }

  const visibleContactColumns = keepMessageTimestamps ? contactDateTimeColumns : contactColumns

  return (
    <main dir="rtl" className="min-h-full px-4 py-6 sm:px-6 lg:px-10">
      <div className="mx-auto max-w-6xl space-y-6">
        <header className="flex flex-wrap items-center justify-between gap-4">
          <div>
            <div className="mb-2 inline-flex items-center gap-2 rounded-full border border-primary/15 bg-primary/5 px-3 py-1 text-xs font-semibold text-primary">
              <Sparkles className="h-3.5 w-3.5" /> تحويل ذكي للمستندات
            </div>
            <h1 className="text-2xl font-bold tracking-tight text-foreground sm:text-3xl">حوّل النصوص والصور إلى ملفات جاهزة</h1>
            <p className="mt-2 max-w-2xl text-sm leading-6 text-muted-foreground">اقرأ مستندات PDF وWord وPowerPoint وExcel والصور أو ألصق النص، ثم ترجمه أو لخّصه أو نسّقه ونزّله كملف مرتب.</p>
          </div>
          <div className="hidden items-center gap-2 rounded-xl border border-border bg-card px-4 py-3 text-xs text-muted-foreground md:flex">
            <ScanText className="h-4 w-4 text-primary" /> قراءة الملفات محلياً · AI بموافقتك فقط
          </div>
        </header>

        <section className="grid gap-5 lg:grid-cols-[minmax(0,1.2fr)_minmax(300px,.8fr)]">
          <div className="rounded-2xl border border-border bg-card p-5 shadow-sm sm:p-6">
            <div className="mb-5 flex flex-wrap items-center justify-between gap-3">
              <div>
                <h2 className="font-bold text-foreground">1. أضف المحتوى</h2>
                <p className="mt-1 text-xs text-muted-foreground">يدعم الجداول والنصوص العربية والإنجليزية</p>
              </div>
              <div className="inline-flex flex-wrap rounded-lg bg-muted p-1">
                <button type="button" onClick={() => { setMode("text"); setResult(null); setAiOutput(""); setSelectedFormat(null); setError("") }} className={`rounded-md px-3 py-2 text-sm font-semibold transition ${mode === "text" ? "bg-card text-foreground shadow-sm" : "text-muted-foreground"}`}>
                  <span className="inline-flex items-center gap-2"><ClipboardPaste className="h-4 w-4" /> نص</span>
                </button>
                <button type="button" onClick={() => { setMode("image"); setResult(null); setAiOutput(""); setSelectedFormat(null); setError("") }} className={`rounded-md px-3 py-2 text-sm font-semibold transition ${mode === "image" ? "bg-card text-foreground shadow-sm" : "text-muted-foreground"}`}>
                  <span className="inline-flex items-center gap-2"><FileImage className="h-4 w-4" /> صورة</span>
                </button>
                <button type="button" onClick={() => { setMode("document"); setResult(null); setAiOutput(""); setSelectedFormat(null); setError("") }} className={`rounded-md px-3 py-2 text-sm font-semibold transition ${mode === "document" ? "bg-card text-foreground shadow-sm" : "text-muted-foreground"}`}>
                  <span className="inline-flex items-center gap-2"><FileUp className="h-4 w-4" /> PDF / PowerPoint</span>
                </button>
              </div>
            </div>

            {mode === "text" ? (
              <textarea
                value={text}
                onChange={event => { setText(event.target.value); setResult(null); setAiOutput(""); setSelectedFormat(null); setError("") }}
                placeholder={"ألصق النص أو الجدول هنا...\n\nمثال:\nالتاريخ\tنوع الحركة\tالمبلغ\n05/10/2026\tتحويل إلى البنك\t₪ 20,684.00"}
                className="min-h-64 w-full resize-y rounded-xl border border-input bg-background px-4 py-3 text-sm leading-7 outline-none transition placeholder:text-muted-foreground/70 focus:border-primary focus:ring-2 focus:ring-primary/15"
                dir="auto"
              />
            ) : mode === "image" ? (
              <div>
                <input ref={inputRef} type="file" accept="image/*" multiple className="hidden" onChange={event => { selectFiles(event.target.files); event.currentTarget.value = "" }} />
                <button type="button" onClick={() => inputRef.current?.click()} className="flex min-h-44 w-full flex-col items-center justify-center rounded-xl border-2 border-dashed border-primary/25 bg-primary/[0.025] px-5 py-7 text-center transition hover:border-primary/50 hover:bg-primary/[0.05]">
                  <span className="mb-3 flex h-12 w-12 items-center justify-center rounded-2xl bg-primary/10 text-primary"><ImagePlus className="h-6 w-6" /></span>
                  <span className="font-semibold text-foreground">اضغط لاختيار صورة أو أكثر</span>
                  <span className="mt-1 text-xs text-muted-foreground">PNG أو JPG · حتى 15 ميغابايت للصورة</span>
                </button>
                {files.length > 0 && (
                  <div className="mt-4 grid grid-cols-2 gap-3 sm:grid-cols-3">
                    {files.map((file, index) => (
                      <div key={`${file.name}-${file.lastModified}`} className="relative overflow-hidden rounded-lg border border-border bg-muted">
                        <img src={imageUrls[index]} alt={`معاينة ${file.name}`} className="h-28 w-full object-contain" />
                        <div className="truncate border-t border-border bg-card px-2 py-1.5 text-xs text-muted-foreground">{file.name}</div>
                        <label className="flex items-start gap-2 border-t border-border bg-card p-2 text-[11px] leading-4 text-muted-foreground">
                          <input type="checkbox" checked={visionConsents.has(index)} onChange={event => setVisionConsents(previous => {
                            const next = new Set(previous)
                            if (event.target.checked) next.add(index)
                            else next.delete(index)
                            return next
                          })} className="mt-0.5 h-3.5 w-3.5 shrink-0 accent-primary" />
                          <span>أوافق على إرسال هذه الصورة لمزود AI بصري لاستخراج النص. لا تُرسل إلا عند ضغط زر OCR السحابي.</span>
                        </label>
                        <button type="button" aria-label={`إزالة ${file.name}`} onClick={() => {
                          URL.revokeObjectURL(imageUrlsRef.current[index])
                          imageUrlsRef.current = imageUrlsRef.current.filter((_, fileIndex) => fileIndex !== index)
                          setFiles(previous => previous.filter((_, fileIndex) => fileIndex !== index))
                          setImageUrls(previous => previous.filter((_, fileIndex) => fileIndex !== index))
                          setVisionConsents(previous => new Set([...previous].filter(consentedIndex => consentedIndex !== index).map(consentedIndex => consentedIndex > index ? consentedIndex - 1 : consentedIndex)))
                          setRecognizedImageTexts(previous => previous.filter((_, fileIndex) => fileIndex !== index))
                          setResult(null)
                          setAiOutput("")
                          setSelectedFormat(null)
                        }} className="absolute left-2 top-2 rounded-full bg-card/90 p-1 text-foreground shadow hover:text-destructive"><X className="h-4 w-4" /></button>
                      </div>
                    ))}
                  </div>
                )}
                {files.length > 0 && (
                  <button type="button" disabled={busy || visionConsents.size === 0} onClick={() => void runVisionOcr()} className="mt-3 inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-2 text-xs font-semibold text-primary-foreground disabled:cursor-not-allowed disabled:opacity-50">
                    <Sparkles className="h-4 w-4" /> تحسين النص بالصورة عبر AI
                  </button>
                )}
                {files.length > 0 && <button type="button" onClick={() => void downloadCleanScan()} className="mt-3 inline-flex items-center gap-2 rounded-lg border border-border bg-background px-3 py-2 text-xs font-semibold text-foreground hover:border-primary/40"><ScanLine className="h-4 w-4 text-primary" /> مسح الورقة وإخراجها A4 أبيض</button>}
                {cleanedScanUrl && (
                  <div className="mt-4 rounded-xl border border-border bg-muted/40 p-3">
                    <p className="mb-2 text-xs font-semibold text-foreground">معاينة الصفحة بعد المسح · A4</p>
                    <img src={cleanedScanUrl} alt="معاينة الورقة بعد تنظيفها وتحويلها إلى A4" className="mx-auto max-h-[560px] max-w-full rounded border bg-white object-contain shadow-sm" />
                    <button type="button" onClick={() => {
                      const source = files[0]
                      if (source) void fetch(cleanedScanUrl).then(response => response.blob()).then(blob =>
                        downloadBlob(blob, `${safeFileName(source.name.replace(/\.[^.]+$/, ""))}-مسح.png`))
                        .catch(cause => {
                          console.error("تعذر تنزيل صورة المسح:", cause)
                          setError("تعذر تنزيل صورة المسح. أعد إنشاء المسح وحاول مرة أخرى.")
                        })
                    }} className="mt-3 inline-flex items-center gap-2 rounded-lg border border-border bg-background px-3 py-2 text-xs font-semibold text-foreground hover:border-primary/40">
                      <Download className="h-4 w-4" /> تنزيل صورة A4
                    </button>
                  </div>
                )}
                <p className="mt-3 text-xs leading-5 text-muted-foreground">تُقرأ الصور داخل متصفحك باستخدام OCR. تُحمّل ملفات اللغة عند أول استخدام، ولا تُرفع صورك إلى خادم.</p>
              </div>
            ) : (
              <div>
                <input ref={documentInputRef} type="file" accept=".pdf,.docx,.docm,.dotx,.doc,.odt,.pptx,.pptm,.ppsx,.ppt,.xlsx,.xlsm,.xlsb,.xls,.ods,.csv,.txt,.md,.markdown,.rtf,.html,.htm,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/plain,text/csv,text/html" className="hidden" onChange={event => { selectDocument(event.target.files?.[0]); event.currentTarget.value = "" }} />
                <button type="button" onClick={() => documentInputRef.current?.click()} className="flex min-h-44 w-full flex-col items-center justify-center rounded-xl border-2 border-dashed border-primary/25 bg-primary/[0.025] px-5 py-7 text-center transition hover:border-primary/50 hover:bg-primary/[0.05]">
                  <span className="mb-3 flex h-12 w-12 items-center justify-center rounded-2xl bg-primary/10 text-primary"><FileUp className="h-6 w-6" /></span>
                  <span className="font-semibold text-foreground">{documentFile ? "اختيار ملف آخر" : "ارفع مستنداً أو جدولاً"}</span>
                  <span className="mt-1 text-xs text-muted-foreground">PDF · DOCX · ODT · PPTX · XLSX · XLS · ODS · CSV · TXT · MD · RTF · HTML · حتى 40 ميغابايت</span>
                  {documentFile && <span className="mt-3 max-w-full truncate rounded-md bg-muted px-3 py-1.5 text-xs text-foreground">{documentFile.name}</span>}
                </button>
                {documentFile && /\.(ppt|doc)$/i.test(documentFile.name) && <p className="mt-3 text-xs text-amber-800 dark:text-amber-200">صيغة Office القديمة تحتاج إلى حفظ الملف كـ .pptx أو .docx قبل القراءة.</p>}
                {documentText && (
                  <label className="mt-3 block">
                    <span className="mb-1 block text-xs font-semibold text-muted-foreground">النص المستخرج · يمكنك تصحيح العربية قبل التصدير</span>
                    <textarea value={documentText} onChange={event => {
                      const value = event.target.value
                      setDocumentText(value)
                      setResult(extractDocument(value))
                      setAiOutput("")
                      setSelectedFormat(null)
                    }} className="max-h-56 min-h-32 w-full resize-y overflow-auto rounded-lg border border-border bg-background p-3 text-xs leading-6 text-foreground" dir="auto" />
                  </label>
                )}
                <p className="mt-3 text-xs leading-5 text-muted-foreground">تُقرأ الملفات محلياً داخل المتصفح. الملف نفسه لا يُرسل لمزود AI؛ عند طلب الترجمة أو التلخيص يُرسل النص المستخرج فقط.</p>
              </div>
            )}

            <div className="mt-5 rounded-xl border border-border bg-background p-4">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-start gap-3">
                  <span className="flex h-10 w-10 shrink-0 items-center justify-center rounded-lg bg-emerald-500/10 text-emerald-700"><FileSpreadsheet className="h-5 w-5" /></span>
                  <div>
                    <div className="text-sm font-semibold text-foreground">إلحاق البيانات بملف Excel موجود</div>
                    <p className="mt-1 text-xs leading-5 text-muted-foreground">ارفع XLSX أو XLS، وستضاف البيانات بعد آخر صف مع مطابقة الأعمدة قدر الإمكان. يُنزّل النظام نسخة محدّثة ويترك الأصل دون تغيير.</p>
                  </div>
                </div>
                <input
                  ref={templateInputRef}
                  type="file"
                  accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
                  className="hidden"
                  onChange={event => { void selectTemplate(event.target.files?.[0]); event.currentTarget.value = "" }}
                />
                <button type="button" onClick={() => templateInputRef.current?.click()} className="inline-flex shrink-0 items-center gap-2 rounded-lg border border-border bg-card px-3 py-2 text-xs font-semibold text-foreground transition hover:border-primary/40 hover:bg-muted/50">
                  <Upload className="h-4 w-4" /> {templateFile ? "تغيير الملف" : "اختيار ملف Excel"}
                </button>
              </div>
              {templateFile && (
                <div className="mt-3 flex flex-wrap items-center gap-3 border-t border-border pt-3">
                  <span className="max-w-full truncate rounded-md bg-muted px-2.5 py-1.5 text-xs text-foreground" title={templateFile.name}>{templateFile.name}</span>
                  <label className="flex items-center gap-2 text-xs text-muted-foreground">
                    ورقة العمل
                    <select value={templateSheetName} onChange={event => setTemplateSheetName(event.target.value)} className="max-w-56 rounded-md border border-input bg-card px-2 py-1.5 text-foreground">
                      {templateSheets.map(sheet => <option key={sheet} value={sheet}>{sheet}</option>)}
                    </select>
                  </label>
                  <button type="button" onClick={() => { setTemplateFile(null); setTemplateSheets([]); setTemplateSheetName(""); setNotice("") }} className="inline-flex items-center gap-1 text-xs text-destructive hover:underline">
                    <X className="h-3.5 w-3.5" /> إزالة
                  </button>
                </div>
              )}
            </div>

            {error && <div role="alert" className="mt-4 flex items-start gap-2 rounded-lg border border-destructive/25 bg-destructive/5 p-3 text-sm text-destructive"><AlertCircle className="mt-0.5 h-4 w-4 shrink-0" />{error}</div>}
            {busy && (
              <div className="mt-4 rounded-xl border border-primary/20 bg-primary/5 p-4" aria-live="polite">
                <div className="mb-2 flex items-center justify-between gap-2 text-xs"><span className="flex items-center gap-2 font-medium text-foreground"><LoaderCircle className="h-4 w-4 animate-spin text-primary" />{status}</span><span className="text-muted-foreground">{progress}%</span></div>
                <div className="h-2 overflow-hidden rounded-full bg-primary/10"><div className="h-full rounded-full bg-primary transition-all" style={{ width: `${progress}%` }} /></div>
              </div>
            )}
            <div className="mt-4 flex flex-wrap items-center justify-between gap-3">
              <span className="text-xs text-muted-foreground">{mode === "text" ? `${text.length.toLocaleString("ar")} حرف` : mode === "image" ? `${files.length.toLocaleString("ar")} صورة` : documentFile?.name ?? "لم يتم اختيار ملف"}</span>
              <button type="button" disabled={busy} onClick={() => void analyze()} className="inline-flex items-center gap-2 rounded-lg bg-primary px-5 py-2.5 text-sm font-semibold text-primary-foreground shadow-sm transition hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-60">
                {busy ? <LoaderCircle className="h-4 w-4 animate-spin" /> : <ScanText className="h-4 w-4" />}
                {busy ? "جارٍ التحليل..." : "حلّل المحتوى"}
              </button>
            </div>
          </div>

          <div className="rounded-2xl border border-border bg-card p-5 shadow-sm sm:p-6">
            <h2 className="font-bold text-foreground">2. اختر صيغة الإخراج</h2>
            <p className="mt-1 text-xs leading-5 text-muted-foreground">تظهر الصيغ المتوافقة مع المحتوى بعد التحليل.</p>
            <div className="mt-4 space-y-2">
              {formatOptions.map(option => {
                const Icon = option.icon
                const allowed = currentResult
                  ? aiOutput.trim()
                    ? option.id !== "excel"
                    : formatAllowed(option.id, currentResult, sourceText)
                  : false
                const selected = selectedFormat === option.id
                return (
                  <button key={option.id} type="button" disabled={!allowed} onClick={() => setSelectedFormat(option.id)} className={`flex w-full items-center gap-3 rounded-xl border p-3 text-right transition ${selected ? "border-primary bg-primary/5 ring-1 ring-primary/20" : allowed ? "border-border hover:border-primary/40 hover:bg-muted/50" : "cursor-not-allowed border-border/70 opacity-50"}`}>
                    <span className={`flex h-10 w-10 shrink-0 items-center justify-center rounded-lg ${selected ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground"}`}><Icon className="h-5 w-5" /></span>
                    <span className="min-w-0 flex-1"><span className="block text-sm font-semibold text-foreground">{option.label}</span><span className="mt-0.5 block text-xs text-muted-foreground">{option.description}</span></span>
                    {selected && <Check className="h-4 w-4 text-primary" />}
                  </button>
                )
              })}
            </div>
            {currentResult && !compatibleFormats.some(option => option.id === "excel") && (
              <div className="mt-3 rounded-lg bg-amber-500/10 p-3 text-xs leading-5 text-amber-900 dark:text-amber-200">تعذر تجهيز Excel: لم أجد جدولاً أو أرقاماً مالية واضحة لاستخراجها. يمكنك تنزيل النص كتقرير Word أو PDF أو PowerPoint.</div>
            )}
            <button type="button" disabled={!selectedFormat || !currentResult} onClick={() => void exportSelected()} className="mt-4 inline-flex w-full items-center justify-center gap-2 rounded-lg bg-foreground px-4 py-3 text-sm font-semibold text-background transition hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-40">
              <Download className="h-4 w-4" /> تنزيل الملف
            </button>
            {notice && <p role="status" className="mt-3 rounded-lg border border-emerald-500/25 bg-emerald-500/10 p-3 text-xs leading-5 text-emerald-800 dark:text-emerald-200">{notice}</p>}
            {selectedFormat === "pdf" && <p className="mt-2 text-center text-[11px] text-muted-foreground">سيُفتح مربع الطباعة؛ اختر «حفظ بتنسيق PDF».</p>}
          </div>
        </section>

        <section className="space-y-4 rounded-2xl border border-primary/20 bg-card p-5 shadow-sm sm:p-6">
          <div className="flex items-center gap-2">
            <WandSparkles className="h-5 w-5 text-primary" />
            <div>
              <h2 className="font-bold text-foreground">مساعد AI للمستند</h2>
              <p className="mt-1 text-xs leading-5 text-muted-foreground">ترجمة فقرة بفقرة، تلخيص، استخراج مصطلحات، تنسيق، أو سؤال عن المحتوى.</p>
            </div>
          </div>
          <label className="flex items-start gap-2 rounded-lg border border-amber-500/25 bg-amber-500/5 p-3 text-xs leading-5 text-foreground">
            <input type="checkbox" checked={allowAiUpload} onChange={event => setAllowAiUpload(event.target.checked)} className="mt-1 h-4 w-4 shrink-0 accent-primary" />
            <span>أوافق على إرسال <strong>النص المستخرج فقط</strong> إلى مزود AI المهيأ على الخادم لمعالجته. لن يُرسل ملف PDF أو الصورة نفسها. لا تتابع إذا كان النص يحتوي معلومات لا تريد مشاركتها.</span>
          </label>
          <div className="flex flex-wrap items-center gap-2">
            <label className="flex items-center gap-2 text-xs text-muted-foreground">
              اتجاه الترجمة
              <select value={translationDirection} onChange={event => setTranslationDirection(event.target.value)} className="rounded-md border border-input bg-background px-2 py-2 text-foreground">
                <option value="auto">اكتشاف تلقائي</option>
                <option value="ar-en">العربية إلى الإنجليزية</option>
                <option value="en-ar">الإنجليزية إلى العربية</option>
              </select>
            </label>
            <button type="button" disabled={busy || !sourceText.trim() || !allowAiUpload} onClick={() => void runAiAction("translate")} className="inline-flex items-center gap-2 rounded-lg bg-primary px-3 py-2 text-xs font-semibold text-primary-foreground disabled:cursor-not-allowed disabled:opacity-50"><Languages className="h-4 w-4" /> ترجمة مع الأصل تحته</button>
            <button type="button" disabled={busy || !sourceText.trim() || !allowAiUpload} onClick={() => void runAiAction("summarize")} className="inline-flex items-center gap-2 rounded-lg border border-border bg-background px-3 py-2 text-xs font-semibold text-foreground disabled:cursor-not-allowed disabled:opacity-50"><BookOpen className="h-4 w-4" /> تلخيص</button>
            <button type="button" disabled={busy || !sourceText.trim() || !allowAiUpload} onClick={() => void runAiAction("glossary")} className="inline-flex items-center gap-2 rounded-lg border border-border bg-background px-3 py-2 text-xs font-semibold text-foreground disabled:cursor-not-allowed disabled:opacity-50"><FileText className="h-4 w-4" /> مصطلحات وترجمتها</button>
            <button type="button" disabled={busy || !sourceText.trim() || !allowAiUpload} onClick={() => void runAiAction("format")} className="inline-flex items-center gap-2 rounded-lg border border-border bg-background px-3 py-2 text-xs font-semibold text-foreground disabled:cursor-not-allowed disabled:opacity-50"><WandSparkles className="h-4 w-4" /> تنسيق النص</button>
          </div>
          <div className="flex flex-col gap-2 sm:flex-row">
            <input value={aiQuestion} onChange={event => setAiQuestion(event.target.value)} maxLength={1000} placeholder="اسأل عن محتوى الملف أو اطلب إجراءً محدداً..." className="min-w-0 flex-1 rounded-lg border border-input bg-background px-3 py-2 text-sm outline-none focus:border-primary" dir="auto" />
            <button type="button" disabled={busy || !sourceText.trim() || !allowAiUpload || !aiQuestion.trim()} onClick={() => void runAiAction("ask")} className="inline-flex items-center justify-center gap-2 rounded-lg border border-border bg-background px-4 py-2 text-xs font-semibold text-foreground disabled:cursor-not-allowed disabled:opacity-50"><Sparkles className="h-4 w-4" /> اسأل المساعد</button>
          </div>
          {busy && status.includes("مزود AI") && <p role="status" className="text-xs text-primary">{status}</p>}
          {aiOutput && (
            <div className="rounded-xl border border-border bg-white p-4 text-slate-900 shadow-inner">
              <div className="mb-2 text-xs font-semibold text-slate-500">معاينة النتيجة · يمكنك تعديل النص قبل التصدير</div>
              <textarea value={aiOutput} onChange={event => setAiOutput(event.target.value)} className="min-h-48 w-full resize-y bg-transparent text-sm leading-7 outline-none" dir="auto" aria-label="نتيجة معالجة AI" />
            </div>
          )}
        </section>

        {currentResult && (
          <section className="space-y-4 rounded-2xl border border-border bg-card p-5 shadow-sm sm:p-6">
            <div className="flex flex-wrap items-start justify-between gap-3">
              <div>
                <div className="flex items-center gap-2"><CheckCircle2 className="h-5 w-5 text-emerald-600" /><h2 className="font-bold text-foreground">3. راجع البيانات المستخرجة</h2></div>
                <p className="mt-1 text-xs text-muted-foreground">راجع النتائج قبل تنزيلها، خصوصاً الأرقام التي قُرئت من صورة.</p>
              </div>
              <span className="rounded-full bg-muted px-3 py-1 text-xs text-muted-foreground">{currentResult.rows.length ? `${currentResult.rows.length} صف` : currentResult.contacts.length ? `${currentResult.contacts.length} جهة اتصال` : `${currentResult.facts.length} قيمة مالية`}</span>
            </div>
            {mode === "image" && <div className="rounded-lg border border-amber-500/25 bg-amber-500/10 px-3 py-2 text-xs leading-5 text-amber-900 dark:text-amber-200">قد يخطئ OCR في قراءة الأرقام أو النصوص. اضغط على أي خلية لتصحيحها قبل تنزيل الملف.</div>}
            {currentResult.contacts.length > 0 && (
              <>
                <label className="flex w-fit cursor-pointer items-center gap-2 rounded-lg border border-border bg-background px-3 py-2 text-xs text-foreground">
                  <input type="checkbox" checked={keepMessageTimestamps} onChange={event => setKeepMessageTimestamps(event.target.checked)} className="h-4 w-4 accent-primary" />
                  <CalendarClock className="h-4 w-4 text-primary" />
                  تضمين تاريخ ووقت الرسائل في الملف (اختياري)
                </label>
                <div className="overflow-x-auto rounded-xl border border-border">
                  <table className="w-full min-w-max border-collapse text-right text-xs">
                    <thead className="bg-muted/70"><tr>{visibleContactColumns.map(column => <th key={column} className="whitespace-nowrap border-b border-border px-3 py-2.5 font-semibold text-foreground">{column}</th>)}</tr></thead>
                    <tbody>{currentResult.contacts.map((contact, rowIndex) => {
                      const values = [contact.name, contact.email, contact.phone, contact.sender, contact.date, contact.time]
                      return <tr key={rowIndex} className="odd:bg-card even:bg-muted/30">{values.slice(0, visibleContactColumns.length).map((value, columnIndex) => <td key={columnIndex} className="border-b border-border/70 px-2 py-1"><input aria-label={`${visibleContactColumns[columnIndex]}، جهة الاتصال ${rowIndex + 1}`} value={value} onChange={event => updateContact(rowIndex, columnIndex, event.target.value)} className="w-full min-w-32 rounded px-1 py-1 text-xs text-foreground outline-none focus:bg-background focus:ring-1 focus:ring-primary/40" dir="auto" /></td>)}</tr>
                    })}</tbody>
                  </table>
                  <div className="border-t border-border bg-muted/40 px-3 py-2 text-center text-xs text-muted-foreground">{keepMessageTimestamps ? "سيُضمّن تاريخ ووقت كل رسالة في الأعمدة المخصصة. يمكنك تعديلهما قبل التصدير." : "لن يُضمّن تاريخ أو وقت الرسائل. فعّل الخيار أعلاه لإضافتهما."} راجع بيانات الاتصال قبل التصدير.</div>
                </div>
              </>
            )}
            {currentResult.rows.length > 0 && (
              <div className="overflow-x-auto rounded-xl border border-border">
                <table className="w-full min-w-max border-collapse text-right text-xs">
                  <thead className="bg-muted/70"><tr>{currentResult.columns.map((column, index) => <th key={`${column}-${index}`} className="whitespace-nowrap border-b border-border px-3 py-2.5 font-semibold text-foreground">{column || `عمود ${index + 1}`}</th>)}</tr></thead>
                  <tbody>{currentResult.rows.slice(0, 8).map((row, rowIndex) => <tr key={rowIndex} className="odd:bg-card even:bg-muted/30">{row.map((cell, cellIndex) => <td key={cellIndex} className="border-b border-border/70 px-2 py-1"><input aria-label={`${currentResult.columns[cellIndex] || `حقل ${cellIndex + 1}`}، الصف ${rowIndex + 1}`} value={cell} onChange={event => updateRowCell(rowIndex, cellIndex, event.target.value)} className="w-full min-w-20 rounded px-1 py-1 text-xs text-foreground outline-none focus:bg-background focus:ring-1 focus:ring-primary/40" dir="auto" /></td>)}</tr>)}</tbody>
                </table>
                {currentResult.rows.length > 8 && <div className="border-t border-border bg-muted/40 px-3 py-2 text-center text-xs text-muted-foreground">يُعرض أول 8 صفوف فقط هنا؛ يتضمن ملف Excel كل الصفوف ({currentResult.rows.length}).</div>}
              </div>
            )}
            {currentResult.facts.length > 0 && (
              <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
                {currentResult.facts.map((fact, index) => <label key={`${fact.label}-${index}`} className="block rounded-xl border border-border bg-background p-4"><span className="text-xs text-muted-foreground">{fact.label}</span><input aria-label={fact.label} value={fact.value} onChange={event => updateFact(index, event.target.value)} className="mt-1 block w-full rounded bg-transparent text-xl font-bold text-foreground outline-none focus:ring-1 focus:ring-primary/40" dir="auto" /></label>)}
              </div>
            )}
            {!currentResult.rows.length && !currentResult.facts.length && !currentResult.contacts.length && <div className="rounded-xl border border-dashed border-border p-4 text-sm text-muted-foreground">تم حفظ النص كما هو للتصدير، لكن لم أتمكن من التعرف على جدول أو مبالغ أو جهات اتصال واضحة.</div>}
            {mode === "image" && recognizedText && (
              <details className="rounded-xl border border-border bg-background p-4">
                <summary className="cursor-pointer text-sm font-semibold text-foreground">النص الذي قرأته من الصورة</summary>
                <label className="mt-3 block">
                  <span className="mb-1 block text-xs text-muted-foreground">صحّح النص هنا؛ ستُحدَّث البيانات المستخرجة والملفات بناءً على تعديلاتك.</span>
                  <textarea value={recognizedText} onChange={event => {
                    const value = event.target.value
                    setRecognizedText(value)
                    setResult(extractDocument(value))
                    setAiOutput("")
                    setSelectedFormat(null)
                  }} className="min-h-40 w-full resize-y rounded-lg border border-border bg-background p-3 text-xs leading-6 text-foreground" dir="auto" aria-label="النص المستخرج من الصورة" />
                </label>
              </details>
            )}
          </section>
        )}

        {!currentResult && (
          <div className="grid gap-3 sm:grid-cols-3">
            <div className="flex items-start gap-3 rounded-xl border border-border/70 bg-card/60 p-4"><ClipboardPaste className="mt-0.5 h-4 w-4 shrink-0 text-primary" /><div><div className="text-sm font-semibold text-foreground">الصق أي نص</div><p className="mt-1 text-xs leading-5 text-muted-foreground">تقارير، قوائم مالية، أو جداول منسوخة من موقع.</p></div></div>
            <div className="flex items-start gap-3 rounded-xl border border-border/70 bg-card/60 p-4"><Upload className="mt-0.5 h-4 w-4 shrink-0 text-primary" /><div><div className="text-sm font-semibold text-foreground">ارفع صورة واضحة</div><p className="mt-1 text-xs leading-5 text-muted-foreground">استخرج الكتابة والجداول بالعربية والإنجليزية.</p></div></div>
            <div className="flex items-start gap-3 rounded-xl border border-border/70 bg-card/60 p-4"><FileSpreadsheet className="mt-0.5 h-4 w-4 shrink-0 text-primary" /><div><div className="text-sm font-semibold text-foreground">راجع ثم نزّل</div><p className="mt-1 text-xs leading-5 text-muted-foreground">Excel للبيانات، وWord وPDF وPowerPoint للتقارير.</p></div></div>
          </div>
        )}
      </div>
    </main>
  )
}
