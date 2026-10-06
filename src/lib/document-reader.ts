import JSZip from "jszip"
import { GlobalWorkerOptions, getDocument } from "pdfjs-dist"
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url"
import { createWorker } from "tesseract.js"
import * as XLSX from "xlsx"

GlobalWorkerOptions.workerSrc = pdfWorkerUrl
GlobalWorkerOptions.workerPort ??= new Worker(pdfWorkerUrl, { type: "module" })

export async function readUploadedDocument(file: File, onProgress?: (status: string, progress: number) => void): Promise<string> {
  const extension = file.name.toLowerCase().split(".").pop()
  if (extension === "pdf") {
    const pdf = await getDocument({ data: await file.arrayBuffer() }).promise
    if (pdf.numPages > 200) throw new Error("عدد صفحات PDF يتجاوز الحد المدعوم (200 صفحة).")
    const pages: string[] = []
    let ocrWorker: Awaited<ReturnType<typeof createWorker>> | undefined
    let ocrPageCount = 0
    try {
      for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
        const page = await pdf.getPage(pageNumber)
        const content = await page.getTextContent()
        const lines: Array<{ y: number; direction: string; parts: Array<{ x: number; text: string }> }> = []
        for (const item of content.items) {
          if (!("str" in item) || !item.str.trim()) continue
          const y = Math.round(item.transform[5])
          let line = lines.find(candidate => Math.abs(candidate.y - y) <= 2 && candidate.direction === item.dir)
          if (!line) {
            line = { y, direction: item.dir, parts: [] }
            lines.push(line)
          }
          line.parts.push({ x: item.transform[4], text: item.str.trim() })
        }
        let text = lines
          .sort((left, right) => right.y - left.y)
          .map(line => line.parts
            .sort((left, right) => line.direction === "rtl" ? right.x - left.x : left.x - right.x)
            .map(part => part.text)
            .join(" ")
            .replace(/\s+/g, " ")
            .trim())
          .filter(Boolean)
          .join("\n")
        if (!text.trim()) {
          ocrPageCount++
          if (ocrPageCount > 50) throw new Error("الملف ممسوح ضوئياً ويتجاوز حد OCR (50 صفحة). قسّم الملف أو ارفع صفحات أقل.")
          onProgress?.(`استخراج النص من الصفحة المصورة ${pageNumber}...`, Math.round((pageNumber / pdf.numPages) * 100))
          ocrWorker ??= await createWorker("ara+eng", 1, {
            logger: message => {
              if (message.status === "recognizing text" && typeof message.progress === "number") {
                onProgress?.(`التعرف على النص في الصفحة ${pageNumber}...`, Math.round(((pageNumber - 1 + message.progress) / pdf.numPages) * 100))
              }
            },
          })
          const viewport = page.getViewport({ scale: 1.6 })
          const canvas = document.createElement("canvas")
          canvas.width = Math.ceil(viewport.width)
          canvas.height = Math.ceil(viewport.height)
          const context = canvas.getContext("2d")
          if (!context) throw new Error("تعذر تجهيز صفحة PDF المصورة لـ OCR.")
          await page.render({ canvasContext: context, viewport }).promise
          const recognized = await ocrWorker.recognize(canvas)
          text = recognized.data.text.trim()
        }
        if (text.trim()) pages.push(`صفحة ${pageNumber}\n${text}`)
        onProgress?.(`تمت قراءة الصفحة ${pageNumber} من ${pdf.numPages}`, Math.round((pageNumber / pdf.numPages) * 100))
      }
    } finally {
      await ocrWorker?.terminate()
    }
    if (!pages.length) throw new Error("لم أجد نصاً واضحاً في PDF. جرّب ملفاً أوضح أو تحقق من الصفحات.")
    return pages.join("\n\n")
  }

  if (["pptx", "pptm", "ppsx"].includes(extension ?? "")) {
    const archive = await JSZip.loadAsync(await file.arrayBuffer())
    const slideNames = Object.keys(archive.files)
      .filter(name => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
      .sort((left, right) => Number(left.match(/slide(\d+)/i)?.[1]) - Number(right.match(/slide(\d+)/i)?.[1]))
    if (!slideNames.length) throw new Error("لم أجد شرائح قابلة للقراءة في ملف PowerPoint.")
    const slides: string[] = []
    for (const [index, name] of slideNames.entries()) {
      const xml = await archive.file(name)?.async("string")
      if (!xml) continue
      const document = new DOMParser().parseFromString(xml, "application/xml")
      const paragraphs = Array.from(document.getElementsByTagName("a:p"))
        .map(paragraph => Array.from(paragraph.getElementsByTagName("a:t")).map(node => node.textContent ?? "").join("").trim())
        .filter(Boolean)
      if (paragraphs.length) slides.push(`الشريحة ${index + 1}\n${paragraphs.join("\n")}`)
    }
    if (!slides.length) throw new Error("لم أجد نصاً قابلاً للقراءة في شرائح PowerPoint.")
    return slides.join("\n\n")
  }

  if (["docx", "docm", "dotx", "odt"].includes(extension ?? "")) {
    const archive = await JSZip.loadAsync(await file.arrayBuffer())
    const isWordXml = extension !== "odt"
    const documentPath = isWordXml ? "word/document.xml" : "content.xml"
    const xml = await archive.file(documentPath)?.async("string")
    if (!xml) throw new Error(`لم أجد محتوى قابلاً للقراءة في ملف ${(extension ?? "").toUpperCase()}.`)
    const document = new DOMParser().parseFromString(xml, "application/xml")
    if (isWordXml) {
      const paragraphs = Array.from(document.getElementsByTagName("w:p")).map(paragraph => {
        let value = ""
        for (const node of Array.from(paragraph.childNodes)) {
          if (node.nodeName === "w:r") {
            for (const part of Array.from(node.childNodes)) {
              if (part.nodeName === "w:t") value += part.textContent ?? ""
              else if (part.nodeName === "w:tab") value += "\t"
              else if (part.nodeName === "w:br") value += "\n"
            }
          } else if (node.nodeName === "w:hyperlink") {
            value += Array.from(node.childNodes)
              .flatMap(run => Array.from(run.childNodes))
              .filter(part => part.nodeName === "w:t")
              .map(part => part.textContent ?? "")
              .join("")
          }
        }
        return value.trim()
      }).filter(Boolean)
      if (!paragraphs.length) throw new Error("لم أجد نصاً قابلاً للقراءة في مستند Word.")
      return paragraphs.join("\n")
    }
    const textNamespace = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    const paragraphs = [
      ...Array.from(document.getElementsByTagNameNS(textNamespace, "h")),
      ...Array.from(document.getElementsByTagNameNS(textNamespace, "p")),
    ]
      .sort((left, right) => {
        const order = left.compareDocumentPosition(right)
        return order & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : order & Node.DOCUMENT_POSITION_PRECEDING ? 1 : 0
      })
      .map(paragraph => paragraph.textContent?.trim() ?? "")
      .filter(Boolean)
    if (!paragraphs.length) throw new Error("لم أجد نصاً قابلاً للقراءة في مستند OpenDocument.")
    return paragraphs.join("\n")
  }

  if (["xlsx", "xls", "xlsm", "xlsb", "ods", "csv"].includes(extension ?? "")) {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true })
    const sheets = workbook.SheetNames.flatMap(name => {
      const worksheet = workbook.Sheets[name]
      if (!worksheet) return []
      const text = XLSX.utils.sheet_to_csv(worksheet, { FS: "\t", blankrows: false }).trim()
      return text ? [`${name}\n${text}`] : []
    })
    if (!sheets.length) throw new Error("لم أجد بيانات قابلة للقراءة في الجدول.")
    return sheets.join("\n\n")
  }

  if (extension === "html" || extension === "htm") {
    const html = await file.text()
    const document = new DOMParser().parseFromString(html, "text/html")
    document.querySelectorAll("script,style,noscript,template").forEach(node => node.remove())
    const text = (document.body.innerText || document.body.textContent || "").replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim()
    if (text) return text
    throw new Error("لم أجد نصاً قابلاً للقراءة في صفحة HTML.")
  }

  if (extension === "rtf") {
    const rtf = await file.text()
    const text = rtf
      .replace(/\\'[0-9a-f]{2}/gi, byte => String.fromCharCode(parseInt(byte.slice(2), 16)))
      .replace(/\\(?:par|line)\b ?/gi, "\n")
      .replace(/\\tab\b ?/gi, "\t")
      .replace(/\\[a-z]+-?\d* ?/gi, "")
      .replace(/[{}]/g, "")
      .replace(/\\([\\{}])/g, "$1")
      .replace(/\n{3,}/g, "\n\n")
      .trim()
    if (text) return text
    throw new Error("لم أجد نصاً قابلاً للقراءة في ملف RTF.")
  }

  if (extension === "txt" || extension === "md" || extension === "markdown") {
    const text = await file.text()
    if (text.trim()) return text
    throw new Error("الملف النصي فارغ.")
  }

  if (extension === "ppt" || extension === "doc") {
    throw new Error(`صيغة Office القديمة (.${extension}) غير مدعومة للقراءة مباشرة. افتح الملف في Office واحفظه بصيغة ${extension === "ppt" ? ".pptx" : ".docx"} ثم أعد رفعه.`)
  }
  throw new Error("الصيغ المدعومة: PDF، DOCX، ODT، PPTX، XLSX، XLS، CSV، TXT، MD، RTF، HTML، والصور عبر تبويب الصورة.")
}
