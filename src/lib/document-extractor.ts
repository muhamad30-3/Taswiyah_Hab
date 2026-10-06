export interface ExtractedFact {
  label: string
  value: string
}

export interface ExtractedContact {
  name: string
  email: string
  phone: string
  sender: string
  date: string
  time: string
}

export interface ExtractedDocument {
  columns: string[]
  rows: string[][]
  facts: ExtractedFact[]
  contacts: ExtractedContact[]
}

const factLabels: Array<{ label: string; pattern: RegExp }> = [
  { label: "المصروفات اليومية", pattern: /(?:إجمالي\s+)?(?:المصروفات|مصروفات)(?:\s+اليومية)?|daily\s+expenses?|expenses?/i },
  { label: "الخصم", pattern: /(?:مبلغ\s+)?الخصم|الخصومات|discounts?/i },
  { label: "الأصول", pattern: /الأصول|الاصول|assets?/i },
  { label: "الإيرادات", pattern: /الإيرادات|المبيعات|revenues?|sales|income/i },
  { label: "الربح", pattern: /(?:صافي\s+)?الربح|profits?/i },
  { label: "الرصيد", pattern: /الرصيد|balance/i },
]

function normalizeDigits(value: string) {
  return value
    .replace(/[٠-٩]/g, digit => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[\u200e\u200f\u202a-\u202e\u2066-\u2069]/g, "")
}

function extractMessageDateTime(timestamp: string) {
  const dateMatch = timestamp.match(/\d{4}[/-]\d{1,2}[/-]\d{1,2}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/)
  const timeMatch = timestamp.match(/\d{1,2}:\d{2}\s*(?:[ap]\.?m\.?|ص|م)?/i)
  return {
    date: dateMatch?.[0] ?? "",
    time: timeMatch?.[0]?.trim() ?? "",
  }
}

function contactNameCandidates(context: string) {
  const excluded = /(?:^|\s)(?:انا|أنا|من|في|بقدر|ما|وجاهي|الكتروني|إلكتروني|لقاء|لقاءاته|الفصل|كله|مش|اجباري|إجباري|i|from|can|cannot|attend|meeting|meetings|online|in[ -]person)(?=\s|$)/i
  const candidates = context.split(/\n/).flatMap(line => {
    let candidate = line.replace(/\[[^\]]*\]/g, " ").replace(/<[^>]*>/g, " ").trim()
    const separator = candidate.lastIndexOf(":")
    if (separator >= 0) candidate = candidate.slice(separator + 1).trim()
    candidate = candidate.replace(/^[\s\-–—•|,،;؛]+|[\s\-–—•|,،;؛]+$/g, "")
    if (!candidate || /[@\d]/.test(candidate) || excluded.test(candidate)) return []
    const words = candidate.split(/\s+/)
    if (words.length > 7 || !words.every(word => /^\p{L}[\p{L}'’\-]*$/u.test(word))) return []
    if (words.length === 1 && !/[\u0600-\u06ff]/.test(candidate)) return []
    return [candidate]
  })
  return candidates
}

function extractWhatsAppContacts(text: string): ExtractedContact[] {
  const messages: Array<{ text: string; timestamp: string }> = []
  let currentMessage: string[] = []
  let currentTimestamp = ""
  let looseLines: string[] = []
  const flushLooseLines = () => {
    if (looseLines.length) messages.push({ text: looseLines.join("\n"), timestamp: "" })
    looseLines = []
  }
  const flushMessage = () => {
    if (currentMessage.length) messages.push({ text: currentMessage.join(" ").trim(), timestamp: currentTimestamp })
    currentMessage = []
    currentTimestamp = ""
  }

  for (const line of text.split("\n")) {
    const timestampPattern = /\[([^\]]{5,80})\]\s*/g
    const timestamps = [...line.matchAll(timestampPattern)].filter(match =>
      /\d{1,2}:\d{2}/.test(match[1]) ||
      /\d{4}[/-]\d{1,2}[/-]\d{1,2}/.test(match[1]) ||
      /\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/.test(match[1])
    )
    if (timestamps.length) {
      flushMessage()
      flushLooseLines()
      const prefix = line.slice(0, timestamps[0].index).trim()
      if (prefix) messages.push({ text: prefix, timestamp: "" })
      timestamps.forEach((timestamp, index) => {
        const start = (timestamp.index ?? 0) + timestamp[0].length
        const end = index + 1 < timestamps.length ? timestamps[index + 1].index ?? line.length : line.length
        const messageText = line.slice(start, end).trim()
        if (index > 0) flushMessage()
        currentTimestamp = timestamp[1]
        if (messageText) currentMessage.push(messageText)
      })
    } else if (line.trim()) {
      if (currentMessage.length) currentMessage.push(line.trim())
      else looseLines.push(line.trim())
    }
  }
  flushMessage()
  flushLooseLines()

  const contacts: ExtractedContact[] = []
  const seen = new Set<string>()
  const emailPattern = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi
  const phonePattern = /(?<![A-Z0-9])(?:\+\s?\d{1,3}[\s().-]*)?\d(?:[\s().-]*\d){7,14}(?!\d)/gi

  for (const { text, timestamp } of messages) {
    const message = text
    const emailMatches = [...message.matchAll(emailPattern)]
    let previousEmailEnd = 0
    const messagePhones = [...message.matchAll(phonePattern)].map(match => match[0].replace(/[^\d+]/g, ""))
    const senderSeparator = message.match(/^(.{1,100}?):\s*/)
    const senderText = senderSeparator?.[1].replace(/^["'⁦⁩]+|["'⁦⁩]+$/g, "") ?? ""
    const senderPhones = [...senderText.matchAll(phonePattern)].map(match => match[0].replace(/[^\d+]/g, ""))

    for (const emailMatch of emailMatches) {
      const email = emailMatch[0]
      const emailStart = emailMatch.index ?? 0
      const context = message.slice(previousEmailEnd, emailStart)
      previousEmailEnd = emailStart + email.length
      const candidates = contactNameCandidates(context)
      const name = candidates[candidates.length - 1] ?? ""
      const phone = [...context.matchAll(phonePattern)].map(match => match[0].replace(/[^\d+]/g, "")).at(-1)
        ?? (emailMatches.length === 1 ? messagePhones[0] : "")
        ?? ""
      if (!name && !phone) continue
      const key = `${name.toLowerCase()}|${email.toLowerCase()}|${phone}`
      if (seen.has(key)) continue
      seen.add(key)
      contacts.push({
        name,
        email,
        phone,
        sender: senderPhones.length ? "" : senderText,
        ...extractMessageDateTime(timestamp),
      })
    }

    if (emailMatches.length) continue
    const sender = senderText
    const senderPhone = senderPhones[0] ?? ""
    const body = message.slice(senderSeparator?.[0].length ?? 0).replace(phonePattern, " ").trim()
    const name = contactNameCandidates(body).at(-1) ?? ""
    const phone = senderPhone || messagePhones[0] || ""
    if (!phone || !name) continue
    const key = `${name.toLowerCase()}||${phone}`
    if (seen.has(key)) continue
    seen.add(key)
    contacts.push({
      name,
      email: "",
      phone,
      sender: senderPhone ? "" : sender,
      ...extractMessageDateTime(timestamp),
    })
  }
  return contacts
}

function splitDelimitedLine(line: string): string[] {
  if (line.includes("\t")) return line.split("\t").map(cell => cell.trim())
  return line.trim().split(/\s{2,}/).map(cell => cell.trim())
}

function extractTable(lines: string[]) {
  for (let headerIndex = 0; headerIndex < lines.length; headerIndex++) {
    const header = lines[headerIndex]
    if (!header.includes("\t") && !/\s{2,}/.test(header)) continue

    const columns = splitDelimitedLine(header)
    if (columns.length < 2 || !/(date|التاريخ|amount|المبلغ|transaction|البيان)/i.test(header)) continue

    const rows: string[][] = []
    for (const line of lines.slice(headerIndex + 1)) {
      if (!line.trim()) continue
      if (!line.includes("\t") && !/\s{2,}/.test(line)) break
      const cells = splitDelimitedLine(line)
      if (cells.length < 2) break
      rows.push(Array.from({ length: columns.length }, (_, index) => cells[index] ?? ""))
    }
    if (rows.length) return { columns, rows }
  }
  return { columns: [], rows: [] }
}

function extractFacts(text: string): ExtractedFact[] {
  const normalized = normalizeDigits(text)
  const facts: ExtractedFact[] = []
  const seen = new Set<string>()

  for (const { label, pattern } of factLabels) {
    for (const match of normalized.matchAll(new RegExp(pattern.source, pattern.flags.includes("g") ? pattern.flags : `${pattern.flags}g`))) {
      const followingText = normalized.slice((match.index ?? 0) + match[0].length, (match.index ?? 0) + match[0].length + 80)
      const amountMatch = followingText.match(/^\s*(?::|=|يبلغ|بلغت|هو|هي|مقداره)?\s*(?:₪|ريال(?:\s+سعودي)?|SAR)?\s*([-+]?\d[\d,]*(?:\.\d+)?)/i)
        ?? followingText.match(/^[^.!?\n]{0,30}?([-+]?\d[\d,]*(?:\.\d+)?)/)
      if (!amountMatch) continue

      const value = amountMatch[1].replace(/,$/, "")
      const key = `${label}:${value}`
      if (seen.has(key)) continue
      seen.add(key)
      facts.push({ label, value })
    }
  }
  return facts
}

function extractLooseTransactions(text: string) {
  const pattern = /^\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{6,8})\s+(\d{1,2}:\d{2}|\d{3,4})\s+(.+?)\s+(?:₪\s*)?([-+]?\d[\d,]*(?:\.\d{2})?)\s*$/i
  const rows = text.split("\n").flatMap(line => {
    const match = line.match(pattern)
    if (!match) return []
    const time = match[2].includes(":") ? match[2] : match[2].padStart(4, "0").replace(/^(\d{2})(\d{2})$/, "$1:$2")
    return [[match[1], time, match[3].trim(), match[4]]]
  })
  return rows.length >= 3
    ? { columns: ["التاريخ (مقروء من الصورة)", "الوقت", "تفاصيل الحركة", "المبلغ المقروء"], rows }
    : { columns: [], rows: [] }
}

export function extractDocument(text: string): ExtractedDocument {
  const normalizedText = normalizeDigits(text).replace(/\r\n?/g, "\n")
  const table = extractTable(normalizedText.split("\n"))
  const looseTable = table.rows.length ? table : extractLooseTransactions(normalizedText)
  return {
    ...looseTable,
    facts: extractFacts(normalizedText),
    contacts: extractWhatsAppContacts(normalizedText),
  }
}
