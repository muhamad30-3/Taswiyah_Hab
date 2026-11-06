import {
  centralStorageDelete,
  centralStorageEnabled,
  centralStorageGet,
  centralStorageSet,
} from "./central-storage"
import {
  EMPTY_TRANSFER_DATA,
  type TransferData,
  type TransferImport,
  type TransferRecord,
} from "./transfers"

const LOCAL_DATABASE_NAME = "taswiyah-transfer-ledger"
const LOCAL_STORE_NAME = "ledgers"
const TRANSFER_CHUNK_SIZE = 2_000
const STORAGE_BATCH_SIZE = 4
let localDatabase: Promise<IDBDatabase> | null = null

type TransferManifest = {
  version: 1
  chunks: string[]
  imports: TransferImport[]
}

export type TransferDataLoadResult = {
  data: TransferData
  warning?: string
}

export type TransferLoadProgress = {
  loadedRecords: number
  loadedChunks: number
  totalChunks: number
}

function openLocalDatabase() {
  if (localDatabase) return localDatabase
  if (typeof indexedDB === "undefined") return Promise.reject(new Error("المتصفح لا يدعم التخزين المحلي الدائم."))
  localDatabase = new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(LOCAL_DATABASE_NAME, 1)
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(LOCAL_STORE_NAME)) {
        request.result.createObjectStore(LOCAL_STORE_NAME, { keyPath: "id" })
      }
    }
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error ?? new Error("تعذر فتح قاعدة بيانات الحوالات المحلية."))
  }).catch(error => {
    localDatabase = null
    throw error
  })
  return localDatabase
}

async function readLocalValue<T>(key: string): Promise<T | undefined> {
  const database = await openLocalDatabase()
  return new Promise<T | undefined>((resolve, reject) => {
    const request = database.transaction(LOCAL_STORE_NAME, "readonly").objectStore(LOCAL_STORE_NAME).get(key)
    request.onsuccess = () => resolve((request.result as { value?: T } | undefined)?.value)
    request.onerror = () => reject(request.error ?? new Error("تعذرت قراءة سجل الحوالات المحلي."))
  })
}

async function writeLocalValue<T>(key: string, value: T): Promise<void> {
  const database = await openLocalDatabase()
  return new Promise<void>((resolve, reject) => {
    const transaction = database.transaction(LOCAL_STORE_NAME, "readwrite")
    transaction.objectStore(LOCAL_STORE_NAME).put({ id: key, value })
    transaction.oncomplete = () => resolve()
    transaction.onerror = () => reject(transaction.error ?? new Error("تعذر حفظ سجل الحوالات المحلي."))
    transaction.onabort = () => reject(transaction.error ?? new Error("أُلغي حفظ سجل الحوالات المحلي."))
  })
}

function manifestKey(storageKey: string) {
  return `${storageKey}:manifest`
}

function createChunkKey() {
  const id = typeof globalThis.crypto?.randomUUID === "function"
    ? globalThis.crypto.randomUUID()
    : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`
  return `transfers_chunk:${id}`
}

async function readValue<T>(key: string): Promise<T | undefined> {
  return centralStorageEnabled ? centralStorageGet<T>(key) : readLocalValue<T>(key)
}

async function writeValue<T>(key: string, value: T): Promise<void> {
  if (centralStorageEnabled) await centralStorageSet(key, value)
  else await writeLocalValue(key, value)
}

async function deleteValue(key: string): Promise<void> {
  if (centralStorageEnabled) await centralStorageDelete(key)
  else {
    const database = await openLocalDatabase()
    await new Promise<void>((resolve, reject) => {
      const transaction = database.transaction(LOCAL_STORE_NAME, "readwrite")
      transaction.objectStore(LOCAL_STORE_NAME).delete(key)
      transaction.oncomplete = () => resolve()
      transaction.onerror = () => reject(transaction.error ?? new Error("تعذر تنظيف جزء سجل الحوالات المحلي."))
      transaction.onabort = () => reject(transaction.error ?? new Error("أُلغي تنظيف جزء سجل الحوالات المحلي."))
    })
  }
}

function isTransferData(value: unknown): value is TransferData {
  return !!value
    && typeof value === "object"
    && Array.isArray((value as TransferData).records)
    && Array.isArray((value as TransferData).imports)
}

function isTransferManifest(value: unknown): value is TransferManifest {
  return !!value
    && typeof value === "object"
    && (value as TransferManifest).version === 1
    && Array.isArray((value as TransferManifest).chunks)
    && (value as TransferManifest).chunks.every(key => typeof key === "string")
    && Array.isArray((value as TransferManifest).imports)
}

function sliceIntoChunks(records: TransferRecord[]) {
  const chunks: TransferRecord[][] = []
  for (let index = 0; index < records.length; index += TRANSFER_CHUNK_SIZE) {
    chunks.push(records.slice(index, index + TRANSFER_CHUNK_SIZE))
  }
  return chunks
}

async function writeRecordChunks(records: TransferRecord[]) {
  const chunkKeys: string[] = []
  const chunks = sliceIntoChunks(records)
  for (let index = 0; index < chunks.length; index += STORAGE_BATCH_SIZE) {
    const batch = chunks.slice(index, index + STORAGE_BATCH_SIZE)
    const batchKeys = batch.map(createChunkKey)
    await Promise.all(batch.map((recordsInChunk, offset) =>
      writeValue(batchKeys[offset], recordsInChunk),
    ))
    chunkKeys.push(...batchKeys)
  }
  return chunkKeys
}

export async function loadTransferData(
  storageKey: string,
  onProgress?: (progress: TransferLoadProgress) => void,
): Promise<TransferDataLoadResult> {
  const storedManifest = await readValue<unknown>(manifestKey(storageKey))
  if (storedManifest !== undefined) {
    if (!isTransferManifest(storedManifest)) throw new Error("فهرس أجزاء سجل الحوالات غير صالح.")
    const records: TransferRecord[] = []
    let loadedChunks = 0
    onProgress?.({ loadedRecords: 0, loadedChunks, totalChunks: storedManifest.chunks.length })
    for (let index = 0; index < storedManifest.chunks.length; index += STORAGE_BATCH_SIZE) {
      const keys = storedManifest.chunks.slice(index, index + STORAGE_BATCH_SIZE)
      const chunks = await Promise.all(keys.map(key => readValue<unknown>(key)))
      for (const chunk of chunks) {
        if (!Array.isArray(chunk)) throw new Error("تعذر تحميل أحد أجزاء سجل الحوالات؛ لم يتم اعتماد بيانات ناقصة.")
        records.push(...chunk as TransferRecord[])
        loadedChunks++
        onProgress?.({
          loadedRecords: records.length,
          loadedChunks,
          totalChunks: storedManifest.chunks.length,
        })
      }
    }
    return { data: { records, imports: storedManifest.imports } }
  }

  if (!centralStorageEnabled) {
    const serialized = localStorage.getItem(storageKey)
    if (serialized) {
      const legacyValue: unknown = JSON.parse(serialized)
      if (!isTransferData(legacyValue)) throw new Error("بيانات الحوالات المحلية المحفوظة غير صالحة.")
      onProgress?.({ loadedRecords: legacyValue.records.length, loadedChunks: 1, totalChunks: 1 })
      return migrateLegacyData(storageKey, legacyValue)
    }
  }
  const legacyData = await readValue<unknown>(storageKey)
  if (legacyData === undefined) {
    onProgress?.({ loadedRecords: 0, loadedChunks: 0, totalChunks: 0 })
    return { data: EMPTY_TRANSFER_DATA }
  }
  if (!isTransferData(legacyData)) throw new Error("بيانات الحوالات المحفوظة غير صالحة.")
  onProgress?.({ loadedRecords: legacyData.records.length, loadedChunks: 1, totalChunks: 1 })
  return migrateLegacyData(storageKey, legacyData)
}

async function migrateLegacyData(storageKey: string, legacyData: TransferData) {
  const manifest: TransferManifest = {
    version: 1,
    chunks: await writeRecordChunks(legacyData.records),
    imports: legacyData.imports,
  }
  await writeValue(manifestKey(storageKey), manifest)
  let warning: string | undefined
  try {
    if (centralStorageEnabled) await centralStorageDelete(storageKey)
    else localStorage.removeItem(storageKey)
  } catch (error) {
    warning = error instanceof Error
      ? `تم تقسيم البيانات وحفظها، لكن تعذر تنظيف النسخة القديمة: ${error.message}`
      : "تم تقسيم البيانات وحفظها، لكن تعذر تنظيف النسخة القديمة."
  }
  return { data: legacyData, warning }
}

export async function appendTransferData(
  storageKey: string,
  existingData: TransferData,
  newRecords: TransferRecord[],
  imported: TransferImport,
): Promise<{ data: TransferData; cleanupWarning: string }> {
  const indexKey = manifestKey(storageKey)
  const storedManifest = await readValue<unknown>(indexKey)
  const isMigration = storedManifest === undefined
  let manifest: TransferManifest
  if (storedManifest !== undefined) {
    if (!isTransferManifest(storedManifest)) throw new Error("فهرس أجزاء سجل الحوالات غير صالح.")
    manifest = storedManifest
  } else {
    const legacyChunks = await writeRecordChunks(existingData.records)
    manifest = {
      version: 1,
      chunks: legacyChunks,
      imports: existingData.imports,
    }
  }

  const appendedChunks = await writeRecordChunks(newRecords)
  const nextManifest: TransferManifest = {
    version: 1,
    chunks: [...manifest.chunks, ...appendedChunks],
    imports: [imported, ...manifest.imports],
  }
  await writeValue(indexKey, nextManifest)

  let cleanupWarning = ""
  if (isMigration) {
    try {
      if (centralStorageEnabled) await centralStorageDelete(storageKey)
      else localStorage.removeItem(storageKey)
    } catch (error) {
      cleanupWarning = error instanceof Error
        ? `تم حفظ سجل الحوالات، لكن تعذر تنظيف النسخة القديمة: ${error.message}`
        : "تم حفظ سجل الحوالات، لكن تعذر تنظيف النسخة القديمة."
    }
  }
  return {
    data: {
      records: [...existingData.records, ...newRecords],
      imports: nextManifest.imports,
    },
    cleanupWarning,
  }
}

export async function replaceTransferData(
  storageKey: string,
  nextData: TransferData,
): Promise<{ cleanupWarning: string }> {
  const indexKey = manifestKey(storageKey)
  const storedManifest = await readValue<unknown>(indexKey)
  if (storedManifest !== undefined && !isTransferManifest(storedManifest)) {
    throw new Error("فهرس أجزاء سجل الحوالات غير صالح.")
  }
  const isLegacyData = storedManifest === undefined
  const previousManifest = isTransferManifest(storedManifest) ? storedManifest : null
  const previousChunkKeys = previousManifest?.chunks ?? []
  const nextRecordIds = new Set(nextData.records.map(record => record.id))
  const coveredRecordIds = new Set<string>()
  const nextChunkKeys: string[] = []
  if (previousManifest) {
    for (let index = 0; index < previousChunkKeys.length; index += STORAGE_BATCH_SIZE) {
      const keys = previousChunkKeys.slice(index, index + STORAGE_BATCH_SIZE)
      const chunks = await Promise.all(keys.map(key => readValue<unknown>(key)))
      for (let offset = 0; offset < keys.length; offset++) {
        const records = chunks[offset]
        if (!Array.isArray(records)) throw new Error("تعذر تحميل أحد أجزاء سجل الحوالات؛ لم يتم اعتماد الحذف.")
        const retained = (records as TransferRecord[]).filter(record => {
          coveredRecordIds.add(record.id)
          return nextRecordIds.has(record.id)
        })
        if (retained.length === records.length) {
          nextChunkKeys.push(keys[offset])
        } else if (retained.length) {
          nextChunkKeys.push(...await writeRecordChunks(retained))
        }
      }
    }
  }
  const newRecords = nextData.records.filter(record => !coveredRecordIds.has(record.id))
  nextChunkKeys.push(...await writeRecordChunks(newRecords))
  const nextManifest: TransferManifest = {
    version: 1,
    chunks: nextChunkKeys,
    imports: nextData.imports,
  }
  await writeValue(indexKey, nextManifest)

  try {
    const retainedChunkKeys = new Set(nextChunkKeys)
    await Promise.all(previousChunkKeys.filter(key => !retainedChunkKeys.has(key)).map(deleteValue))
    if (isLegacyData) {
      if (centralStorageEnabled) await centralStorageDelete(storageKey)
      else localStorage.removeItem(storageKey)
    }
    return { cleanupWarning: "" }
  } catch (error) {
    return {
      cleanupWarning: error instanceof Error
        ? `تم حذف الكشف، لكن تعذر تنظيف بعض الأجزاء القديمة: ${error.message}`
        : "تم حذف الكشف، لكن تعذر تنظيف بعض الأجزاء القديمة.",
    }
  }
}
