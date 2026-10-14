import { centralSession, centralStorageEnabled } from "./central-storage"

interface SourceFileGroup {
  key: string
  owner: string
  stage: "stageA" | "stageB" | "stageC"
  source: string
  files: File[]
}

const databaseName = "taswiyah-session-files"
const storeName = "files"
let databasePromise: Promise<IDBDatabase> | null = null

function openDatabase(): Promise<IDBDatabase> {
  if (!databasePromise) {
    databasePromise = new Promise((resolve, reject) => {
      const request = indexedDB.open(databaseName, 2)
      request.onupgradeneeded = () => {
        const database = request.result
        if (database.objectStoreNames.contains(storeName) && request.transaction?.objectStore(storeName).keyPath !== "key") {
          database.deleteObjectStore(storeName)
        }
        if (!database.objectStoreNames.contains(storeName)) {
          database.createObjectStore(storeName, { keyPath: "key" })
        }
      }
      request.onsuccess = () => resolve(request.result)
      request.onerror = () => reject(request.error)
    })
  }
  return databasePromise
}

export async function saveStageASourceFiles(stage: SourceFileGroup["stage"], source: string, files: File[], append = false): Promise<void> {
  const owner = await getOwner()
  const key = `${owner}:${stage}:${source}`
  const database = await openDatabase()
  await new Promise<void>((resolve, reject) => {
    const transaction = database.transaction(storeName, "readwrite")
    const store = transaction.objectStore(storeName)
    const request = store.get(key)
    request.onsuccess = () => {
      const persist = (previous?: SourceFileGroup) => {
        store.put({ key, owner, stage, source, files: append ? [...(previous?.files ?? []), ...files] : files } satisfies SourceFileGroup)
      }
      if (request.result || stage !== "stageA") {
        persist(request.result as SourceFileGroup | undefined)
        return
      }
      const legacyKey = `${owner}:${source}`
      const legacyRequest = store.get(legacyKey)
      legacyRequest.onsuccess = () => {
        persist(legacyRequest.result as SourceFileGroup | undefined)
        if (legacyRequest.result) store.delete(legacyKey)
      }
      legacyRequest.onerror = () => reject(legacyRequest.error)
    }
    request.onerror = () => reject(request.error)
    transaction.oncomplete = () => resolve()
    transaction.onerror = () => reject(transaction.error)
  })
}

export async function getStageASourceFiles(): Promise<SourceFileGroup[]> {
  const owner = await getOwner()
  const database = await openDatabase()
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(storeName, "readonly")
    const request = transaction.objectStore(storeName).getAll()
    request.onsuccess = () => resolve((request.result as SourceFileGroup[]).filter(group => group.owner === owner))
    request.onerror = () => reject(request.error)
  })
}

export async function clearStageASourceFiles(stage?: SourceFileGroup["stage"]): Promise<void> {
  const owner = await getOwner()
  const database = await openDatabase()
  await new Promise<void>((resolve, reject) => {
    const transaction = database.transaction(storeName, "readwrite")
    const store = transaction.objectStore(storeName)
    const request = store.getAllKeys()
    request.onsuccess = () => {
      for (const key of request.result) {
        const entryRequest = store.get(key)
        entryRequest.onsuccess = () => {
          const group = entryRequest.result as SourceFileGroup | undefined
          if (group?.owner === owner && (!stage || group.stage === stage)) store.delete(key)
        }
      }
    }
    request.onerror = () => reject(request.error)
    transaction.oncomplete = () => resolve()
    transaction.onerror = () => reject(transaction.error)
  })
}

async function getOwner(): Promise<string> {
  if (!centralStorageEnabled) return "local"
  const user = await centralSession()
  if (!user) throw new Error("يجب تسجيل الدخول لحفظ ملفات الجلسة بأمان.")
  return user.username
}
