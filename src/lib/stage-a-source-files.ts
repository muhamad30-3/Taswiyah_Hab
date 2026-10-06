import { centralSession, centralStorageEnabled } from "./central-storage"

interface SourceFileGroup {
  key: string
  owner: string
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

export async function saveStageASourceFiles(source: string, files: File[], append = false): Promise<void> {
  const owner = await getOwner()
  const key = `${owner}:${source}`
  const database = await openDatabase()
  await new Promise<void>((resolve, reject) => {
    const transaction = database.transaction(storeName, "readwrite")
    const store = transaction.objectStore(storeName)
    const request = store.get(key)
    request.onsuccess = () => {
      const previous = request.result as SourceFileGroup | undefined
      store.put({ key, owner, source, files: append ? [...(previous?.files ?? []), ...files] : files } satisfies SourceFileGroup)
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

export async function clearStageASourceFiles(): Promise<void> {
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
          if (group?.owner === owner) store.delete(key)
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
