export const centralStorageEnabled = import.meta.env.VITE_CENTRAL_STORAGE === "true"

type ApiError = { error?: string }
export type CentralUser = { username: string; role: "admin" | "user" }
export type AdminUser = {
  id: string
  username: string
  email: string | null
  role: "admin" | "user"
  status: "pending" | "approved" | "rejected"
  createdAt: string
}
export type StoredFile = {
  id: string
  fileName: string
  mimeType: string
  size: number
  uploadedAt: string
}

let currentUserRole: CentralUser["role"] | null = null
let currentUsername: string | null = null
const storageCacheLifetime = 15_000
const localBackupPrefix = "taswiyah:central-backup:"
const localBackupDatabaseName = "taswiyah-central-backups"
const localBackupStoreName = "backups"
const storageCache = new Map<string, { value: unknown; expiresAt: number }>()
const storageRequests = new Map<string, Promise<unknown>>()
const storageVersions = new Map<string, number>()
let storageGeneration = 0
let localBackupDatabase: Promise<IDBDatabase> | null = null
let isFlushingLocalBackups = false
let localBackupRetryTimer: number | null = null
let localBackupRetryAttempt = 0
const localBackupVersions = new Map<string, number>()
const pendingStorageWrites = new Map<string, { backup: LocalBackup; waiters: Array<{ resolve: () => void; reject: (error: unknown) => void }> }>()
const activeStorageWrites = new Set<string>()

type LocalBackup = {
  id: string
  username: string
  key: string
  updatedAt: number
  syncPending: boolean
  deleted?: boolean
  value?: unknown
}

function clearStorageCache() {
  storageGeneration++
  storageCache.clear()
  storageRequests.clear()
  storageVersions.clear()
}

function invalidateStorageKey(key: string) {
  storageVersions.set(key, (storageVersions.get(key) ?? 0) + 1)
  storageCache.delete(key)
  storageRequests.delete(key)
}

if (typeof window !== "undefined") {
  window.addEventListener("focus", clearStorageCache)
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") clearStorageCache()
  })
}

export function isCentralAdmin() {
  return currentUserRole === "admin"
}

function reportStorageFailure(error: unknown) {
  if (typeof window === "undefined") return
  const message = error instanceof Error ? error.message : "حدث خطأ في التخزين المركزي."
  window.dispatchEvent(new CustomEvent("central-storage-error", { detail: message }))
}

async function readResponse<T>(response: Response): Promise<T> {
  const body = await response.json().catch(() => ({})) as T & ApiError
  if (!response.ok) throw new Error(body.error || `تعذر الاتصال بخادم التخزين (HTTP ${response.status}).`)
  return body
}

function localBackupId(username: string, key: string) {
  return `${username}:${key}`
}

function localBackupStorageKey(username: string, key: string) {
  return `${localBackupPrefix}${encodeURIComponent(username)}:${encodeURIComponent(key)}`
}

function openLocalBackupDatabase(): Promise<IDBDatabase> {
  if (localBackupDatabase) return localBackupDatabase
  const opening = new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(localBackupDatabaseName, 1)
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(localBackupStoreName)) {
        request.result.createObjectStore(localBackupStoreName, { keyPath: "id" })
      }
    }
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error ?? new Error("تعذر فتح التخزين المحلي الاحتياطي."))
  })
  const tracked = opening.catch(error => {
    localBackupDatabase = null
    throw error
  })
  localBackupDatabase = tracked
  return tracked
}

async function readLocalBackup(username: string, key: string): Promise<LocalBackup | undefined> {
  const storageKey = localBackupStorageKey(username, key)
  let backup: LocalBackup | undefined
  try {
    const serialized = localStorage.getItem(storageKey)
    if (serialized) backup = JSON.parse(serialized) as LocalBackup
  } catch (error) {
    console.warn("تعذرت قراءة النسخة المحلية من Local Storage؛ ستتم محاولة IndexedDB:", error)
  }
  try {
    const database = await openLocalBackupDatabase()
    const indexedDbBackup = await new Promise<LocalBackup | undefined>((resolve, reject) => {
      const request = database.transaction(localBackupStoreName, "readonly").objectStore(localBackupStoreName).get(localBackupId(username, key))
      request.onsuccess = () => resolve(request.result as LocalBackup | undefined)
      request.onerror = () => reject(request.error)
    })
    if (indexedDbBackup && (!backup || indexedDbBackup.updatedAt >= backup.updatedAt)) backup = indexedDbBackup
  } catch (error) {
    console.error("تعذرت قراءة النسخة المحلية الاحتياطية:", error)
  }
  return backup
}

async function saveLocalBackup(backup: LocalBackup): Promise<void> {
  const storageKey = localBackupStorageKey(backup.username, backup.key)
  try {
    localStorage.setItem(storageKey, JSON.stringify(backup))
    return
  } catch (localStorageError) {
    try {
      const database = await openLocalBackupDatabase()
      await new Promise<void>((resolve, reject) => {
        const transaction = database.transaction(localBackupStoreName, "readwrite")
        transaction.objectStore(localBackupStoreName).put(backup)
        transaction.oncomplete = () => resolve()
        transaction.onerror = () => reject(transaction.error)
        transaction.onabort = () => reject(transaction.error ?? new Error("تعذر إتمام النسخة الاحتياطية المحلية."))
      })
      console.warn("تعذر Local Storage؛ حُفظت النسخة الاحتياطية في IndexedDB:", localStorageError)
      return
    } catch (indexedDbError) {
      throw new Error("تعذر إنشاء نسخة احتياطية محلية. أفرغ مساحة من المتصفح قبل متابعة التعديلات.", { cause: indexedDbError })
    }
  }
}

async function removeLocalBackup(username: string, key: string): Promise<void> {
  try {
    localStorage.removeItem(localBackupStorageKey(username, key))
  } catch (error) {
    console.warn("تعذر حذف النسخة الاحتياطية من Local Storage:", error)
  }
  try {
    const database = await openLocalBackupDatabase()
    await new Promise<void>((resolve, reject) => {
      const transaction = database.transaction(localBackupStoreName, "readwrite")
      transaction.objectStore(localBackupStoreName).delete(localBackupId(username, key))
      transaction.oncomplete = () => resolve()
      transaction.onerror = () => reject(transaction.error)
    })
  } catch (error) {
    console.error("تعذر حذف النسخة الاحتياطية من IndexedDB:", error)
  }
}

async function listLocalBackups(): Promise<LocalBackup[]> {
  const backups = new Map<string, LocalBackup>()
  if (currentUsername) {
    const prefix = `${localBackupPrefix}${encodeURIComponent(currentUsername)}:`
    try {
      for (let index = 0; index < localStorage.length; index++) {
        const storageKey = localStorage.key(index)
        if (!storageKey?.startsWith(prefix)) continue
        const serialized = localStorage.getItem(storageKey)
        if (!serialized) continue
        const backup = JSON.parse(serialized) as LocalBackup
        if (backup.username === currentUsername && backup.syncPending) backups.set(backup.id, backup)
      }
    } catch (error) {
      console.warn("تعذر فحص النسخ الاحتياطية في Local Storage:", error)
    }
  }
  try {
    const database = await openLocalBackupDatabase()
    const stored = await new Promise<LocalBackup[]>((resolve, reject) => {
      const request = database.transaction(localBackupStoreName, "readonly").objectStore(localBackupStoreName).getAll()
      request.onsuccess = () => resolve(request.result as LocalBackup[])
      request.onerror = () => reject(request.error)
    })
    for (const backup of stored) {
      if (backup.username !== currentUsername || !backup.syncPending) continue
      const previous = backups.get(backup.id)
      if (!previous || backup.updatedAt >= previous.updatedAt) backups.set(backup.id, backup)
    }
  } catch (error) {
    console.warn("تعذر فحص النسخ الاحتياطية في IndexedDB:", error)
  }
  return [...backups.values()]
}

async function writeBackupToApi(backup: LocalBackup): Promise<void> {
  const endpoint = `/api/storage/${encodeURIComponent(backup.key)}`
  const response = await fetch(endpoint, {
    method: backup.deleted ? "DELETE" : "PUT",
    credentials: "include",
    ...(backup.deleted ? {} : {
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ value: backup.value }),
    }),
  })
  if (backup.deleted) {
    if (!response.ok) await readResponse(response)
  } else {
    const body = await readResponse<{ saved: boolean }>(response)
    if (!body.saved) throw new Error("لم يؤكد الخادم حفظ البيانات.")
  }
  const latest = await readLocalBackup(backup.username, backup.key)
  if (latest?.updatedAt === backup.updatedAt) {
    if (backup.deleted) await removeLocalBackup(backup.username, backup.key)
    else await saveLocalBackup({ ...backup, syncPending: false })
  }
  invalidateStorageKey(backup.key)
}

function enqueueStorageWrite(backup: LocalBackup): Promise<void> {
  return new Promise((resolve, reject) => {
    const pending = pendingStorageWrites.get(backup.key)
    if (pending) {
      if (backup.updatedAt >= pending.backup.updatedAt) pending.backup = backup
      pending.waiters.push({ resolve, reject })
    } else {
      pendingStorageWrites.set(backup.key, { backup, waiters: [{ resolve, reject }] })
    }
    if (activeStorageWrites.has(backup.key)) return
    activeStorageWrites.add(backup.key)
    void (async () => {
      try {
        while (true) {
          const write = pendingStorageWrites.get(backup.key)
          if (!write) break
          pendingStorageWrites.delete(backup.key)
          try {
            await writeBackupToApi(write.backup)
            for (const waiter of write.waiters) waiter.resolve()
          } catch (error) {
            const reportedError = new Error(`${error instanceof Error ? error.message : "فشل الاتصال بقاعدة البيانات."} حُفظت نسخة احتياطية محلية في هذا المتصفح وستتم إعادة المحاولة عند عودة الاتصال.`)
            reportStorageFailure(reportedError)
            for (const waiter of write.waiters) waiter.reject(error)
          }
        }
      } finally {
        activeStorageWrites.delete(backup.key)
      }
    })()
  })
}

async function flushLocalBackups(): Promise<void> {
  if (localBackupRetryTimer !== null) {
    window.clearTimeout(localBackupRetryTimer)
    localBackupRetryTimer = null
  }
  if (!currentUsername || !navigator.onLine || isFlushingLocalBackups) return
  isFlushingLocalBackups = true
  let allSynced = true
  let pendingBackups: LocalBackup[] = []
  try {
    pendingBackups = await listLocalBackups()
    for (const backup of pendingBackups) {
      if (backup.username !== currentUsername) continue
      try {
        await enqueueStorageWrite(backup)
      } catch (error) {
        allSynced = false
        console.warn(`تعذرت مزامنة النسخة الاحتياطية «${backup.key}»:`, error)
      }
    }
  } finally {
    isFlushingLocalBackups = false
  }
  if (allSynced) {
    localBackupRetryAttempt = 0
    if (pendingBackups.length > 0) window.dispatchEvent(new Event("central-storage-recovered"))
  } else if (navigator.onLine) {
    const delay = Math.min(2_000 * (2 ** localBackupRetryAttempt), 30_000)
    localBackupRetryAttempt = Math.min(localBackupRetryAttempt + 1, 4)
    localBackupRetryTimer = window.setTimeout(() => {
      localBackupRetryTimer = null
      void flushLocalBackups()
    }, delay)
  }
}

if (typeof window !== "undefined") {
  window.addEventListener("online", () => { void flushLocalBackups() })
  window.addEventListener("focus", () => { void flushLocalBackups() })
}

export async function centralLogin(username: string, password: string): Promise<CentralUser> {
  const response = await fetch("/api/auth/login", {
    method: "POST",
    credentials: "include",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ username, password }),
  })
  const body = await readResponse<CentralUser>(response)
  if (currentUsername !== body.username) clearStorageCache()
  currentUsername = body.username
  currentUserRole = body.role
  void flushLocalBackups()
  return body
}

export async function centralRegister(username: string, email: string, password: string): Promise<void> {
  const response = await fetch("/api/auth/register", {
    method: "POST",
    credentials: "include",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ username, email, password }),
  })
  await readResponse<{ registered: boolean; status: string }>(response)
}

export async function centralChangePassword(currentPassword: string, newPassword: string): Promise<void> {
  const response = await fetch("/api/auth/change-password", {
    method: "POST",
    credentials: "include",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ currentPassword, newPassword }),
  })
  await readResponse<{ updated: boolean }>(response)
}

export async function centralSession(): Promise<CentralUser | null> {
  const response = await fetch("/api/auth/me", { credentials: "include" })
  if (response.status === 401) {
    if (currentUsername !== null) clearStorageCache()
    currentUsername = null
    currentUserRole = null
    return null
  }
  const body = await readResponse<CentralUser>(response)
  if (currentUsername !== body.username) clearStorageCache()
  currentUsername = body.username
  currentUserRole = body.role
  void flushLocalBackups()
  return body
}

export async function centralLogout(): Promise<void> {
  try {
    const response = await fetch("/api/auth/logout", { method: "POST", credentials: "include" })
    if (!response.ok) await readResponse(response)
  } finally {
    currentUsername = null
    currentUserRole = null
    clearStorageCache()
  }
}

export async function listAdminUsers(): Promise<AdminUser[]> {
  const response = await fetch("/api/admin/users", { credentials: "include" })
  const body = await readResponse<{ users: AdminUser[] }>(response)
  return body.users
}

export async function updateAdminUser(id: string, action: "approve" | "reject"): Promise<void> {
  const response = await fetch(`/api/admin/users/${encodeURIComponent(id)}/${action}`, {
    method: "POST",
    credentials: "include",
  })
  await readResponse<{ updated: boolean }>(response)
}

export async function listStoredFiles(): Promise<StoredFile[]> {
  const response = await fetch("/api/files", { credentials: "include" })
  const body = await readResponse<{ files: StoredFile[] }>(response)
  return body.files
}

export async function uploadStoredFile(file: File, fileName = file.name): Promise<string> {
  if (file.size > 25 * 1024 * 1024) throw new Error(`«${file.name}» يتجاوز الحد الأقصى 25 ميغابايت.`)
  const response = await fetch("/api/files", {
    method: "POST",
    credentials: "include",
    headers: {
      "content-type": "application/octet-stream",
      "x-file-name": encodeURIComponent(fileName),
      "x-file-type": file.type || "application/octet-stream",
    },
    body: file,
  })
  const body = await readResponse<{ id: string }>(response)
  return body.id
}

export async function fetchStoredFile(id: string, fileName: string, mimeType: string): Promise<File> {
  const response = await fetch(`/api/files/${encodeURIComponent(id)}`, { credentials: "include" })
  if (!response.ok) {
    const body = await response.json().catch(() => ({})) as ApiError
    throw new Error(body.error || `تعذر فتح الملف (HTTP ${response.status}).`)
  }
  return new File([await response.blob()], fileName, { type: mimeType })
}

export async function downloadStoredFile(file: StoredFile): Promise<void> {
  const response = await fetch(`/api/files/${encodeURIComponent(file.id)}`, { credentials: "include" })
  if (!response.ok) {
    const body = await response.json().catch(() => ({})) as ApiError
    throw new Error(body.error || `تعذر تنزيل الملف (HTTP ${response.status}).`)
  }
  const url = URL.createObjectURL(await response.blob())
  const anchor = document.createElement("a")
  anchor.href = url
  anchor.download = file.fileName
  anchor.click()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}

export async function deleteStoredFile(id: string): Promise<void> {
  const response = await fetch(`/api/files/${encodeURIComponent(id)}`, {
    method: "DELETE",
    credentials: "include",
  })
  if (!response.ok) await readResponse(response)
}

export async function centralStorageGet<T>(key: string): Promise<T | undefined> {
  const cached = storageCache.get(key)
  if (cached && cached.expiresAt > Date.now()) return cached.value as T | undefined
  if (cached) storageCache.delete(key)

  const pending = storageRequests.get(key)
  if (pending) return pending as Promise<T | undefined>

  const generation = storageGeneration
  const version = storageVersions.get(key) ?? 0
  let request: Promise<T | undefined> | undefined
  try {
    request = (async (): Promise<T | undefined> => {
      let lastError: unknown
      for (let attempt = 0; attempt < 2; attempt++) {
        let response: Response
        try {
          response = await fetch(`/api/storage/${encodeURIComponent(key)}`, { credentials: "include" })
        } catch (error) {
          lastError = error
          if (attempt === 1) throw error
          await new Promise(resolve => window.setTimeout(resolve, 500))
          continue
        }
        if (response.status === 404) {
          const backup = currentUsername ? await readLocalBackup(currentUsername, key) : undefined
          if (backup?.syncPending) {
            void flushLocalBackups()
            return backup.deleted ? undefined : backup.value as T
          }
          if (generation === storageGeneration && version === (storageVersions.get(key) ?? 0)) {
            storageCache.set(key, { value: undefined, expiresAt: Date.now() + storageCacheLifetime })
          }
          return undefined
        }
        if ((response.status === 502 || response.status === 503) && attempt === 0) {
          await new Promise(resolve => window.setTimeout(resolve, 500))
          continue
        }
        const body = await readResponse<{ value: T }>(response)
        const backup = currentUsername ? await readLocalBackup(currentUsername, key) : undefined
        if (backup?.syncPending) {
          void flushLocalBackups()
          return backup.deleted ? undefined : backup.value as T
        }
        if (generation === storageGeneration && version === (storageVersions.get(key) ?? 0)) {
          storageCache.set(key, { value: body.value, expiresAt: Date.now() + storageCacheLifetime })
        }
        if (currentUsername) {
          void saveLocalBackup({
            id: localBackupId(currentUsername, key),
            username: currentUsername,
            key,
            updatedAt: Date.now(),
            syncPending: false,
            value: body.value,
          }).catch(error => console.warn(`تعذر تحديث النسخة المحلية «${key}»:`, error))
        }
        return body.value
      }
      throw lastError instanceof Error ? lastError : new Error("تعذرت قراءة البيانات من قاعدة البيانات.")
    })()
    storageRequests.set(key, request)
    return await request
  } catch (error) {
    const backup = currentUsername ? await readLocalBackup(currentUsername, key) : undefined
    if (backup) {
      if (backup.syncPending) void flushLocalBackups()
      reportStorageFailure(new Error(`تعذر الاتصال بقاعدة البيانات؛ تم استرجاع نسخة هذا المتصفح${backup.syncPending ? " وستتم إعادة مزامنتها تلقائياً عند عودة الاتصال" : ""}.`))
      return backup.deleted ? undefined : backup.value as T
    }
    reportStorageFailure(error)
    throw error
  } finally {
    if (request && storageRequests.get(key) === request) storageRequests.delete(key)
  }
}

export async function centralStorageSet(key: string, value: unknown): Promise<void> {
  invalidateStorageKey(key)
  const generation = storageGeneration
  const version = storageVersions.get(key)
  if (!currentUsername) throw new Error("لا يمكن حفظ نسخة محلية قبل التحقق من حساب المستخدم.")
  const previousBackup = await readLocalBackup(currentUsername, key)
  const backupId = localBackupId(currentUsername, key)
  const backup: LocalBackup = {
    id: backupId,
    username: currentUsername,
    key,
    updatedAt: Math.max(Date.now(), (previousBackup?.updatedAt ?? 0) + 1, (localBackupVersions.get(backupId) ?? 0) + 1),
    syncPending: true,
    value,
  }
  localBackupVersions.set(backupId, backup.updatedAt)
  try {
    await saveLocalBackup(backup)
  } catch (error) {
    reportStorageFailure(error)
    throw error
  }
  await enqueueStorageWrite(backup)
  if (generation === storageGeneration && version === storageVersions.get(key)) {
    storageCache.set(key, { value, expiresAt: Date.now() + storageCacheLifetime })
  }
}

export async function centralStorageDelete(key: string): Promise<void> {
  invalidateStorageKey(key)
  if (!currentUsername) throw new Error("لا يمكن حذف بيانات مركزية قبل التحقق من حساب المستخدم.")
  const previousBackup = await readLocalBackup(currentUsername, key)
  const backupId = localBackupId(currentUsername, key)
  const backup: LocalBackup = {
    id: backupId,
    username: currentUsername,
    key,
    updatedAt: Math.max(Date.now(), (previousBackup?.updatedAt ?? 0) + 1, (localBackupVersions.get(backupId) ?? 0) + 1),
    syncPending: true,
    deleted: true,
  }
  localBackupVersions.set(backupId, backup.updatedAt)
  await saveLocalBackup(backup)
  await enqueueStorageWrite(backup)
}
