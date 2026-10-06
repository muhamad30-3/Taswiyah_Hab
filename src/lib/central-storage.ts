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
const storageCache = new Map<string, { value: unknown; expiresAt: number }>()
const storageRequests = new Map<string, Promise<unknown>>()
const storageVersions = new Map<string, number>()
let storageGeneration = 0

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
        if (generation === storageGeneration && version === (storageVersions.get(key) ?? 0)) {
          storageCache.set(key, { value: body.value, expiresAt: Date.now() + storageCacheLifetime })
        }
        return body.value
      }
      throw lastError instanceof Error ? lastError : new Error("تعذرت قراءة البيانات من قاعدة البيانات.")
    })()
    storageRequests.set(key, request)
    return await request
  } catch (error) {
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
  try {
    const response = await fetch(`/api/storage/${encodeURIComponent(key)}`, {
      method: "PUT",
      credentials: "include",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ value }),
    })
    const body = await readResponse<{ saved: boolean }>(response)
    if (!body.saved) throw new Error("لم يؤكد الخادم حفظ البيانات في قاعدة البيانات.")
    if (generation === storageGeneration && version === storageVersions.get(key)) {
      storageCache.set(key, { value, expiresAt: Date.now() + storageCacheLifetime })
    }
  } catch (error) {
    reportStorageFailure(error)
    throw error
  }
}

export async function centralStorageDelete(key: string): Promise<void> {
  invalidateStorageKey(key)
  try {
    const response = await fetch(`/api/storage/${encodeURIComponent(key)}`, {
      method: "DELETE",
      credentials: "include",
    })
    if (!response.ok) await readResponse(response)
  } catch (error) {
    reportStorageFailure(error)
    throw error
  }
}
