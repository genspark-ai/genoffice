/**
 * Google Drive API calls (main process only; renderer never sees a token).
 * Uses Node 22's built-in fetch — no HTTP client dependency.
 */
import { getValidAccessToken } from './google-auth'

const DRIVE_FILES = 'https://www.googleapis.com/drive/v3/files'
const DRIVE_UPLOAD_FILES = 'https://www.googleapis.com/upload/drive/v3/files'
const DRIVE_PERMISSIONS = (fileId: string) =>
  `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}/permissions`

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const GDOC_MIME = 'application/vnd.google-apps.document'

export interface GoogleDoc {
  id: string
  name: string
  modifiedTime: string
  webViewLink: string
  iconLink?: string
}

export interface GoogleFolder {
  id: string
  name: string
}

export interface GooglePermission {
  id: string
  type: 'user' | 'anyone' | 'group' | 'domain'
  role: 'reader' | 'commenter' | 'writer' | 'owner'
  emailAddress?: string
  displayName?: string
}

type DriveResult<T> = { ok: true; data: T } | { ok: false; error: string; status?: number }

async function authedFetch(url: string, init: RequestInit = {}): Promise<DriveResult<Response>> {
  const auth = await getValidAccessToken()
  if (!auth.ok) return { ok: false, error: auth.error }
  const res = await fetch(url, {
    ...init,
    headers: { ...(init.headers ?? {}), Authorization: `Bearer ${auth.accessToken}` },
  })
  if (!res.ok) {
    let detail = ''
    try {
      detail = await res.text()
    } catch {
      /* ignore */
    }
    return { ok: false, error: detail || res.statusText, status: res.status }
  }
  return { ok: true, data: res }
}

/** Files the user's Google Drive account owns/has access to that are native
 *  Google Docs, newest-modified first. */
export async function listGoogleDocs(): Promise<DriveResult<GoogleDoc[]>> {
  const params = new URLSearchParams({
    q: "mimeType='application/vnd.google-apps.document' and trashed=false",
    orderBy: 'modifiedTime desc',
    pageSize: '30',
    fields: 'files(id,name,modifiedTime,webViewLink,iconLink)',
  })
  const result = await authedFetch(`${DRIVE_FILES}?${params.toString()}`)
  if (!result.ok) return result
  const body = (await result.data.json()) as { files: GoogleDoc[] }
  return { ok: true, data: body.files ?? [] }
}

/** Minimal metadata lookup (webViewLink) for a file already known by id —
 *  used to resolve the share link for a document imported earlier this session. */
export async function getGoogleFileMeta(
  fileId: string,
): Promise<DriveResult<{ id: string; name: string; webViewLink: string }>> {
  const params = new URLSearchParams({ fields: 'id,name,webViewLink' })
  const result = await authedFetch(
    `${DRIVE_FILES}/${encodeURIComponent(fileId)}?${params.toString()}`,
  )
  if (!result.ok) return result
  return {
    ok: true,
    data: (await result.data.json()) as { id: string; name: string; webViewLink: string },
  }
}

/** Export a Google Doc to .docx bytes. */
export async function exportGoogleDoc(fileId: string): Promise<DriveResult<ArrayBuffer>> {
  const params = new URLSearchParams({ mimeType: DOCX_MIME })
  const result = await authedFetch(
    `${DRIVE_FILES}/${encodeURIComponent(fileId)}/export?${params.toString()}`,
  )
  if (!result.ok) return result
  return { ok: true, data: await result.data.arrayBuffer() }
}

/** RFC 2387 multipart/related body for Drive's multipart upload, as used by
 *  files.create/files.update. Exported standalone so the body layout can be
 *  unit-tested without a network call. */
export function buildMultipartUploadBody(
  boundary: string,
  metadata: Record<string, unknown>,
  fileBytes: Uint8Array,
  fileMime: string,
): Uint8Array {
  const encoder = new TextEncoder()
  const head = encoder.encode(
    `--${boundary}\r\n` +
      'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
      `${JSON.stringify(metadata)}\r\n` +
      `--${boundary}\r\n` +
      `Content-Type: ${fileMime}\r\n\r\n`,
  )
  const tail = encoder.encode(`\r\n--${boundary}--`)
  const combined = new Uint8Array(head.length + fileBytes.length + tail.length)
  combined.set(head, 0)
  combined.set(fileBytes, head.length)
  combined.set(tail, head.length + fileBytes.length)
  return combined
}

function newBoundary(): string {
  return `genoffice-${Date.now()}-${Math.random().toString(36).slice(2)}`
}

/** Folders under `parentId` (defaults to "root" / My Drive), name-sorted, for
 *  the destination-folder picker. Only folders the user can see are returned
 *  — trashed folders are excluded. */
export async function listFolders(parentId?: string): Promise<DriveResult<GoogleFolder[]>> {
  const parent = parentId || 'root'
  const params = new URLSearchParams({
    q: `mimeType='application/vnd.google-apps.folder' and trashed=false and '${parent}' in parents`,
    orderBy: 'name',
    pageSize: '200',
    fields: 'files(id,name)',
  })
  const result = await authedFetch(`${DRIVE_FILES}?${params.toString()}`)
  if (!result.ok) return result
  const body = (await result.data.json()) as { files: GoogleFolder[] }
  return { ok: true, data: body.files ?? [] }
}

/** Create a new native Google Doc from .docx bytes (Drive converts on upload).
 *  Pass `folderId` to place it directly in a Drive folder instead of My Drive root. */
export async function createGoogleDocFromDocx(
  name: string,
  docxBytes: Uint8Array,
  folderId?: string,
): Promise<DriveResult<{ id: string; webViewLink: string }>> {
  const boundary = newBoundary()
  const metadata: Record<string, unknown> = { name, mimeType: GDOC_MIME }
  if (folderId) metadata.parents = [folderId]
  const body = buildMultipartUploadBody(boundary, metadata, docxBytes, DOCX_MIME)
  const result = await authedFetch(
    `${DRIVE_UPLOAD_FILES}?uploadType=multipart&fields=id,webViewLink`,
    {
      method: 'POST',
      headers: { 'Content-Type': `multipart/related; boundary=${boundary}` },
      body: body as unknown as BodyInit,
    },
  )
  if (!result.ok) return result
  const json = (await result.data.json()) as { id: string; webViewLink: string }
  return { ok: true, data: json }
}

/** Push new .docx content into an existing Google Doc, keeping the same fileId. */
export async function updateGoogleDocFromDocx(
  fileId: string,
  docxBytes: Uint8Array,
): Promise<DriveResult<{ id: string; webViewLink: string }>> {
  const boundary = newBoundary()
  const body = buildMultipartUploadBody(boundary, {}, docxBytes, DOCX_MIME)
  const result = await authedFetch(
    `${DRIVE_UPLOAD_FILES}/${encodeURIComponent(fileId)}?uploadType=multipart&fields=id,webViewLink`,
    {
      method: 'PATCH',
      headers: { 'Content-Type': `multipart/related; boundary=${boundary}` },
      body: body as unknown as BodyInit,
    },
  )
  if (!result.ok) return result
  const json = (await result.data.json()) as { id: string; webViewLink: string }
  return { ok: true, data: json }
}

export async function listPermissions(fileId: string): Promise<DriveResult<GooglePermission[]>> {
  const params = new URLSearchParams({
    fields: 'permissions(id,type,role,emailAddress,displayName)',
  })
  const result = await authedFetch(`${DRIVE_PERMISSIONS(fileId)}?${params.toString()}`)
  if (!result.ok) return result
  const body = (await result.data.json()) as { permissions: GooglePermission[] }
  return { ok: true, data: body.permissions ?? [] }
}

export async function addPermission(
  fileId: string,
  emailAddress: string,
  role: 'reader' | 'commenter' | 'writer',
): Promise<DriveResult<GooglePermission>> {
  const result = await authedFetch(
    `${DRIVE_PERMISSIONS(fileId)}?sendNotificationEmail=true&fields=id,type,role,emailAddress,displayName`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'user', role, emailAddress }),
    },
  )
  if (!result.ok) return result
  return { ok: true, data: (await result.data.json()) as GooglePermission }
}

export async function updatePermission(
  fileId: string,
  permissionId: string,
  role: 'reader' | 'commenter' | 'writer',
): Promise<DriveResult<GooglePermission>> {
  const result = await authedFetch(
    `${DRIVE_PERMISSIONS(fileId)}/${encodeURIComponent(permissionId)}?fields=id,type,role,emailAddress,displayName`,
    {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ role }),
    },
  )
  if (!result.ok) return result
  return { ok: true, data: (await result.data.json()) as GooglePermission }
}

export async function removePermission(
  fileId: string,
  permissionId: string,
): Promise<DriveResult<null>> {
  const result = await authedFetch(
    `${DRIVE_PERMISSIONS(fileId)}/${encodeURIComponent(permissionId)}`,
    { method: 'DELETE' },
  )
  if (!result.ok) return result
  return { ok: true, data: null }
}

/** Build the addParents/removeParents query for files.update, given the file's
 *  current parent ids and the target folder. Exported standalone so the query
 *  construction is unit-testable without a network call. */
export function buildMoveParentsQuery(currentParents: string[], targetFolderId: string): string {
  const params = new URLSearchParams({
    addParents: targetFolderId,
    removeParents: currentParents.join(','),
    fields: 'id,parents',
  })
  return params.toString()
}

/** Move a file to a different Drive folder: fetches its current parents, then
 *  swaps them for `targetFolderId` via files.update addParents/removeParents. */
export async function moveFile(
  fileId: string,
  targetFolderId: string,
): Promise<DriveResult<{ id: string; parents: string[] }>> {
  const metaResult = await authedFetch(
    `${DRIVE_FILES}/${encodeURIComponent(fileId)}?fields=parents`,
  )
  if (!metaResult.ok) return metaResult
  const meta = (await metaResult.data.json()) as { parents?: string[] }
  const currentParents = meta.parents ?? []
  const query = buildMoveParentsQuery(currentParents, targetFolderId)
  const result = await authedFetch(`${DRIVE_FILES}/${encodeURIComponent(fileId)}?${query}`, {
    method: 'PATCH',
  })
  if (!result.ok) return result
  return { ok: true, data: (await result.data.json()) as { id: string; parents: string[] } }
}

/** Build the files.copy request body, given the destination name and optional
 *  target folder. Exported standalone so the body shape is unit-testable
 *  without a network call. */
export function buildCopyFileBody(name: string, folderId?: string): Record<string, unknown> {
  const body: Record<string, unknown> = { name }
  if (folderId) body.parents = [folderId]
  return body
}

/** Drive files.copy: duplicate an existing file (Google Docs File ▸ Make a
 *  copy). `folderId` places the copy directly in a Drive folder instead of
 *  wherever the source file lives. */
export async function copyFile(
  fileId: string,
  name: string,
  folderId?: string,
): Promise<DriveResult<{ id: string; name: string; webViewLink: string; parents?: string[] }>> {
  const result = await authedFetch(
    `${DRIVE_FILES}/${encodeURIComponent(fileId)}/copy?fields=id,name,webViewLink,parents`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(buildCopyFileBody(name, folderId)),
    },
  )
  if (!result.ok) return result
  return {
    ok: true,
    data: (await result.data.json()) as {
      id: string
      name: string
      webViewLink: string
      parents?: string[]
    },
  }
}

/** Rename a Drive file (files.update {name}) — GDocsHeader title commit. */
export async function renameFile(
  fileId: string,
  name: string,
): Promise<DriveResult<{ id: string; name: string }>> {
  const result = await authedFetch(`${DRIVE_FILES}/${encodeURIComponent(fileId)}?fields=id,name`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name }),
  })
  if (!result.ok) return result
  return { ok: true, data: (await result.data.json()) as { id: string; name: string } }
}

/** General access: anyone-with-the-link. Pass role=null to revoke it.
 *  Also strips any 'domain' permission (Workspace's org-wide default sharing
 *  policy attaches one of these to new files automatically) — without this,
 *  "Restricted" silently leaves the file visible to the whole domain. */
export async function setAnyoneAccess(
  fileId: string,
  role: 'reader' | 'writer' | null,
): Promise<DriveResult<null>> {
  if (role === null) {
    const perms = await listPermissions(fileId)
    if (!perms.ok) return perms
    const openAccess = perms.data.filter((p) => p.type === 'anyone' || p.type === 'domain')
    if (openAccess.length === 0) return { ok: true, data: null }
    for (const perm of openAccess) {
      const result = await removePermission(fileId, perm.id)
      if (!result.ok) return result
    }
    return { ok: true, data: null }
  }
  // A file can only ever have one 'anyone' permission — POSTing again when one
  // already exists creates a duplicate instead of changing its role. List
  // first and PATCH the existing one; only POST when none exists yet.
  const perms = await listPermissions(fileId)
  if (!perms.ok) return perms
  const anyone = perms.data.find((p) => p.type === 'anyone')
  const result = anyone
    ? await authedFetch(`${DRIVE_PERMISSIONS(fileId)}/${encodeURIComponent(anyone.id)}?fields=id`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ role }),
      })
    : await authedFetch(`${DRIVE_PERMISSIONS(fileId)}?fields=id`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type: 'anyone', role }),
      })
  if (!result.ok) return result
  return { ok: true, data: null }
}
