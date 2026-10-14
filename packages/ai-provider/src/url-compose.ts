/**
 * URL composition shared by the registry, the endpoint-profile resolver and
 * the model-list probe. It lives in its own module so all three can read the
 * same rules without importing each other: the request path and the settings
 * probe must build the same address for an endpoint, or the pane would list
 * models from a route the chat turn never uses.
 */

/**
 * A stored base URL must be a bounded http(s) URL. Anything else (file:/javascript:
 * schemes, megabyte strings) would misroute gateway traffic or overflow request
 * builders.
 *
 * A query string is kept — Azure-style bases pin `?api-version=…` and gateways
 * pin a version there — while the fragment is dropped (it is never sent to the
 * server, and leaving it on would truncate every composed endpoint path).
 * Embedded credentials are refused outright: they would end up in request logs
 * and error messages, and the api key field is the supported place for them.
 */
export function normalizeBaseUrl(raw: string | undefined, fallback: string): string {
  const candidate = (raw ?? fallback).trim()
  if (candidate === '' || candidate.length > 2048) {
    throw new Error('Base URL must be a non-empty http(s) URL under 2048 characters')
  }
  let parsed: URL
  try {
    parsed = new URL(candidate)
  } catch {
    throw new Error('Base URL must be a valid http(s) URL')
  }
  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
    throw new Error('Base URL must use http or https')
  }
  if (parsed.username || parsed.password) {
    throw new Error('Base URL must not embed credentials; put the key in the API key field')
  }
  parsed.hash = ''
  return parsed.toString()
}

/** Strip trailing slashes and a trailing /v1 from the path (before any query) */
export function stripTrailingV1(base: string): string {
  const q = base.indexOf('?')
  const path = (q === -1 ? base : base.slice(0, q)).replace(/\/+$/, '').replace(/\/v1$/, '')
  return q === -1 ? path : `${path}${base.slice(q)}`
}

/** Append a path segment before any query string so `?api-version=…` stays last */
export function appendPath(base: string, path: string): string {
  const q = base.indexOf('?')
  return q === -1 ? `${base}${path}` : `${base.slice(0, q)}${path}${base.slice(q)}`
}
