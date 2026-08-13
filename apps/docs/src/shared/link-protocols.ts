/**
 * Protocol allowlist for hyperlinks clicked inside Docs documents.
 *
 * Extends the suite-wide default (http/https) with `aof-review:` so footage
 * log links ("Artisan's Docs" -> AOF Footage Review deep links) hand off to
 * the registered protocol handler. Everything else (file:, javascript:,
 * arbitrary app schemes) stays rejected by `safeExternalUrl`.
 */
export const DOCS_LINK_PROTOCOLS: readonly string[] = ['http:', 'https:', 'aof-review:']
