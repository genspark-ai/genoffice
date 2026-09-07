/**
 * Process-level constants for the web-server: port, host, static roots,
 * app allow-list. Mirrors the values previously declared inline in
 * `index.ts` so the rest of the modules don't need to know about the
 * server layout.
 */
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'

const __dirname = fileURLToPath(new URL('.', import.meta.url))
export const ROOT = resolve(__dirname, '../../..')

export const PORT = Number(process.env.PORT) || 8080
export const HOST = process.env.HOST || '0.0.0.0'

export const APPS = ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'shell']
export const STATIC_ROOT = resolve(ROOT, 'apps')

export const WEB_TEMP_ROOT = resolve(process.env.TMPDIR || '/tmp', 'genoffice-web-temp')
