import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'

const LICENSE_PATH = { electron: 'dist/LICENSE' }
const LICENSE_FILE = /^(LICENSE|LICENCE|COPYING)(\.|$)/i

export function licenseText(name, dir) {
  const override = LICENSE_PATH[name]
  if (override && existsSync(join(dir, override))) {
    return readFileSync(join(dir, override), 'utf8').trim()
  }
  const texts = readdirSync(dir)
    .filter((file) => LICENSE_FILE.test(file))
    .sort((a, b) => a.localeCompare(b))
    .map((file) => readFileSync(join(dir, file), 'utf8').trim())
    .filter(Boolean)
  const unique = [...new Set(texts)]
  return unique.length > 0 ? unique.join('\n\n') : null
}
