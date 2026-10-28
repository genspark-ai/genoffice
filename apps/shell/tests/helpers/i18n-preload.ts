// locale dictionaries load lazily now; preload every one so sync t()
// assertions see the same dictionaries the static bundle used to ship
import { LANGS } from '@genoffice/i18n'
import { loadStrings } from '../../src/renderer/src/strings'

await Promise.all(LANGS.map((lang) => loadStrings(lang)))
