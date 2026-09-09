// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
import { CJK_RE } from './fonts'

function forEachRunText(ir, fn) {
  ;(function walk(nodes) {
    for (const node of nodes) {
      for (const run of node.runs || []) fn(run.text || '')
      if (node.children) walk(node.children)
      if (node.cells) for (const cell of node.cells) walk(cell.children || [])
      if (node.rows) {
        for (const row of node.rows) {
          for (const cell of row.cells) walk(cell.children || (cell.runs ? [cell] : []))
        }
      }
      if (node.items) {
        for (const item of node.items) {
          walk([{ runs: Array.isArray(item) ? item : item.runs }])
        }
      }
    }
  })(ir)
}

function detectCJK(ir) {
  let cjk = 0
  let total = 0
  forEachRunText(ir, (text) => {
    total += text.length
    for (const char of text) if (CJK_RE.test(char)) cjk++
  })
  return total > 0 && cjk / total > 0.05
}

// w:lang eastAsia drives Word's CJK line breaking and proofing. The html
// lang attribute is declared intent; script counting is the fallback for
// the common lang-less AI page.
function eastAsiaLangOf(lang = '', ir = []) {
  const lower = lang.toLowerCase()
  const primary = lower.split('-')[0]
  if (primary === 'ja') return 'ja-JP'
  if (primary === 'ko') return 'ko-KR'
  if (primary === 'zh') {
    return /hant|tw|hk|mo/.test(lower) ? 'zh-TW' : 'zh-CN'
  }
  let kana = 0
  let hangul = 0
  forEachRunText(ir, (text) => {
    for (const char of text) {
      const code = char.codePointAt(0)
      if ((code >= 0x3040 && code <= 0x30ff) || (code >= 0x31f0 && code <= 0x31ff)) kana++
      else if ((code >= 0xac00 && code <= 0xd7af) || (code >= 0x1100 && code <= 0x11ff)) hangul++
    }
  })
  if (hangul > kana && hangul > 10) return 'ko-KR'
  if (kana > 10) return 'ja-JP'
  return 'zh-CN'
}

function bidiLangOf(lang = '') {
  const primary = lang.toLowerCase().split('-')[0]
  return (
    { ar: 'ar-SA', he: 'he-IL', fa: 'fa-IR', ur: 'ur-PK' }[primary] ||
    (lang.includes('-') ? lang : 'ar-SA')
  )
}

export { bidiLangOf, detectCJK, eastAsiaLangOf }
