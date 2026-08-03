/**
 * Rule-management panels say what each rule actually does.
 *
 * - List-type data validation names show the actual options (literal lists)
 *   or the source range address, instead of the fixed "from range" wording.
 * - Rules whose formula references #REF! (deleted cells — they can never
 *   trigger) get a visible ⚠ prefix.
 * - Conditional-formatting "custom formula" items show the formula text.
 *   Univer keeps that describe function module-private, so this is a DOM
 *   pass over the rendered panel, re-applied on panel re-renders.
 */
import { LocaleService } from '@univerjs/core'
import {
  ConditionalFormattingRuleModel,
  type IConditionFormattingRule,
} from '@univerjs/preset-sheets-conditional-formatting'
import { DataValidatorRegistryService } from '@univerjs/preset-sheets-data-validation'

import { t } from './i18n/locale'
import type { UniverRuntime } from './univer-state'

const BROKEN_REF = /#REF!/

interface DvRule {
  formula1?: string
}

/** list / listMultiple: literal options → the options; range reference → the address */
function listDetail(formula1: string | undefined): string | null {
  const f = formula1?.trim()
  if (!f) return null
  if (f.startsWith('=')) return t('appDvListFromRange', { range: f.slice(1) })
  const items = f
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean)
  if (items.length === 0) return null
  return t('appDvListOptions', { items: items.join(', ') })
}

function withBrokenBadge(name: string, ...formulas: Array<string | undefined>): string {
  return formulas.some((f) => f && BROKEN_REF.test(f))
    ? `⚠ ${name} — ${t('appRuleBrokenRef')}`
    : name
}

interface NamedValidator {
  generateRuleName(rule: DvRule): string
}

function enhanceDvRuleNames(runtime: UniverRuntime): () => void {
  const registry = runtime.univer.__getInjector().get(DataValidatorRegistryService)
  const restores: Array<() => void> = []
  for (const type of ['list', 'listMultiple', 'custom', 'decimal', 'whole', 'date', 'textLength']) {
    const validator = registry.getValidatorItem(type) as NamedValidator | undefined
    if (!validator) continue
    const original = validator.generateRuleName.bind(validator)
    validator.generateRuleName = (rule: DvRule & { formula2?: string }) => {
      const detail = type === 'list' || type === 'listMultiple' ? listDetail(rule.formula1) : null
      return withBrokenBadge(detail ?? original(rule), rule.formula1, rule.formula2)
    }
    restores.push(() => {
      validator.generateRuleName = original
    })
  }
  return () => restores.forEach((restore) => restore())
}

function colName(index: number): string {
  let name = ''
  for (let i = index; i >= 0; i = Math.floor(i / 26) - 1)
    name = String.fromCharCode(65 + (i % 26)) + name
  return name
}

function rangeA1(r: {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}): string {
  const start = `${colName(r.startColumn)}${r.startRow + 1}`
  const end = `${colName(r.endColumn)}${r.endRow + 1}`
  return start === end ? start : `${start}:${end}`
}

/**
 * The CF panel's formula items render the bare localized "custom formula"
 * label. Find those describe lines (matched by exact label text, then by the
 * ranges line below them) and append the rule's formula, in DOM order per
 * ranges key so duplicate ranges stay stable.
 */
function enhanceCfPanel(runtime: UniverRuntime): () => void {
  const injector = runtime.univer.__getInjector()
  const model = injector.get(ConditionalFormattingRuleModel)
  const locale = injector.get(LocaleService)

  const augment = (): void => {
    const label = locale.t('sheets-conditional-formatting-ui.ruleType.formula')
    if (!label || label.includes('ruleType')) return
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const sheet = workbook?.getActiveSheet()
    if (!workbook || !sheet) return
    const rules =
      (model.getSubunitRules(workbook.getId(), sheet.getSheetId()) as
        IConditionFormattingRule[] | undefined) ?? []
    const byRanges = new Map<string, string[]>()
    for (const rule of rules) {
      const config = rule.rule as { type: string; subType?: string; value?: string }
      if (config.type !== 'highlightCell' || config.subType !== 'formula') continue
      const key = rule.ranges.map(rangeA1).join(',')
      const formula = String(config.value ?? '').replace(/^=/, '')
      if (!byRanges.has(key)) byRanges.set(key, [])
      byRanges.get(key)!.push(formula)
    }
    if (byRanges.size === 0) return
    for (const div of document.querySelectorAll<HTMLElement>('div')) {
      if (div.textContent !== label || div.dataset.gsRuleDetail) continue
      const rangesText = div.nextElementSibling?.textContent ?? ''
      const queue = byRanges.get(rangesText)
      const formula = queue?.shift()
      if (formula === undefined) continue
      div.dataset.gsRuleDetail = '1'
      div.textContent = BROKEN_REF.test(formula)
        ? `⚠ ${label}: ${formula} — ${t('appRuleBrokenRef')}`
        : `${label}: ${formula}`
      div.title = div.textContent
      div.classList.remove('univer-truncate')
      div.style.whiteSpace = 'normal'
      div.style.overflowWrap = 'anywhere'
    }
  }

  let scheduled = false
  const observer = new MutationObserver(() => {
    if (scheduled) return
    scheduled = true
    requestAnimationFrame(() => {
      scheduled = false
      augment()
    })
  })
  observer.observe(document.body, { childList: true, subtree: true })
  return () => observer.disconnect()
}

export function installRuleDetail(runtime: UniverRuntime): () => void {
  const restoreDv = enhanceDvRuleNames(runtime)
  const disconnectCf = enhanceCfPanel(runtime)
  return () => {
    restoreDv()
    disconnectCf()
  }
}
