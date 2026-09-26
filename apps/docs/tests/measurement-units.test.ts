import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import {
  defaultMeasurementUnit,
  formatLength,
  fromUnit,
  measurementUnit,
  parseLength,
  setMeasurementUnit,
  toUnit,
} from '../src/renderer/units'
import { LengthInput } from '../src/renderer/components/LengthInput'
import { MarginDialog } from '../src/renderer/components/MarginDialog'
import { ParagraphDialog } from '../src/renderer/components/ParagraphDialog'
import { rulerTicks } from '../src/renderer/components/Ruler'
import {
  firstLineFromSpecial,
  pickSpecial,
  specialFromFirstLine,
} from '../src/renderer/components/paragraph-special-indent'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { LocaleProvider, setModuleLang } from '../src/renderer/i18n/locale'

Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })
setModuleLang('en')

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

afterEach(() => {
  setMeasurementUnit(null)
})

function render(element: React.ReactElement) {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(createElement(LocaleProvider, { initial: 'en', children: element })))
  return {
    container,
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

function typeAndBlur(input: HTMLInputElement, text: string) {
  act(() => {
    input.focus()
    input.value = text
    input.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
  })
}

describe('measurement units', () => {
  it('defaults to inches only for the inch countries', () => {
    expect(defaultMeasurementUnit('en-US')).toBe('in')
    expect(defaultMeasurementUnit('en')).toBe('in')
    expect(defaultMeasurementUnit('en-LR')).toBe('in')
    expect(defaultMeasurementUnit('my-MM')).toBe('in')
    expect(defaultMeasurementUnit('en-GB')).toBe('cm')
    expect(defaultMeasurementUnit('zh-CN')).toBe('cm')
    expect(defaultMeasurementUnit('zh-Hant-TW')).toBe('cm')
    expect(defaultMeasurementUnit('de')).toBe('cm')
  })

  it('persists the user choice over the locale default', () => {
    setMeasurementUnit('pt')
    expect(measurementUnit()).toBe('pt')
    expect(localStorage.getItem('aidocs.measurementUnit')).toBe('pt')
    setMeasurementUnit(null)
    expect(localStorage.getItem('aidocs.measurementUnit')).toBeNull()
  })

  it('converts and rounds per unit', () => {
    expect(toUnit(1440, 'in')).toBe(1)
    expect(toUnit(1440, 'cm')).toBe(2.54)
    expect(toUnit(1440, 'mm')).toBe(25.4)
    expect(toUnit(1440, 'pt')).toBe(72)
    expect(toUnit(1440, 'pi')).toBe(6)
    expect(toUnit(1080, 'cm')).toBe(1.91)
    expect(toUnit(1000, 'mm')).toBe(17.6)
    expect(fromUnit(2.54, 'cm')).toBe(1440)
    expect(fromUnit(12, 'pt')).toBe(240)
    expect(fromUnit(3, 'pi')).toBe(720)
  })

  it('formats with the unit suffix like Word', () => {
    expect(formatLength(1440, 'in')).toBe('1"')
    expect(formatLength(720, 'in')).toBe('0.5"')
    expect(formatLength(1440, 'cm')).toBe('2.54 cm')
    expect(formatLength(567, 'cm')).toBe('1 cm')
    expect(formatLength(1440, 'mm')).toBe('25.4 mm')
    expect(formatLength(240, 'pt')).toBe('12 pt')
    expect(formatLength(720, 'pi')).toBe('3 pi')
    expect(formatLength(-360, 'pt')).toBe('-18 pt')
    expect(formatLength(0, 'cm')).toBe('0 cm')
    expect(formatLength(1440, 'cm', '\u5398\u7c73')).toBe('2.54 \u5398\u7c73')
  })

  it('parses any unit suffix, a bare number in the current unit and comma decimals', () => {
    expect(parseLength('2cm', 'in')).toBe(1134)
    expect(parseLength('1"', 'cm')).toBe(1440)
    expect(parseLength('1 in', 'cm')).toBe(1440)
    expect(parseLength('36pt', 'cm')).toBe(720)
    expect(parseLength('3pi', 'cm')).toBe(720)
    expect(parseLength('25.4 mm', 'in')).toBe(1440)
    expect(parseLength('2,54', 'cm')).toBe(1440)
    expect(parseLength('-0.5', 'in')).toBe(-720)
    expect(parseLength('96px', 'cm')).toBe(1440)
    expect(parseLength('2 \u5398\u7c73', 'in', { cm: '\u5398\u7c73' })).toBe(1134)
    expect(parseLength('abc', 'cm')).toBeNull()
    expect(parseLength('2 furlongs', 'cm')).toBeNull()
    expect(parseLength('', 'cm')).toBeNull()
  })
})

describe('LengthInput', () => {
  it('shows the preferred unit, commits twips on blur and keeps untouched values exact', () => {
    setMeasurementUnit('cm')
    const onCommit = vi.fn()
    const { container, unmount } = render(
      createElement(LengthInput, { value: 1080, onCommit, ariaLabel: 'Left' }),
    )
    const input = container.querySelector<HTMLInputElement>('input')!
    expect(input.value).toBe('1.91 cm')
    typeAndBlur(input, '1.91 cm')
    expect(onCommit).not.toHaveBeenCalled()
    typeAndBlur(input, '1"')
    expect(onCommit).toHaveBeenLastCalledWith(1440)
    expect(input.value).toBe('2.54 cm')
    typeAndBlur(input, 'nonsense')
    expect(onCommit).toHaveBeenCalledTimes(1)
    unmount()
  })

  it('clamps to the bounds and steps on the unit grid with the arrow keys', () => {
    setMeasurementUnit('in')
    const onCommit = vi.fn()
    const { container, unmount } = render(
      createElement(LengthInput, { value: 1440, onCommit, min: -720, max: 2880 }),
    )
    const input = container.querySelector<HTMLInputElement>('input')!
    expect(input.value).toBe('1"')
    typeAndBlur(input, '9')
    expect(onCommit).toHaveBeenLastCalledWith(2880)
    expect(input.value).toBe('2"')
    typeAndBlur(input, '-3')
    expect(onCommit).toHaveBeenLastCalledWith(-720)
    act(() => {
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }))
    })
    expect(onCommit).toHaveBeenLastCalledWith(-576)
    expect(input.value).toBe('-0.4"')
    unmount()
  })

  it('does not drift a committed value when Enter is followed by blur', () => {
    setMeasurementUnit('cm')
    const onCommit = vi.fn()
    const { container, unmount } = render(createElement(LengthInput, { value: 0, onCommit }))
    const input = container.querySelector<HTMLInputElement>('input')!
    act(() => {
      input.focus()
      input.value = '0.7 in'
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    })
    // 1008 twips shows as 1.78 cm, which would re-parse to 1009
    expect(input.value).toBe('1.78 cm')
    act(() => input.dispatchEvent(new FocusEvent('focusout', { bubbles: true })))
    expect(onCommit).toHaveBeenCalledTimes(1)
    expect(onCommit).toHaveBeenCalledWith(1008)
    unmount()
  })

  it('commits readable keystrokes live for dialogs, formatting only on blur', () => {
    setMeasurementUnit('cm')
    const onCommit = vi.fn()
    const { container, unmount } = render(
      createElement(LengthInput, { value: 0, onCommit, live: true }),
    )
    const input = container.querySelector<HTMLInputElement>('input')!
    const type = (v: string) =>
      act(() => {
        input.focus()
        input.value = v
        input.dispatchEvent(new InputEvent('input', { bubbles: true }))
      })
    type('0 cm ')
    expect(onCommit).not.toHaveBeenCalled()
    type('1"')
    expect(onCommit).toHaveBeenLastCalledWith(1440)
    type('1" x')
    expect(onCommit).toHaveBeenCalledTimes(1)
    type('2')
    expect(onCommit).toHaveBeenLastCalledWith(1134)
    expect(input.value).toBe('2')
    act(() => input.dispatchEvent(new FocusEvent('focusout', { bubbles: true })))
    expect(onCommit).toHaveBeenCalledTimes(2)
    expect(input.value).toBe('2 cm')
    unmount()
  })

  it('re-renders in the new unit when the preference changes', () => {
    setMeasurementUnit('cm')
    const { container, unmount } = render(
      createElement(LengthInput, { value: 1440, onCommit: () => {} }),
    )
    const input = container.querySelector<HTMLInputElement>('input')!
    expect(input.value).toBe('2.54 cm')
    act(() => setMeasurementUnit('pt'))
    expect(input.value).toBe('72 pt')
    unmount()
  })
})

describe('ruler graduation', () => {
  const labels = (unit: Parameters<typeof rulerTicks>[2]) =>
    rulerTicks(12240, 1440, unit)
      .filter((t) => t.kind === 'num')
      .sort((a, b) => a.pos - b.pos)
      .map((t) => `${t.label}@${Math.round(t.pos)}`)

  it('numbers outward from the left margin per unit', () => {
    expect(labels('in')).toEqual([
      '1@0',
      '1@2880',
      '2@4320',
      '3@5760',
      '4@7200',
      '5@8640',
      '6@10080',
      '7@11520',
    ])
    expect(labels('pt').slice(0, 4)).toEqual(['72@0', '36@720', '36@2160', '72@2880'])
    expect(labels('pi').slice(0, 3)).toEqual(['6@0', '6@2880', '12@4320'])
    expect(labels('cm').slice(0, 5)).toEqual(['2@306', '1@873', '1@2007', '2@2574', '3@3141'])
    expect(labels('mm').slice(0, 3)).toEqual(['20@306', '10@873', '10@2007'])
  })

  it('places eighth-inch minors and half-inch mids for inches', () => {
    const ticks = rulerTicks(2880, 1440, 'in').sort((a, b) => a.pos - b.pos)
    expect(ticks.map((t) => t.kind)).toEqual([
      'num',
      'minor',
      'minor',
      'minor',
      'mid',
      'minor',
      'minor',
      'minor',
      'minor',
      'minor',
      'minor',
      'mid',
      'minor',
      'minor',
      'minor',
      'num',
    ])
  })
})

describe('MarginDialog', () => {
  it('shows the margins in the preferred unit and applies parsed twips', () => {
    setMeasurementUnit('in')
    const onApply = vi.fn()
    const { container, unmount } = render(
      createElement(MarginDialog, {
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        mirror: false,
        pageWidth: 12240,
        pageHeight: 15840,
        onApply,
        onClose: () => {},
      }),
    )
    const inputs = container.querySelectorAll<HTMLInputElement>('input')
    expect([...inputs].map((i) => i.value)).toEqual(['1"', '1"', '1"', '1"'])
    typeAndBlur(inputs[2], '2 cm')
    const ok = [...container.querySelectorAll('button')].find((b) => b.textContent === 'OK')!
    act(() => ok.click())
    expect(onApply).toHaveBeenCalledWith(
      { top: 1440, right: 1440, bottom: 1440, left: 1134 },
      false,
    )
    unmount()
  })
})

describe('Paragraph dialog Special indents', () => {
  it('maps w:ind firstLine to Word’s Special dropdown and back', () => {
    expect(specialFromFirstLine(720)).toEqual({ special: 'firstLine', by: 720 })
    expect(specialFromFirstLine(-360)).toEqual({ special: 'hanging', by: 360 })
    expect(specialFromFirstLine(null)).toEqual({ special: 'none', by: 0 })
    expect(firstLineFromSpecial({ special: 'firstLine', by: 720 })).toBe(720)
    expect(firstLineFromSpecial({ special: 'hanging', by: 360 })).toBe(-360)
    expect(firstLineFromSpecial({ special: 'none', by: 360 })).toBeNull()
    expect(firstLineFromSpecial({ special: 'hanging', by: 0 })).toBeNull()
    expect(pickSpecial({ special: 'none', by: 0 }, 'hanging')).toEqual({
      special: 'hanging',
      by: 720,
    })
    expect(pickSpecial({ special: 'firstLine', by: 300 }, 'hanging')).toEqual({
      special: 'hanging',
      by: 300,
    })
    expect(pickSpecial({ special: 'hanging', by: 300 }, 'none')).toEqual({ special: 'none', by: 0 })
  })

  it('writes a hanging indent and a negative left indent through the dialog', () => {
    setMeasurementUnit('cm')
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0, indentFirstLine: 720 },
            content: [{ type: 'text', text: 'Indented' }],
          },
        ],
      },
    })
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 2, 2)))
    const { container, unmount } = render(
      createElement(ParagraphDialog, { editor, onClose: () => {}, pageWidth: 12240 }),
    )
    const special = container.querySelector<HTMLButtonElement>('.gs-dd-btn[aria-label="Special"]')!
    expect(special.dataset.value).toBe('firstLine')
    const by = container.querySelector<HTMLInputElement>('input[aria-label="By"]')!
    expect(by.value).toBe('1.27 cm')
    act(() => special.click())
    act(() =>
      container.querySelector<HTMLButtonElement>('.gs-dd-item[data-value="hanging"]')!.click(),
    )
    typeAndBlur(container.querySelector<HTMLInputElement>('input[aria-label="By"]')!, '0.5"')
    typeAndBlur(
      container.querySelector<HTMLInputElement>('input[aria-label="Left indent"]')!,
      '-1 cm',
    )
    const ok = [...container.querySelectorAll('button')].find((b) => b.textContent === 'OK')!
    act(() => ok.click())
    const attrs = editor.getAttributes('docParagraph')
    expect(attrs.indentFirstLine).toBe(-720)
    expect(attrs.indentLeft).toBe(-567)
    unmount()
    editor.destroy()
  })
})
