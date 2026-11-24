import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  GRAMMAR_MAX_CHARS,
  checkGrammar,
  grammarBatches,
  grammarParagraphs,
  locateQuote,
  mapGrammarFindings,
  parseGrammarReply,
  type GrammarParagraph,
} from '../src/renderer/ai/grammar'
import {
  applyGrammarSuggestion,
  grammarFindingAt,
  grammarState,
  ignoreGrammarFinding,
  resetIgnoredGrammar,
  setGrammarFindings,
  setGrammarMarksHidden,
} from '../src/renderer/editor/grammar-marks'
import {
  CHECK_GRAMMAR_KEY,
  HIDE_GRAMMAR_KEY,
  checkGrammarEnabled,
  grammarHidden,
  setCheckGrammarPref,
  setGrammarHiddenPref,
} from '../src/renderer/grammar-pref'

const para = (...content: unknown[]) => ({ type: 'docParagraph', content })
const text = (value: string, marks?: unknown[]) => ({ type: 'text', text: value, marks })
const del = { type: 'del', attrs: { id: '1', author: 'a', date: '' } }
const hidden = { type: 'redaction', attrs: { label: 'client' } }

function open(...paragraphs: unknown[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: paragraphs } as never,
  })
}

const item = (p: number, quote: string, occurrence = 0, suggestions = ['fixed']) => ({
  p,
  quote,
  occurrence,
  message: 'wrong',
  suggestions,
})

const plain = (value: string): GrammarParagraph => ({
  text: value,
  segments: [{ offset: 0, length: value.length, pos: 1 }],
  blocked: [],
})

describe('what a grammar check sends', () => {
  it('reads paragraphs in order and leaves out empty ones, deleted text and inline objects', () => {
    const editor = open(
      para(text('They is '), text('gone ', [del]), text('here'), { type: 'hardBreak' }, text('.')),
      para(),
      para(text('Second.')),
    )
    expect(grammarParagraphs(editor.state.doc).map((p) => p.text)).toEqual([
      'They is here.',
      'Second.',
    ])
    editor.destroy()
  })

  it('sends a placeholder for text hidden from AI, never the words', () => {
    const editor = open(para(text('Call '), text('Ada Lovelace', [hidden]), text(' they is in.')))
    const [only] = grammarParagraphs(editor.state.doc)
    expect(only.text).toBe('Call {{client}} they is in.')
    expect(JSON.stringify(only)).not.toContain('Lovelace')
    editor.destroy()
  })

  it('batches paragraphs under the request size and stops at the cap', () => {
    const long = plain('x'.repeat(2500))
    expect(grammarBatches([long, long, long]).batches).toEqual([[0, 1], [2]])
    const many = Array.from({ length: 40 }, () => long)
    const { batches, checked } = grammarBatches(many)
    expect(checked).toBe(GRAMMAR_MAX_CHARS / 2500)
    expect(batches.flat()).toHaveLength(checked)
  })
})

describe('where a quoted span is in the document', () => {
  it('picks the occurrence the model counted, and none past the last', () => {
    const editor = open(para(text('the cat and the dog and the bird')))
    const [p] = grammarParagraphs(editor.state.doc)
    const { doc } = editor.state
    const second = locateQuote(p, 'the', 1)!
    expect(second.from).toBe(13)
    expect(doc.textBetween(second.from, second.to)).toBe('the')
    expect(locateQuote(p, 'the', 2)!.from).toBe(25)
    expect(locateQuote(p, 'the', 3)).toBeNull()
    editor.destroy()
  })

  it('maps a quote across formatting, an inline object and deleted text', () => {
    const editor = open(
      para(
        text('We '),
        text('was', [{ type: 'bold' }]),
        text(' go'),
        { type: 'hardBreak' },
        text('ing '),
        text('not ', [del]),
        text('home'),
      ),
    )
    const [p] = grammarParagraphs(editor.state.doc)
    const range = locateQuote(p, 'was going home', 0)!
    expect(range).toEqual({ from: 4, to: 23 })
    editor.destroy()
  })

  it('drops a quote that is not in the paragraph exactly, an unknown paragraph and an empty quote', () => {
    const paragraphs = [plain('They is here.')]
    const found = mapGrammarFindings(paragraphs, [
      item(0, 'they is'),
      item(0, 'They  is'),
      item(3, 'They is'),
      item(0, ' '),
      item(0, 'They is'),
    ])
    expect(found.map((f) => f.quote)).toEqual(['They is'])
    expect(found[0]).toMatchObject({ from: 1, to: 8, suggestions: ['fixed'] })
  })

  it('drops a quote that touches text hidden from AI or no-proofing text', () => {
    const editor = open(
      para(text('Call '), text('Ada', [hidden]), text(' they is in.')),
      para(text('skip me', [{ type: 'docTextStyle', attrs: { noProof: true } }]), text(' ok is')),
    )
    const paragraphs = grammarParagraphs(editor.state.doc)
    const found = mapGrammarFindings(paragraphs, [
      item(0, '{{client}} they'),
      item(0, 'they is'),
      item(1, 'me ok'),
    ])
    expect(found.map((f) => f.quote)).toEqual(['they is'])
    expect(editor.state.doc.textBetween(found[0].from, found[0].to)).toBe('they is')
    editor.destroy()
  })

  it('leaves misspelled words to the spelling checker, drops ignored findings and overlaps', () => {
    const paragraphs = [plain('Teh dogs barks at at night.')]
    const first = mapGrammarFindings(paragraphs, [
      item(0, 'Teh'),
      item(0, 'dogs barks'),
      item(0, 'barks at'),
      item(0, 'at at'),
    ])
    expect(first.map((f) => f.quote)).toEqual(['Teh', 'dogs barks', 'at at'])
    const again = mapGrammarFindings(
      paragraphs,
      [item(0, 'Teh'), item(0, 'dogs barks'), item(0, 'at at')],
      { isMisspelled: (w) => w === 'Teh', ignored: new Set([first[2].key]) },
    )
    expect(again.map((f) => f.quote)).toEqual(['dogs barks'])
  })
})

describe('the reply of the model', () => {
  it('takes the JSON list out of a fenced or chatty reply and drops malformed items', () => {
    const reply =
      'Here you go:\n```json\n' +
      JSON.stringify([
        {
          p: 0,
          quote: 'They is',
          occurrence: 0,
          message: ' Agreement. ',
          suggestions: ['They are', 'They is', 'They are', 7],
        },
        { p: '0', quote: 'x', message: 'm' },
        { p: 0, quote: 'no message' },
        { p: 1, quote: 'a a', occurrence: -2, message: 'Repeated word', suggestions: 'a' },
        'text',
      ]) +
      '\n```'
    expect(parseGrammarReply(reply)).toEqual([
      { p: 0, quote: 'They is', occurrence: 0, message: 'Agreement.', suggestions: ['They are'] },
      { p: 1, quote: 'a a', occurrence: 0, message: 'Repeated word', suggestions: [] },
    ])
    expect(parseGrammarReply('I found nothing wrong.')).toEqual([])
    expect(parseGrammarReply('[{"p": 0,')).toEqual([])
  })

  it('checks batch by batch, ignores paragraphs a request did not carry and reports a failure', async () => {
    const editor = open(para(text('They is here.')), para(text('It are fine.')))
    const chat = vi.fn(async (_system: string, user: string) => ({
      ok: true,
      content: JSON.stringify([
        { p: 0, quote: 'They is', message: 'Agreement', suggestions: ['They are'] },
        { p: 9, quote: 'They is', message: 'Agreement' },
        { p: 1, quote: 'It are', message: 'Agreement', suggestions: ['It is'] },
      ]),
      user,
    }))
    const result = await checkGrammar(editor.state.doc, { lang: 'en', chat })
    expect(chat).toHaveBeenCalledOnce()
    expect(JSON.parse(chat.mock.calls[0][1])).toEqual([
      { p: 0, text: 'They is here.' },
      { p: 1, text: 'It are fine.' },
    ])
    expect(result).toMatchObject({ checked: 2, total: 2 })
    expect(result.error).toBeUndefined()
    expect(result.findings.map((f) => [f.from, f.quote])).toEqual([
      [1, 'They is'],
      [16, 'It are'],
    ])
    const failed = await checkGrammar(editor.state.doc, {
      lang: 'en',
      chat: async () => ({ ok: false, error: 'No API key configured for openai' }),
    })
    expect(failed).toMatchObject({ findings: [], error: 'No API key configured for openai' })
    editor.destroy()
  })
})

describe('grammar marks on the page', () => {
  afterEach(() => localStorage.removeItem(HIDE_GRAMMAR_KEY))

  const checked = (editor: Editor, ...quotes: Array<[number, string, number?]>) => {
    const paragraphs = grammarParagraphs(editor.state.doc)
    setGrammarFindings(
      editor,
      mapGrammarFindings(
        paragraphs,
        quotes.map(([p, quote, occurrence]) => item(p, quote, occurrence)),
      ),
    )
    return grammarState(editor.state)!.findings
  }

  it('keeps marks out of the document and out of undo', () => {
    const editor = open(para(text('They is here.')))
    const before = editor.getJSON()
    const [finding] = checked(editor, [0, 'They is'])
    expect(finding).toMatchObject({ from: 1, to: 8 })
    expect(editor.getJSON()).toEqual(before)
    expect(editor.can().undo()).toBe(false)
    editor.destroy()
  })

  it('drops the marks of an edited paragraph and moves the marks of the others', () => {
    const editor = open(para(text('They is here.')), para(text('It are fine and it are.')))
    checked(editor, [0, 'They is'], [1, 'It are'], [1, 'it are'])
    editor.commands.insertContentAt(1, 'Oh, ')
    const left = grammarState(editor.state)!.findings
    expect(left.map((f) => f.quote)).toEqual(['It are', 'it are'])
    expect(editor.state.doc.textBetween(left[0].from, left[0].to)).toBe('It are')
    // formatting is not an edit of the words
    editor.chain().setTextSelection({ from: left[0].from, to: left[0].to }).setMark('bold').run()
    expect(grammarState(editor.state)!.findings).toHaveLength(2)
    editor.destroy()
  })

  it('accepting a suggestion is one undo step, keeps an inline object and the other marks of the paragraph', () => {
    const editor = open(
      para(text('We was go'), { type: 'hardBreak' }, text('ing and it are late.')),
    )
    const [first] = checked(editor, [0, 'was going'], [0, 'it are'])
    expect(applyGrammarSuggestion(editor, first.id, 'were going')).toBe(true)
    expect(editor.state.doc.textContent).toBe('We were going and it are late.')
    let breaks = 0
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'hardBreak') breaks++
    })
    expect(breaks).toBe(1)
    const rest = grammarState(editor.state)!.findings
    expect(rest.map((f) => f.quote)).toEqual(['it are'])
    expect(editor.state.doc.textBetween(rest[0].from, rest[0].to)).toBe('it are')
    editor.commands.undo()
    expect(editor.state.doc.textContent).toBe('We was going and it are late.')
    editor.destroy()
  })

  it('records an accepted suggestion as a tracked change while Track Changes is on', () => {
    const editor = open(para(text('They is here.')))
    editor.storage.trackChanges.enabled = true
    const [finding] = checked(editor, [0, 'They is'])
    expect(applyGrammarSuggestion(editor, finding.id, 'They are')).toBe(true)
    const marks = new Set<string>()
    editor.state.doc.descendants((node) => node.marks.forEach((m) => marks.add(m.type.name)))
    expect(marks.has('ins')).toBe(true)
    expect(marks.has('del')).toBe(true)
    expect(grammarState(editor.state)!.findings).toHaveLength(0)
    editor.destroy()
  })

  it('does not replace in a document that cannot be edited, or with a stale finding', () => {
    const editor = open(para(text('They is here.')))
    const [finding] = checked(editor, [0, 'They is'])
    editor.setEditable(false)
    expect(applyGrammarSuggestion(editor, finding.id, 'They are')).toBe(false)
    expect(editor.state.doc.textContent).toBe('They is here.')
    expect(grammarFindingAt(editor.state, 3)?.id).toBe(finding.id)
    editor.setEditable(true)
    // findings of a document that moved on while the model answered are not taken
    setGrammarFindings(editor, [{ ...finding, id: 'late', from: 3, to: 10 }])
    expect(grammarState(editor.state)!.findings).toHaveLength(0)
    // nor ones past the end of a document that shrank
    setGrammarFindings(editor, [{ ...finding, id: 'gone', from: 400, to: 407 }])
    expect(grammarState(editor.state)!.findings).toHaveLength(0)
    editor.destroy()
  })

  it('Hide Grammar Errors keeps the findings and stops offering them; the choice is remembered', () => {
    expect(grammarHidden()).toBe(false)
    const editor = open(para(text('They is here.')))
    const [finding] = checked(editor, [0, 'They is'])
    expect(grammarFindingAt(editor.state, 1)?.id).toBe(finding.id)
    expect(grammarFindingAt(editor.state, 8)?.id).toBe(finding.id)
    expect(grammarFindingAt(editor.state, 9)).toBeNull()
    setGrammarMarksHidden(editor, true)
    expect(grammarFindingAt(editor.state, 3)).toBeNull()
    expect(grammarState(editor.state)!.findings).toHaveLength(1)
    setGrammarMarksHidden(editor, false)
    expect(grammarFindingAt(editor.state, 3)?.id).toBe(finding.id)
    editor.destroy()

    setGrammarHiddenPref(true)
    expect(grammarHidden()).toBe(true)
    const reopened = open(para(text('They is here.')))
    expect(grammarState(reopened.state)!.hidden).toBe(true)
    reopened.destroy()
  })

  it('Ignore Once keeps the finding out of the next check until the ignored ones are reset', () => {
    const editor = open(para(text('They is here.')))
    const [finding] = checked(editor, [0, 'They is'])
    ignoreGrammarFinding(editor, finding.id)
    expect(grammarState(editor.state)!.findings).toHaveLength(0)
    const recheck = () =>
      mapGrammarFindings(grammarParagraphs(editor.state.doc), [item(0, 'They is')], {
        ignored: grammarState(editor.state)!.ignored,
      })
    expect(recheck()).toHaveLength(0)
    resetIgnoredGrammar(editor)
    expect(recheck()).toHaveLength(1)
    editor.destroy()
  })
})

describe('Check grammar', () => {
  afterEach(() => localStorage.removeItem(CHECK_GRAMMAR_KEY))

  it('is off until ticked, and stays as it was left', () => {
    expect(checkGrammarEnabled()).toBe(false)
    setCheckGrammarPref(true)
    expect(checkGrammarEnabled()).toBe(true)
    setCheckGrammarPref(false)
    expect(checkGrammarEnabled()).toBe(false)
  })
})
