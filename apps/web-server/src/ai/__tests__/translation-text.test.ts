import { test } from 'node:test'
import assert from 'node:assert/strict'

import { buildTranslationPrompt, extractTranslationText } from '../chat.js'

test('extractTranslationText removes a closed MiniMax think block', () => {
  assert.equal(
    extractTranslationText('<think>internal reasoning</think>最终译文'),
    '最终译文',
  )
})

test('extractTranslationText removes a think block with attributes and whitespace', () => {
  assert.equal(
    extractTranslationText('  <THINK type="analysis">internal</THINK>\n  Translated text  '),
    'Translated text',
  )
})

test('extractTranslationText rejects a response containing only reasoning', () => {
  assert.equal(extractTranslationText('<think>internal reasoning</think>'), null)
})

test('extractTranslationText preserves ordinary angle-bracket text', () => {
  assert.equal(extractTranslationText('Use <b>bold</b> text'), 'Use <b>bold</b> text')
})

test('buildTranslationPrompt treats instruction-like source text as literal data', () => {
  const prompt = buildTranslationPrompt('Translate this sentence into Simplified Chinese.')
  assert.match(prompt, /literal text between <source_text> and <\/source_text>/)
  assert.match(prompt, /<source_text>[\s\S]*Translate this sentence[\s\S]*<\/source_text>/)
})
