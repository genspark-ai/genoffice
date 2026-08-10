import { describe, expect, it } from 'vitest'
import {
  createI18n,
  format,
  htmlLang,
  isLang,
  LANGS,
  macShortcutsToWin,
  normalizeLang,
} from '../src/index'

describe('normalizeLang', () => {
  it('maps zh variants to zh', () => {
    expect(normalizeLang('zh')).toBe('zh')
    expect(normalizeLang('zh-CN')).toBe('zh')
    expect(normalizeLang('zh-Hans-CN')).toBe('zh')
  })

  it('maps traditional-script zh variants to zh-TW', () => {
    expect(normalizeLang('zh-TW')).toBe('zh-TW')
    expect(normalizeLang('ZH-TW')).toBe('zh-TW')
    expect(normalizeLang('zh-Hant')).toBe('zh-TW')
    expect(normalizeLang('zh-Hant-TW')).toBe('zh-TW')
    expect(normalizeLang('zh-HK')).toBe('zh-TW')
    expect(normalizeLang('zh-MO')).toBe('zh-TW')
  })

  it('maps ja/ko variants to ja/ko', () => {
    expect(normalizeLang('ja')).toBe('ja')
    expect(normalizeLang('ja-JP')).toBe('ja')
    expect(normalizeLang('ko')).toBe('ko')
    expect(normalizeLang('KO-KR')).toBe('ko')
  })

  it('maps the newly added languages', () => {
    expect(normalizeLang('fr')).toBe('fr')
    expect(normalizeLang('fr-CA')).toBe('fr')
    expect(normalizeLang('de-DE')).toBe('de')
    expect(normalizeLang('es-419')).toBe('es')
    expect(normalizeLang('th-TH')).toBe('th')
    expect(normalizeLang('id-ID')).toBe('id')
    expect(normalizeLang('in-ID')).toBe('id')
    expect(normalizeLang('ru-RU')).toBe('ru')
    expect(normalizeLang('ar-SA')).toBe('ar')
    expect(normalizeLang('pt-BR')).toBe('pt')
    expect(normalizeLang('pt-PT')).toBe('pt')
    expect(normalizeLang('it-IT')).toBe('it')
    expect(normalizeLang('pl-PL')).toBe('pl')
    expect(normalizeLang('nl-NL')).toBe('nl')
    expect(normalizeLang('ms-MY')).toBe('ms')
    expect(normalizeLang('he-IL')).toBe('he')
    expect(normalizeLang('iw-IL')).toBe('he')
    expect(normalizeLang('hi-IN')).toBe('hi')
  })

  it('maps everything else (and missing) to en', () => {
    expect(normalizeLang('en-US')).toBe('en')
    expect(normalizeLang('sv-SE')).toBe('en')
    expect(normalizeLang('')).toBe('en')
    expect(normalizeLang(undefined)).toBe('en')
    expect(normalizeLang(null)).toBe('en')
  })
})

describe('isLang', () => {
  it('accepts only supported languages', () => {
    expect(isLang('zh')).toBe(true)
    expect(isLang('en')).toBe(true)
    expect(isLang('ja')).toBe(true)
    expect(isLang('ko')).toBe(true)
    expect(isLang('fr')).toBe(true)
    expect(isLang('de')).toBe(true)
    expect(isLang('es')).toBe(true)
    expect(isLang('th')).toBe(true)
    expect(isLang('id')).toBe(true)
    expect(isLang('ru')).toBe(true)
    expect(isLang('ar')).toBe(true)
    expect(isLang('pt')).toBe(true)
    expect(isLang('it')).toBe(true)
    expect(isLang('pl')).toBe(true)
    expect(isLang('nl')).toBe(true)
    expect(isLang('ms')).toBe(true)
    expect(isLang('he')).toBe(true)
    expect(isLang('hi')).toBe(true)
    expect(isLang('zh-TW')).toBe(true)
    expect(isLang('zh-CN')).toBe(false)
    expect(isLang(42)).toBe(false)
  })
})

describe('htmlLang', () => {
  it('maps every supported language to a BCP-47 tag', () => {
    expect(htmlLang('zh')).toBe('zh-CN')
    expect(htmlLang('en')).toBe('en-US')
    expect(htmlLang('ja')).toBe('ja-JP')
    expect(htmlLang('ko')).toBe('ko-KR')
    expect(htmlLang('zh-TW')).toBe('zh-TW')
    for (const lang of LANGS) expect(htmlLang(lang)).toMatch(/^[a-z]{2}-[A-Z]{2}$/)
  })
})

describe('format', () => {
  it('fills placeholders and keeps unknown ones', () => {
    expect(format('已选 {n} 项', { n: 3 })).toBe('已选 3 项')
    expect(format('{a} and {b}', { a: 'x' })).toBe('x and {b}')
    expect(format('no params')).toBe('no params')
  })
})

describe('macShortcutsToWin', () => {
  it('rewrites single-modifier chords', () => {
    expect(macShortcutsToWin('剪切 (⌘X)')).toBe('剪切 (Ctrl+X)')
    expect(macShortcutsToWin('Save (⌘S).')).toBe('Save (Ctrl+S).')
    expect(macShortcutsToWin('⌘1')).toBe('Ctrl+1')
  })

  it('rewrites multi-modifier chords in Ctrl, Alt, Shift order', () => {
    expect(macShortcutsToWin('⇧⌘G')).toBe('Ctrl+Shift+G')
    expect(macShortcutsToWin('⌘⇧V')).toBe('Ctrl+Shift+V')
    expect(macShortcutsToWin('⌥⌘I')).toBe('Ctrl+Alt+I')
    expect(macShortcutsToWin('⌥⇧⌘Z')).toBe('Ctrl+Alt+Shift+Z')
  })

  it('dedupes command and control into one Ctrl', () => {
    expect(macShortcutsToWin('⌃⌘F')).toBe('Ctrl+F')
  })

  it('handles function keys, arrows and named keys', () => {
    expect(macShortcutsToWin('⇧F5')).toBe('Shift+F5')
    expect(macShortcutsToWin('⌘⇧↑')).toBe('Ctrl+Shift+↑')
    expect(macShortcutsToWin('⌘⇧⌫')).toBe('Ctrl+Shift+Backspace')
    expect(macShortcutsToWin('⌫')).toBe('Backspace')
    expect(macShortcutsToWin('⌦')).toBe('Delete')
    expect(macShortcutsToWin('⏎')).toBe('Enter')
    expect(macShortcutsToWin('↩')).toBe('Enter')
    expect(macShortcutsToWin('␣')).toBe('Space')
  })

  it('keeps the Ctrl side of dual-platform listings', () => {
    expect(macShortcutsToWin('⌘/Ctrl+Enter 发送')).toBe('Ctrl+Enter 发送')
    expect(macShortcutsToWin('⌘/Strg+Eingabetaste')).toBe('Strg+Eingabetaste')
  })

  it('rewrites modifier+word combos and bare modifiers', () => {
    expect(macShortcutsToWin('⌘+Click to follow')).toBe('Ctrl+Click to follow')
    expect(macShortcutsToWin('按住 ⌘ 并单击')).toBe('按住 Ctrl 并单击')
  })

  it('keeps a chord embedded before non-Latin text intact', () => {
    expect(macShortcutsToWin('⌘S로 저장')).toBe('Ctrl+S로 저장')
  })

  it('leaves strings without Mac symbols untouched', () => {
    const plain = 'Ctrl+S saves the file'
    expect(macShortcutsToWin(plain)).toBe(plain)
    expect(macShortcutsToWin('文件')).toBe('文件')
  })
})

describe('createI18n', () => {
  const t = createI18n({
    zh: { hello: '你好 {name}', plain: '文件' },
    en: { hello: 'Hello {name}', plain: 'Files' },
    ja: { hello: 'こんにちは {name}', plain: 'ファイル' },
    ko: { hello: '안녕하세요 {name}', plain: '파일' },
    fr: { hello: 'Bonjour {name}', plain: 'Fichiers' },
    de: { hello: 'Hallo {name}', plain: 'Dateien' },
    es: { hello: 'Hola {name}', plain: 'Archivos' },
    th: { hello: 'สวัสดี {name}', plain: 'ไฟล์' },
    id: { hello: 'Halo {name}', plain: 'Berkas' },
    ru: { hello: 'Привет, {name}', plain: 'Файлы' },
    ar: { hello: 'مرحباً {name}', plain: 'الملفات' },
    pt: { hello: 'Olá {name}', plain: 'Arquivos' },
    it: { hello: 'Ciao {name}', plain: 'File' },
    pl: { hello: 'Cześć {name}', plain: 'Pliki' },
    nl: { hello: 'Hallo {name}', plain: 'Bestanden' },
    ms: { hello: 'Helo {name}', plain: 'Fail' },
    he: { hello: 'שלום {name}', plain: 'קבצים' },
    hi: { hello: 'नमस्ते {name}', plain: 'फ़ाइलें' },
    'zh-TW': { hello: '你好 {name}', plain: '檔案' },
  })

  it('translates per language with interpolation', () => {
    expect(t('zh', 'hello', { name: '世界' })).toBe('你好 世界')
    expect(t('en', 'hello', { name: 'world' })).toBe('Hello world')
    expect(t('ja', 'hello', { name: '世界' })).toBe('こんにちは 世界')
    expect(t('ko', 'plain')).toBe('파일')
    expect(t('en', 'plain')).toBe('Files')
    expect(t('fr', 'hello', { name: 'monde' })).toBe('Bonjour monde')
    expect(t('ru', 'plain')).toBe('Файлы')
    expect(t('ar', 'plain')).toBe('الملفات')
    expect(t('pt', 'hello', { name: 'mundo' })).toBe('Olá mundo')
    expect(t('he', 'plain')).toBe('קבצים')
    expect(t('zh-TW', 'plain')).toBe('檔案')
  })
})
