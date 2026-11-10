/**
 * word/settings.xml for the converted document: the two East Asian
 * compatibility flags under which Word sets a space touching Hangul half an
 * em wide (w:balanceSingleByteDoubleByteWidth + w:useFELayout; either alone
 * keeps the face's own 0.25-0.36 em space). That is Hangul's own rule: the
 * original PDFs measure 0.5 em per space in Korean faces, and the narrower
 * spaces pulled one or two extra syllables onto every body line.
 * No compatibilityMode is declared, as before the part existed.
 */

import type JSZip from 'jszip'
import { addDocumentPart } from './parts'

const SETTINGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings'
const SETTINGS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'

/** CT_Compat order: balanceSingleByteDoubleByteWidth precedes useFELayout */
const COMPAT = '<w:compat><w:balanceSingleByteDoubleByteWidth/><w:useFELayout/></w:compat>'

/**
 * The document language, Korean: without a w:lang Word (and GenOffice) keep
 * the Japanese line rules, hanging or squeezing a closing bracket into the
 * margin to pull a syllable up (a statute form keeps "…관한 법률」" on one line), which
 * Hangul never does. CT_RPr puts w:lang after w:szCs.
 */
const KOREAN = '<w:lang w:val="ko-KR" w:eastAsia="ko-KR"/>'

async function withKoreanDefaults(zip: JSZip): Promise<void> {
  const styles = await zip.file('word/styles.xml')?.async('string')
  if (!styles || /<w:rPrDefault>.*?<w:lang\b.*?<\/w:rPrDefault>/s.test(styles)) return
  const next = styles.replace(
    /(<w:rPrDefault><w:rPr>.*?)(<\/w:rPr><\/w:rPrDefault>)/s,
    `$1${KOREAN}$2`,
  )
  if (next !== styles) zip.file('word/styles.xml', next)
}

export async function addHangulLayoutSettings(zip: JSZip): Promise<void> {
  if (zip.file('word/settings.xml')) return
  await withKoreanDefaults(zip)
  await addDocumentPart(
    zip,
    'settings.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${COMPAT}</w:settings>`,
    SETTINGS_REL,
    SETTINGS_CT,
  )
}
