// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
import { LevelFormat, LevelSuffix } from 'docx'

function orderedNumberingLevels(format: any = LevelFormat.DECIMAL) {
  return Array.from({ length: 9 }, (_, level) => ({
    level,
    format,
    text: `%${level + 1}.`,
    suffix: LevelSuffix.SPACE,
    start: 1,
    style: {
      paragraph: {
        indent: { left: 360 * (level + 1), hanging: 240 },
      },
    },
  }))
}

function orderedReference(markerType) {
  return (
    {
      'decimal-leading-zero': 'h2d-ol-decimal-zero',
      'lower-alpha': 'h2d-ol-lower-letter',
      'lower-latin': 'h2d-ol-lower-letter',
      'upper-alpha': 'h2d-ol-upper-letter',
      'upper-latin': 'h2d-ol-upper-letter',
      'lower-roman': 'h2d-ol-lower-roman',
      'upper-roman': 'h2d-ol-upper-roman',
    }[markerType] || 'h2d-ol'
  )
}

function bulletNumberingLevels(text) {
  return Array.from({ length: 9 }, (_, level) => ({
    level,
    format: LevelFormat.BULLET,
    text,
    style: {
      paragraph: {
        indent: { left: 360 * (level + 1), hanging: 240 },
      },
    },
  }))
}

function buildNumberingConfig() {
  return [
    { reference: 'h2d-ol', levels: orderedNumberingLevels() },
    {
      reference: 'h2d-ol-decimal-zero',
      levels: orderedNumberingLevels(LevelFormat.DECIMAL_ZERO),
    },
    {
      reference: 'h2d-ol-lower-letter',
      levels: orderedNumberingLevels(LevelFormat.LOWER_LETTER),
    },
    {
      reference: 'h2d-ol-upper-letter',
      levels: orderedNumberingLevels(LevelFormat.UPPER_LETTER),
    },
    {
      reference: 'h2d-ol-lower-roman',
      levels: orderedNumberingLevels(LevelFormat.LOWER_ROMAN),
    },
    {
      reference: 'h2d-ol-upper-roman',
      levels: orderedNumberingLevels(LevelFormat.UPPER_ROMAN),
    },
    { reference: 'h2d-ul', levels: bulletNumberingLevels('\u2022') },
    { reference: 'h2d-ul-circle', levels: bulletNumberingLevels('\u25e6') },
    { reference: 'h2d-ul-square', levels: bulletNumberingLevels('\u25aa') },
  ]
}

export { buildNumberingConfig, orderedReference }
