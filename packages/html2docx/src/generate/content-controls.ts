// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
import { CheckBox, ImportedXmlComponent, TextRun } from 'docx'

function createContentControlFactory(startId = 1000) {
  let nextId = startId
  return function contentControl(node) {
    if (node.controlType === 'checkbox') {
      return new CheckBox({
        alias: node.alias || node.name || 'Checkbox',
        checked: Boolean(node.checked),
      })
    }
    if (node.controlType === 'radio') {
      return new TextRun({ text: node.checked ? '\u25c9' : '\u25cb', noProof: true })
    }
    const sdt = new ImportedXmlComponent('w:sdt')
    const properties = new ImportedXmlComponent('w:sdtPr')
    if (node.alias || node.name) {
      properties.push(new ImportedXmlComponent('w:alias', { 'w:val': node.alias || node.name }))
    }
    if (node.name) {
      properties.push(new ImportedXmlComponent('w:tag', { 'w:val': node.name }))
    }
    properties.push(new ImportedXmlComponent('w:id', { 'w:val': String(nextId++) }))
    if (node.controlType === 'dropdown') {
      const dropdown = new ImportedXmlComponent('w:dropDownList')
      for (const option of node.options || []) {
        dropdown.push(
          new ImportedXmlComponent('w:listItem', {
            'w:displayText': option.label,
            'w:value': option.value,
          }),
        )
      }
      properties.push(dropdown)
    } else {
      properties.push(
        new ImportedXmlComponent('w:text', {
          'w:multiLine': node.controlType === 'multiline' ? '1' : '0',
        }),
      )
    }
    const content = new ImportedXmlComponent('w:sdtContent')
    content.push(
      new TextRun({
        text: node.value || '\u00a0',
        noProof: true,
        // placeholder hints keep their authored light gray
        color: node.valueColor || undefined,
      }),
    )
    sdt.push(properties)
    sdt.push(content)
    return sdt
  }
}

export { createContentControlFactory }
