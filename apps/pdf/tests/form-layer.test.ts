import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import {
  FormLayer,
  widgetsFromAnnotations,
  type RawFormAnnotation,
} from '../src/renderer/FormLayer'

const rect = [10, 20, 110, 40]

describe('widgetsFromAnnotations', () => {
  it('maps supported AcroForm widgets and ignores unsupported annotations', () => {
    const annots: RawFormAnnotation[] = [
      {
        subtype: 'Widget',
        fieldType: 'Tx',
        fieldName: 'name',
        rect,
        fieldValue: 'Alice',
        required: true,
        maxLen: 12,
        textAlignment: 2,
      },
      {
        subtype: 'Widget',
        fieldType: 'Btn',
        fieldName: 'agree',
        rect,
        fieldValue: 'Yes',
        checkBox: true,
      },
      {
        subtype: 'Widget',
        fieldType: 'Btn',
        fieldName: 'submit',
        rect,
        pushButton: true,
      },
      { subtype: 'Link', fieldName: 'not-a-widget', rect },
    ]

    const widgets = widgetsFromAnnotations(annots)

    expect(widgets).toHaveLength(2)
    expect(widgets[0]).toMatchObject({
      fieldName: 'name',
      kind: 'text',
      value: 'Alice',
      required: true,
      maxLen: 12,
      textAlignment: 'right',
    })
    expect(widgets[1]).toMatchObject({
      fieldName: 'agree',
      kind: 'checkbox',
      checked: true,
    })
  })

  it('preserves text constraints and normalizes invalid maximum lengths', () => {
    const [password, multiline, invalid] = widgetsFromAnnotations([
      {
        subtype: 'Widget',
        fieldType: 'Tx',
        fieldName: 'pin',
        rect,
        password: true,
        comb: true,
        maxLen: 4.8,
        textAlignment: 1,
      },
      {
        subtype: 'Widget',
        fieldType: 'Tx',
        fieldName: 'notes',
        rect,
        multiLine: true,
      },
      {
        subtype: 'Widget',
        fieldType: 'Tx',
        fieldName: 'unlimited',
        rect,
        maxLen: -1,
      },
    ])

    expect(password).toMatchObject({
      password: true,
      comb: true,
      maxLen: 4,
      textAlignment: 'center',
    })
    expect(multiline?.multiLine).toBe(true)
    expect(invalid?.maxLen).toBeNull()
  })

  it('keeps radio export values and choice display labels', () => {
    const widgets = widgetsFromAnnotations([
      {
        subtype: 'Widget',
        fieldType: 'Btn',
        fieldName: 'color',
        rect,
        fieldValue: 'blue',
        radioButton: true,
        buttonValue: 'blue',
      },
      {
        subtype: 'Widget',
        fieldType: 'Ch',
        fieldName: 'country',
        rect,
        fieldValue: ['CN'],
        options: [{ exportValue: 'CN', displayValue: 'China' }, { displayValue: 'Other' }],
      },
    ])

    expect(widgets[0]).toMatchObject({
      kind: 'radio',
      value: 'blue',
      buttonValue: 'blue',
    })
    expect(widgets[1]).toMatchObject({
      kind: 'choice',
      value: 'CN',
      options: [
        { exportValue: 'CN', displayValue: 'China' },
        { exportValue: 'Other', displayValue: 'Other' },
      ],
    })
  })

  it('orders controls top-to-bottom and then left-to-right for native Tab navigation', () => {
    const widgets = widgetsFromAnnotations([
      {
        subtype: 'Widget',
        fieldType: 'Tx',
        fieldName: 'bottom',
        rect: [10, 20, 100, 40],
      },
      {
        subtype: 'Widget',
        fieldType: 'Tx',
        fieldName: 'top-right',
        rect: [150, 200, 250, 220],
      },
      {
        subtype: 'Widget',
        fieldType: 'Tx',
        fieldName: 'top-left',
        rect: [10, 200, 110, 220],
      },
    ])

    expect(widgets.map((widget) => widget.fieldName)).toEqual(['top-left', 'top-right', 'bottom'])
  })

  it('hides the click-to-sign prompt after a visual signature is placed', () => {
    const widgets = widgetsFromAnnotations([
      {
        id: 'sig',
        subtype: 'Widget',
        fieldType: 'Sig',
        fieldName: 'signature',
        rect,
      },
    ])
    const html = renderToStaticMarkup(
      createElement(FormLayer, {
        widgets,
        geom: { pw: 612, ph: 792, rot: 0 },
        scale: 1,
        edits: new Map(),
        signedWidgetIds: new Set(['sig']),
        signatureLabel: 'Click to sign',
        onEdit: () => {},
      }),
    )

    expect(html).toContain('pdf-form-signature is-signed')
    expect(html).toContain('aria-label="Click to sign"')
    expect(html).toContain('aria-disabled="true"')
    expect(html).not.toContain('disabled=""')
    expect(html).not.toContain('>Click to sign</button>')
  })
})
