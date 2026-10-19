import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFDocument, PDFName } from 'pdf-lib'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('Form Design authors text, checkbox and signature fields that Done writes as real AcroForm widgets', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-pdf-formdesign-'))
  const pdfPath = join(dir, 'blank.pdf')
  const blank = await PDFDocument.create()
  blank.addPage([612, 792])
  await writeFile(pdfPath, await blank.save({ useObjectStreams: false }))

  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'pdf-form-design',
    openFile: pdfPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://pdf/')
    const pdfPage = page.locator('.pdf-page').first()
    await expect(pdfPage).toBeVisible()

    await page.getByRole('button', { name: 'Form Design', exact: true }).click()
    const layer = pdfPage.locator('.pdf-formdesign-layer')
    const box = await pdfPage.boundingBox()
    if (!box) throw new Error('page box')
    const textTool = page.getByRole('button', { name: 'Text field', exact: true })

    // Click drops a default-size text field, disarms the tool and opens the card; names are generated, never shown
    await textTool.click()
    await page.mouse.click(box.x + 80, box.y + 120)
    await expect(layer.locator('.pdf-formdesign-field.is-text')).toHaveCount(1)
    await expect(textTool).toHaveAttribute('aria-pressed', 'false')
    const props = pdfPage.locator('.pdf-formdesign-props')
    await expect(props).toBeVisible()
    await props.getByLabel('Required').check()
    // Done closes the card and writes the field to the PDF: after the reload it is an existing widget
    await props.getByRole('button', { name: 'Done', exact: true }).click()
    await expect(props).toHaveCount(0)
    await expect(layer.locator('.pdf-formdesign-field.is-text.is-existing')).toHaveCount(1)
    await expect(pdfPage.locator('.pdf-form-layer .pdf-form-input')).toHaveCount(1)
    await expect
      .poll(async () => {
        const form = (await PDFDocument.load(await readFile(pdfPath))).getForm()
        return (
          form.getFields().some((f) => f.getName() === 'text_1') &&
          form.getTextField('text_1').isRequired()
        )
      })
      .toBe(true)

    // A second click on the page without a tool must not add anything
    await page.mouse.click(box.x + 300, box.y + 500)
    await expect(layer.locator('.pdf-formdesign-field')).toHaveCount(1)

    // Cancel on a just-placed field removes it again
    await textTool.click()
    await page.mouse.click(box.x + 80, box.y + 160)
    await expect(layer.locator('.pdf-formdesign-field.is-text')).toHaveCount(2)
    await props.getByRole('button', { name: 'Cancel', exact: true }).click()
    await expect(layer.locator('.pdf-formdesign-field.is-text')).toHaveCount(1)

    // Drag draws a checkbox of the dragged size
    await page.getByRole('button', { name: 'Checkbox', exact: true }).click()
    await page.mouse.move(box.x + 80, box.y + 200)
    await page.mouse.down()
    await page.mouse.move(box.x + 110, box.y + 230, { steps: 5 })
    await page.mouse.up()
    await expect(layer.locator('.pdf-formdesign-field.is-checkbox')).toHaveCount(1)
    await props.getByRole('button', { name: 'Done', exact: true }).click()

    await page.getByRole('button', { name: 'Signature field', exact: true }).click()
    await page.mouse.click(box.x + 80, box.y + 300)
    await expect(layer.locator('.pdf-formdesign-field.is-signature')).toHaveCount(1)
    await props.getByRole('button', { name: 'Done', exact: true }).click()
    await expect(layer.locator('.pdf-formdesign-field.is-existing')).toHaveCount(3)

    // Two consecutive radio buttons join one group even though Done saved the first; a dropdown
    // gets editable options; a date field a format
    const radioTool = page.getByRole('button', { name: 'Radio button', exact: true })
    await radioTool.click()
    await page.mouse.click(box.x + 300, box.y + 200)
    await props.getByRole('button', { name: 'Done', exact: true }).click()
    await radioTool.click()
    await page.mouse.click(box.x + 340, box.y + 200)
    await expect(props.getByLabel('Option value')).toHaveValue('option_2')
    await props.getByRole('button', { name: 'Done', exact: true }).click()
    await expect(layer.locator('.pdf-formdesign-field.is-radio')).toHaveCount(2)

    await page.getByRole('button', { name: 'Dropdown', exact: true }).click()
    await page.mouse.click(box.x + 300, box.y + 260)
    // Options commit when the box loses focus (cleaned: trimmed, no blanks, no duplicates)
    await props.getByLabel('Options (one per line)').fill('CN\nUS\n\nJP\nUS ')
    await props.getByLabel('Options (one per line)').blur()
    await expect(props.getByLabel('Options (one per line)')).toHaveValue('CN\nUS\nJP')
    await props.getByLabel('Default value').selectOption('US')
    await props.getByRole('button', { name: 'Done', exact: true }).click()

    await page.getByRole('button', { name: 'Date field', exact: true }).click()
    await page.mouse.click(box.x + 300, box.y + 320)
    await props.getByLabel('Date format').selectOption('dd/mm/yyyy')
    await props.getByRole('button', { name: 'Done', exact: true }).click()
    await expect(layer.locator('.pdf-formdesign-field.is-existing')).toHaveCount(7)
    // The reload lists the saved widgets before it drops their pending twins and resets undo
    await expect(layer.locator('.pdf-formdesign-field')).toHaveCount(7)

    // Dragging a saved field body moves it; the move stays pending until the next save
    const sig = layer.locator('.pdf-formdesign-field.is-signature')
    const sigBefore = await sig.boundingBox()
    if (!sigBefore) throw new Error('sig box')
    await page.mouse.move(sigBefore.x + sigBefore.width / 2, sigBefore.y + sigBefore.height / 2)
    await page.mouse.down()
    await page.mouse.move(
      sigBefore.x + sigBefore.width / 2 + 120,
      sigBefore.y + sigBefore.height / 2 + 60,
      {
        steps: 8,
      },
    )
    await page.mouse.up()
    const sigAfter = await sig.boundingBox()
    if (!sigAfter) throw new Error('sig box')
    expect(sigAfter.x - sigBefore.x).toBeGreaterThan(100)
    expect(sigAfter.y - sigBefore.y).toBeGreaterThan(40)
    await expect(sig).toHaveClass(/is-selected/)
    await expect(sig.locator('.pdf-formdesign-handle')).toHaveCount(8)
    await page.keyboard.press('Shift+ArrowLeft')
    const sigNudged = await sig.boundingBox()
    if (!sigNudged) throw new Error('sig box')
    expect(sigAfter.x - sigNudged.x).toBeGreaterThan(5)

    // Delete removes the selected field, undo brings it back
    await page.keyboard.press('Delete')
    await expect(layer.locator('.pdf-formdesign-field')).toHaveCount(6)
    await page.keyboard.press('ControlOrMeta+z')
    await expect(layer.locator('.pdf-formdesign-field')).toHaveCount(7)

    // Shift-click selects several fields (no handles), Arrange aligns them, Ctrl+D duplicates
    const sigField = layer.getByTitle('signature_1')
    const textField = layer.getByTitle('text_1')
    await textField.click()
    await sigField.click({ modifiers: ['Shift'] })
    await expect(layer.locator('.pdf-formdesign-field.is-selected')).toHaveCount(2)
    await expect(layer.locator('.pdf-formdesign-handle')).toHaveCount(0)
    await expect(props).toHaveCount(0)
    await page.getByRole('button', { name: 'Arrange', exact: true }).click()
    await page.getByRole('button', { name: 'Align left', exact: true }).click()
    const alignedText = await textField.boundingBox()
    const alignedSig = await sigField.boundingBox()
    expect(Math.abs(alignedText!.x - alignedSig!.x)).toBeLessThan(1.5)
    await page.keyboard.press('ControlOrMeta+d')
    await expect(layer.locator('.pdf-formdesign-field')).toHaveCount(9)
    await expect(layer.locator('.pdf-formdesign-field.is-selected')).toHaveCount(2)
    await page.keyboard.press('ControlOrMeta+z')
    await expect(layer.locator('.pdf-formdesign-field')).toHaveCount(7)

    // Context menu on a field offers Delete; undo restores
    await sigField.click({ button: 'right' })
    const menu = page.locator('.thumb-menu')
    await expect(menu).toBeVisible()
    await menu.getByRole('button', { name: 'Delete field', exact: true }).click()
    await expect(layer.locator('.pdf-formdesign-field')).toHaveCount(6)
    await page.keyboard.press('ControlOrMeta+z')
    await expect(layer.locator('.pdf-formdesign-field')).toHaveCount(7)

    // Marquee over the two radio buttons selects both; a plain click on empty space clears
    const radios = layer.locator('.pdf-formdesign-field.is-radio')
    const r1 = await radios.nth(0).boundingBox()
    const r2 = await radios.nth(1).boundingBox()
    await page.mouse.move(r1!.x - 10, r1!.y - 10)
    await page.mouse.down()
    await page.mouse.move(r2!.x + r2!.width + 10, r2!.y + r2!.height + 10, { steps: 6 })
    await page.mouse.up()
    await expect(layer.locator('.pdf-formdesign-field.is-selected')).toHaveCount(2)
    await page.mouse.click(box.x + 500, box.y + 700)
    await expect(layer.locator('.pdf-formdesign-field.is-selected')).toHaveCount(0)

    await page.screenshot({ path: join(dir, 'form-design.png') })
    await page.keyboard.press('ControlOrMeta+s')

    // The authored fields are real widgets rendered by the fill layer and stay editable in Form Design
    await expect(pdfPage.locator('.pdf-form-layer .pdf-form-input')).toHaveCount(2)
    await expect(pdfPage.locator('.pdf-form-layer .pdf-form-checkbox')).toHaveCount(1)
    await expect(pdfPage.locator('.pdf-form-layer .pdf-form-radio')).toHaveCount(2)
    await expect(pdfPage.locator('.pdf-form-layer .pdf-form-select')).toHaveCount(1)
    await expect(pdfPage.locator('.pdf-form-layer .pdf-form-signature')).toHaveCount(1)
    await expect(layer.locator('.pdf-formdesign-field.is-existing')).toHaveCount(7)

    // Preview lets the saved widgets be filled without leaving Form Design
    const previewBtn = page.getByRole('button', { name: 'Preview', exact: true })
    await previewBtn.click()
    await pdfPage.locator('.pdf-form-layer .pdf-form-input').first().click()
    await expect(previewBtn).toHaveAttribute('aria-pressed', 'true')
    await expect(page.getByRole('button', { name: 'Form Design', exact: true })).toHaveClass(
      /active/,
    )
    await previewBtn.click()

    // The date field rewrites other layouts into its format, rejects non-dates with an inline
    // message and offers a calendar picker
    await page.getByRole('button', { name: 'Fill Form', exact: true }).click()
    const dateInput = pdfPage.locator('.pdf-form-layer .pdf-form-input[placeholder="dd/mm/yyyy"]')
    await dateInput.click()
    await page.keyboard.type('2026-10-05')
    await page.keyboard.press('Enter')
    await expect(dateInput).toHaveValue('05/10/2026')
    await dateInput.click()
    await page.keyboard.press('ControlOrMeta+a')
    await page.keyboard.type('hello')
    await page.keyboard.press('Enter')
    await expect(dateInput).toHaveValue('05/10/2026')
    await expect(pdfPage.locator('.pdf-form-date-error')).toContainText('dd/mm/yyyy')
    await expect(pdfPage.locator('.pdf-form-date-pick')).toHaveCount(1)
    await pdfPage.locator('.pdf-form-date-native').evaluate((el: HTMLInputElement) => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(
        el,
        '2024-02-29',
      )
      el.dispatchEvent(new Event('change', { bubbles: true }))
    })
    await expect(dateInput).toHaveValue('29/02/2024')
    await page.getByRole('button', { name: 'Form Design', exact: true }).click()

    const saved = await PDFDocument.load(await readFile(pdfPath))
    const form = saved.getForm()
    expect(form.getTextField('text_1').isRequired()).toBe(true)
    const savedRect = form.getTextField('text_1').acroField.getWidgets()[0]!.getRectangle()
    expect(
      form
        .getFields()
        .map((f) => f.getName())
        .sort(),
    ).toEqual(['checkbox_1', 'choice_1', 'date_1', 'radio_1', 'signature_1', 'text_1'])
    expect(form.getRadioGroup('radio_1').getOptions()).toEqual(['option_1', 'option_2'])
    expect(form.getDropdown('choice_1').getOptions()).toEqual(['CN', 'US', 'JP'])
    expect(form.getDropdown('choice_1').getSelected()).toEqual(['US'])
    expect(form.getTextField('date_1').acroField.getWidgets()[0]!.dict.has(PDFName.of('AA'))).toBe(
      true,
    )

    // Move the saved text field and delete the saved checkbox, then save again
    const existingText = layer.locator('.pdf-formdesign-field.is-existing[title="text_1"]')
    const textBefore = await existingText.boundingBox()
    if (!textBefore) throw new Error('text box')
    await page.mouse.move(textBefore.x + textBefore.width / 2, textBefore.y + textBefore.height / 2)
    await page.mouse.down()
    await page.mouse.move(
      textBefore.x + textBefore.width / 2,
      textBefore.y + textBefore.height / 2 + 100,
      {
        steps: 8,
      },
    )
    await page.mouse.up()
    await layer.locator('.pdf-formdesign-field.is-existing.is-checkbox').click()
    await page.keyboard.press('Delete')
    await expect(layer.locator('.pdf-formdesign-field')).toHaveCount(6)
    await page.keyboard.press('ControlOrMeta+s')
    await expect(pdfPage.locator('.pdf-form-layer .pdf-form-checkbox')).toHaveCount(0)

    const fieldNames = async () =>
      (await PDFDocument.load(await readFile(pdfPath)))
        .getForm()
        .getFields()
        .map((f) => f.getName())
        .sort()
    await expect
      .poll(fieldNames)
      .toEqual(['choice_1', 'date_1', 'radio_1', 'signature_1', 'text_1'])
    const resaved = await PDFDocument.load(await readFile(pdfPath))
    const moved = resaved.getForm().getTextField('text_1').acroField.getWidgets()[0]!.getRectangle()
    expect(moved.y).toBeLessThan(savedRect.y - 50)
    expect(moved.width).toBeCloseTo(savedRect.width, 0)
    console.log(`screenshot: ${join(dir, 'form-design.png')}`)
  } finally {
    await closeAndSaveVideo(launched, 'pdf-form-design')
  }
})
