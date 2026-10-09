import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFDocument, PDFName } from 'pdf-lib'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('Form Design authors text, checkbox and signature fields that save as real AcroForm widgets', async () => {
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

    // Click drops a default-size text field, disarms the tool and opens the card with the name focused
    await textTool.click()
    await page.mouse.click(box.x + 80, box.y + 120)
    await expect(layer.locator('.pdf-formdesign-field.is-text')).toHaveCount(1)
    await expect(textTool).toHaveAttribute('aria-pressed', 'false')
    const props = pdfPage.locator('.pdf-formdesign-props')
    await expect(props).toBeVisible()
    const nameInput = props.locator('input').first()
    await expect(nameInput).toBeFocused()
    await expect(nameInput).toHaveValue('text_1')
    await nameInput.fill('full_name')
    await props.getByLabel('Required').check()
    // Done closes the card and keeps the field
    await props.getByRole('button', { name: 'Done', exact: true }).click()
    await expect(props).toHaveCount(0)
    await expect(layer.locator('.pdf-formdesign-field.is-text')).toContainText('full_name')

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
    await expect(page.locator('.form-ribbon-progress')).toContainText('3 new fields')

    // Two consecutive radio buttons join one group; a dropdown gets editable options; a date field a format
    const radioTool = page.getByRole('button', { name: 'Radio button', exact: true })
    await radioTool.click()
    await page.mouse.click(box.x + 300, box.y + 200)
    await props.getByRole('button', { name: 'Done', exact: true }).click()
    await radioTool.click()
    await page.mouse.click(box.x + 340, box.y + 200)
    await expect(props.locator('input').first()).toHaveValue('radio_1')
    await expect(props.getByLabel('Option value')).toHaveValue('option_2')
    await props.getByRole('button', { name: 'Done', exact: true }).click()
    await expect(layer.locator('.pdf-formdesign-field.is-radio')).toHaveCount(2)

    await page.getByRole('button', { name: 'Dropdown', exact: true }).click()
    await page.mouse.click(box.x + 300, box.y + 260)
    await props.getByLabel('Options (one per line)').fill('CN\nUS\nJP')
    await props.getByLabel('Default value').selectOption('US')
    await props.getByRole('button', { name: 'Done', exact: true }).click()

    await page.getByRole('button', { name: 'Date field', exact: true }).click()
    await page.mouse.click(box.x + 300, box.y + 320)
    await props.getByLabel('Date format').selectOption('dd/mm/yyyy')
    await props.getByRole('button', { name: 'Done', exact: true }).click()
    await expect(page.locator('.form-ribbon-progress')).toContainText('7 new fields')

    // Dragging a field body moves it; arrow keys nudge it
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

    // Undo under an open property card refreshes its name draft; blur must not reapply the undone name
    await layer.getByTitle('full_name').click()
    const nameAgain = props.locator('input').first()
    await nameAgain.fill('renamed_twice')
    await nameAgain.blur()
    await expect(layer.locator('.pdf-formdesign-field.is-selected')).toContainText('renamed_twice')
    await page.keyboard.press('ControlOrMeta+z')
    await expect(nameAgain).toHaveValue('full_name')
    await nameAgain.focus()
    await nameAgain.blur()
    await expect(layer.locator('.pdf-formdesign-field.is-selected')).toContainText('full_name')

    // Shift-click selects several fields (no handles), Arrange aligns them, Ctrl+D duplicates
    const sigField = layer.getByTitle('signature_1')
    const textField = layer.getByTitle('full_name')
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

    // Preview shows fillable stand-ins for the pending fields
    await page.getByRole('button', { name: 'Preview', exact: true }).click()
    await expect(pdfPage.locator('.pdf-formdesign-preview')).toHaveCount(7)
    await page.getByRole('button', { name: 'Preview', exact: true }).click()
    await expect(pdfPage.locator('.pdf-formdesign-preview')).toHaveCount(0)

    await page.screenshot({ path: join(dir, 'form-design.png') })
    await page.keyboard.press('ControlOrMeta+s')

    // After the save reload the authored fields are real widgets rendered by the fill layer and
    // stay editable in Form Design
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

    const saved = await PDFDocument.load(await readFile(pdfPath))
    const form = saved.getForm()
    expect(form.getTextField('full_name').isRequired()).toBe(true)
    const savedRect = form.getTextField('full_name').acroField.getWidgets()[0]!.getRectangle()
    expect(
      form
        .getFields()
        .map((f) => f.getName())
        .sort(),
    ).toEqual(['checkbox_1', 'choice_1', 'date_1', 'full_name', 'radio_1', 'signature_1'])
    expect(form.getRadioGroup('radio_1').getOptions()).toEqual(['option_1', 'option_2'])
    expect(form.getDropdown('choice_1').getOptions()).toEqual(['CN', 'US', 'JP'])
    expect(form.getDropdown('choice_1').getSelected()).toEqual(['US'])
    expect(form.getTextField('date_1').acroField.getWidgets()[0]!.dict.has(PDFName.of('AA'))).toBe(
      true,
    )

    // Move the saved text field and delete the saved checkbox, then save again
    const existingText = layer.locator('.pdf-formdesign-field.is-existing[title="full_name"]')
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
    await expect(props.locator('input').first()).toBeDisabled()
    await layer.locator('.pdf-formdesign-field.is-existing.is-checkbox').click()
    await page.keyboard.press('Delete')
    await expect(layer.locator('.pdf-formdesign-field')).toHaveCount(6)
    await page.keyboard.press('ControlOrMeta+s')
    await expect(pdfPage.locator('.pdf-form-layer .pdf-form-checkbox')).toHaveCount(0)

    const resaved = await PDFDocument.load(await readFile(pdfPath))
    expect(
      resaved
        .getForm()
        .getFields()
        .map((f) => f.getName())
        .sort(),
    ).toEqual(['choice_1', 'date_1', 'full_name', 'radio_1', 'signature_1'])
    const moved = resaved
      .getForm()
      .getTextField('full_name')
      .acroField.getWidgets()[0]!
      .getRectangle()
    expect(moved.y).toBeLessThan(savedRect.y - 50)
    expect(moved.width).toBeCloseTo(savedRect.width, 0)
    console.log(`screenshot: ${join(dir, 'form-design.png')}`)
  } finally {
    await closeAndSaveVideo(launched, 'pdf-form-design')
  }
})
