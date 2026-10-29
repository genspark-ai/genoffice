import { test, expect } from '@playwright/test'
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * The composer's slash palette and the settings center's skill editor
 * (genoffice: skills-as-commands).
 *
 * A skill installed under `<userData>/skills/<name>/SKILL.md` must surface in
 * the AI panel when the user types `/`, and picking it must attach the skill
 * chip while stripping the `/query` token from the draft. The settings
 * center's "New skill" editor must write that same file shape, so what the UI
 * creates is what the palette finds.
 */

const SKILL = `---
name: e2e-tables
description: Build spreadsheet-style tables in the document.
---

Always add a header row.
`

async function seedSkill(userDataDir: string): Promise<void> {
  const dir = join(userDataDir, 'skills', 'e2e-tables')
  await mkdir(dir, { recursive: true })
  await writeFile(join(dir, 'SKILL.md'), SKILL)
}

test('slash palette: typing / lists installed skills, Enter attaches and strips the token', async () => {
  test.setTimeout(300_000)
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-slash-'))
  await seedSkill(userDataDir)
  const launched = await launchShell({ userDataDir, onboardingSeen: true, videoDir: 'ai-slash' })
  try {
    await expect(launched.page.locator('.quick-card').first()).toContainText('AI Docs')
    await launched.page.locator('.quick-card').first().click()
    const editor = await waitForPageWithUrl(launched.app, '://docs/')
    const textarea = editor.locator('.ai-composer textarea')
    await expect(textarea).toBeVisible({ timeout: 30_000 })

    await textarea.click()
    await textarea.pressSequentially('/tab')
    const palette = editor.locator('[role=listbox]')
    await expect(palette).toBeVisible()
    const option = palette.locator('[role=option]', { hasText: 'e2e-tables' })
    await expect(option).toBeVisible()

    await textarea.press('Enter')
    await expect(palette).toBeHidden()
    await expect(textarea).toHaveValue('')
    await expect(editor.locator('.ai-skill-chip').filter({ hasText: 'e2e-tables' })).toBeVisible()

    // typing a query that matches nothing shows the no-match state, Escape closes
    await textarea.pressSequentially('/zzzzz')
    await expect(palette).toBeVisible()
    await expect(palette).toContainText('No matching skills')
    await textarea.press('Escape')
    await expect(palette).toBeHidden()
  } finally {
    await closeAndSaveVideo(launched, 'ai-slash')
  }
})

test('settings center: New skill writes a SKILL.md the palette can find', async () => {
  test.setTimeout(300_000)
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-skill-new-'))
  const launched = await launchShell({ userDataDir, onboardingSeen: true, videoDir: 'skill-new' })
  try {
    const page = launched.page
    await page.getByRole('button', { name: 'Settings', exact: true }).click()
    await page.locator('.set-nav-item').filter({ hasText: 'Skills' }).click()
    await page.getByRole('button', { name: 'New skill', exact: true }).click()

    await page.locator('#skill-doc-name').fill('e2e-made')
    await page.locator('#skill-doc-desc').fill('Built by the e2e test.')
    await page.locator('#skill-doc-body').fill('1. Do the thing.\n2. Stop.')
    await page.locator('.set-btn.primary').click()

    await expect(
      page.locator('.set-intg-skill .set-field-label').filter({ hasText: 'e2e-made' }),
    ).toBeVisible()

    const onDisk = await readFile(join(userDataDir, 'skills', 'e2e-made', 'SKILL.md'), 'utf8')
    expect(onDisk).toContain('name: e2e-made')
    expect(onDisk).toContain('description: Built by the e2e test.')
    expect(onDisk).toContain('1. Do the thing.')
  } finally {
    await closeAndSaveVideo(launched, 'skill-new')
  }
})
