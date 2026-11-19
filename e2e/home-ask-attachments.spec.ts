import { test, expect } from '@playwright/test'
import { createServer, type IncomingMessage, type Server } from 'node:http'
import { mkdtemp, mkdir, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { AddressInfo } from 'node:net'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { launchShell, closeAndSaveVideo, screenshotPath } from './helpers'

/**
 * Home › Ask AI with files attached to the question: a document and an image
 * dropped on the panel become composer chips, the model sees the attachment
 * list and reads the document through read_attachment, the image rides along
 * as vision input, and the question goes through with no starred corpus.
 * The model is a local OpenAI-compatible stub that scripts the two turns.
 */

interface ChatRequest {
  messages: Array<{ role: string; content: unknown; tool_call_id?: string }>
}

const BRIEF_TEXT = 'Quarterly brief: revenue grew twelve percent in the east region.'
// 1x1 PNG
const PNG_BYTES = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  'base64',
)

async function readBody(req: IncomingMessage): Promise<string> {
  const chunks: Buffer[] = []
  for await (const chunk of req) chunks.push(chunk as Buffer)
  return Buffer.concat(chunks).toString('utf8')
}

function startModelStub(requests: ChatRequest[]): Promise<Server> {
  const server = createServer(async (req, res) => {
    if (req.method !== 'POST') {
      res.writeHead(404).end()
      return
    }
    const body = JSON.parse(await readBody(req)) as ChatRequest
    requests.push(body)
    const toolResult = body.messages.find((m) => m.role === 'tool')
    const message = toolResult
      ? {
          role: 'assistant',
          content: `The brief says: ${String(toolResult.content).split('---\n')[1]?.trim()}`,
        }
      : {
          role: 'assistant',
          content: null,
          tool_calls: [
            {
              id: 'call_1',
              type: 'function',
              function: { name: 'read_attachment', arguments: JSON.stringify({ index: 0 }) },
            },
          ],
        }
    res.writeHead(200, { 'content-type': 'application/json' })
    res.end(
      JSON.stringify({ choices: [{ message, finish_reason: toolResult ? 'stop' : 'tool_calls' }] }),
    )
  })
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve(server)))
}

test('files dropped on Ask AI reach the model as attachments', async () => {
  const requests: ChatRequest[] = []
  const server = await startModelStub(requests)
  const port = (server.address() as AddressInfo).port
  const root = await mkdtemp(join(tmpdir(), 'genoffice-e2e-ask-'))
  await mkdir(join(root, 'Documents'))
  const briefPath = join(root, 'brief.md')
  const photoPath = join(root, 'photo.png')
  await writeFile(briefPath, `# Brief\n\n${BRIEF_TEXT}\n`)
  await writeFile(photoPath, PNG_BYTES)
  const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-e2e-'))
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = {
    apiKey: 'stub',
    model: 'stub-model',
    baseUrl: `http://127.0.0.1:${port}/v1`,
  }
  await writeFile(
    join(userDataDir, 'ai-settings.json'),
    JSON.stringify(resolveCustomEndpoints(settings)),
  )
  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    settings: { defaultSaveDir: join(root, 'Documents'), starPrompt: { resolved: true } },
    videoDir: 'home-ask-attachments',
  })
  const { page } = launched
  try {
    await page.getByRole('button', { name: 'Ask AI', exact: true }).click()
    const panel = page.locator('.kb-ask')
    await expect(panel).toBeVisible()

    // A drop with real paths: the files come from a temporary <input type=file>,
    // the only way a test can mint File objects the preload can resolve to paths.
    await page.evaluate(() => {
      const input = document.createElement('input')
      input.type = 'file'
      input.multiple = true
      input.id = 'e2e-files'
      input.style.display = 'none'
      document.body.appendChild(input)
    })
    await page.locator('#e2e-files').setInputFiles([briefPath, photoPath])
    await panel.evaluate((el) => {
      const input = document.getElementById('e2e-files') as HTMLInputElement
      const dt = new DataTransfer()
      for (const f of Array.from(input.files ?? [])) dt.items.add(f)
      for (const type of ['dragenter', 'dragover', 'drop'] as const) {
        el.dispatchEvent(new DragEvent(type, { bubbles: true, cancelable: true, dataTransfer: dt }))
      }
    })
    const chips = panel.locator('.kb-ask-attachments .kb-ask-attachment')
    await expect(chips).toHaveCount(2)
    await expect(chips.nth(0)).toContainText('brief.md')
    await expect(chips.nth(1).locator('img.kb-ask-attachment-thumb')).toBeVisible()
    // the drop never reached Home's drop-to-open bridge: still on the Ask page, no new tab
    await expect(panel).toBeVisible()
    await page.screenshot({ path: screenshotPath('home-ask-attachments-chips') })

    await panel.locator('textarea').fill('What does the brief say?')
    await panel.locator('textarea').press('Enter')

    const answer = panel.locator('.kb-ask-msg.kb-ask-assistant').last()
    await expect(answer).toContainText('revenue grew twelve percent', { timeout: 30_000 })
    // the composer chips moved onto the sent question
    await expect(
      panel.locator('.kb-ask-attachments:not(.kb-ask-attachments-sent) .kb-ask-attachment'),
    ).toHaveCount(0)
    await expect(
      panel.locator('.kb-ask-user .kb-ask-attachments-sent .kb-ask-attachment'),
    ).toHaveCount(2)
    await expect(answer.locator('.kb-ask-tool-note')).toContainText('brief.md')
    await page.screenshot({ path: screenshotPath('home-ask-attachments-answer') })

    expect(requests.length).toBe(2)
    const first = JSON.stringify(requests[0])
    expect(first).toContain('Attachment list')
    expect(first).toContain('brief.md')
    expect(first).toContain(briefPath)
    expect(first).toContain('image_url')
    const userTurn = requests[0]!.messages.find(
      (m) => m.role === 'user' && Array.isArray(m.content),
    )
    expect(JSON.stringify(userTurn)).toContain('data:image/png;base64,')
    const toolMsg = requests[1]!.messages.find((m) => m.role === 'tool')
    expect(String(toolMsg?.content)).toContain(BRIEF_TEXT)
  } finally {
    await closeAndSaveVideo(launched, 'home-ask-attachments')
    server.close()
    await rm(root, { recursive: true, force: true })
  }
})
