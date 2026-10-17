/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { IntegrationsApi } from '../src/shared/integrations-api'
import type { FoundSkill } from '../src/shared/found-skill'
import { FoundSkills } from '../src/renderer/src/FoundSkills'
import { LocaleProvider } from '../src/renderer/src/locale'
import { strings } from '../src/renderer/src/strings'

/** the real English text, so assertions read like what a user sees */
const en = strings.en as Record<string, string>
const t = ((key: string, params?: Record<string, string>) =>
  (en[key] ?? key).replace(/\{(\w+)\}/g, (_, name) => params?.[name] ?? `{${name}}`)) as never

/**
 * The pane that offers the skills already on this machine.
 *
 * The point it has to keep making is the difference between a skill that is
 * ours and a skill that is only offered: one group has no button, the other
 * does, and pressing the button is the only thing that copies anything.
 */

const actEnvironment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

let host: HTMLDivElement
let root: Root
let imported: string[]

const skill = (over: Partial<FoundSkill>): FoundSkill => ({
  name: 'pptx-builder',
  description: 'Build a pptx',
  path: '/home/me/.agents/skills/pptx-builder/SKILL.md',
  source: 'agent',
  agent: 'agents',
  relevance: { relevant: true, formats: ['pptx'] },
  ...over,
})

beforeEach(() => {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
  imported = []
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

function mountApi(list: FoundSkill[], body = 'Do the thing.\n'): void {
  window.aiOfficeIntegrations = {
    listSkills: async () => list,
    skillBody: async (p: string) => {
      if (!p.endsWith('SKILL.md')) throw new Error('unknown skill')
      return body
    },
    importSkill: async (p: string) => {
      imported.push(p)
      return skill({ path: '/save/skills/pptx-builder/SKILL.md', source: 'genoffice' })
    },
  } as unknown as IntegrationsApi
}

async function render(): Promise<void> {
  await act(async () => {
    root.render(createElement(LocaleProvider, { lang: 'en' }, createElement(FoundSkills, { t })))
  })
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
  })
}

async function click(el: Element | null | undefined): Promise<void> {
  expect(el).toBeTruthy()
  await act(async () => {
    ;(el as HTMLElement).click()
    await Promise.resolve()
  })
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
  })
}

const buttonWithText = (text: string): HTMLButtonElement | undefined =>
  [...host.querySelectorAll('button')].find((b) => b.textContent?.trim() === text)

/** the Read disclosure inside the row that names this skill */
const disclosureFor = (name: string): HTMLDetailsElement | undefined => {
  const row = [...host.querySelectorAll('.set-intg-row')].find((r) =>
    r.querySelector('.set-field-label')?.textContent?.includes(name),
  )
  return row?.querySelector('details') ?? undefined
}

describe('FoundSkills', () => {
  it('renders nothing until the scan answers', async () => {
    mountApi([skill({})])
    let resolve: (v: FoundSkill[]) => void = () => {}
    ;(window.aiOfficeIntegrations as IntegrationsApi).listSkills = () =>
      new Promise<FoundSkill[]>((r) => (resolve = r))
    await act(async () => {
      root.render(createElement(FoundSkills, { t }))
    })
    // an empty section that fills in a moment later reads as "none found"
    expect(host.textContent).toBe('')

    await act(async () => {
      resolve([skill({})])
      await Promise.resolve()
    })
    expect(host.textContent).toContain('pptx-builder')
  })

  it('offers a candidate with an Import button and does nothing until it is pressed', async () => {
    mountApi([skill({})])
    await render()
    expect(host.textContent).toContain('Build a pptx')
    expect(host.textContent).not.toContain('Does not mention a format GenOffice opens')
    expect(imported).toEqual([])

    await click(buttonWithText('Import'))
    expect(imported).toEqual(['/home/me/.agents/skills/pptx-builder/SKILL.md'])
    expect(host.textContent).toContain('Imported pptx-builder')
  })

  it('gives a skill in our own folder no button: there is nothing to import', async () => {
    mountApi([skill({ source: 'genoffice', agent: undefined })])
    await render()
    expect(host.textContent).toContain('In GenOffice')
    expect(buttonWithText('Import')).toBeUndefined()
  })

  it('marks a skill that is about another format, without hiding it', async () => {
    mountApi([
      skill({
        description: 'Roll back a Kubernetes release',
        relevance: { relevant: false, formats: [] },
      }),
    ])
    await render()
    expect(host.textContent).toContain('Roll back a Kubernetes release')
    expect(host.textContent).toContain('Does not mention a format GenOffice opens')
  })

  it('says so when there is nothing to show', async () => {
    mountApi([])
    await render()
    expect(host.textContent).toContain('No skills found.')
  })

  it('fetches a body only when its disclosure is opened', async () => {
    mountApi([skill({})])
    const body = vi.fn(async () => 'Do the thing.\n')
    window.aiOfficeIntegrations = {
      ...(window.aiOfficeIntegrations as IntegrationsApi),
      skillBody: body,
    } as unknown as IntegrationsApi
    await render()
    expect(body).not.toHaveBeenCalled()

    const details = disclosureFor('pptx-builder')
    expect(details).toBeTruthy()
    await act(async () => {
      details!.open = true
      details!.dispatchEvent(new Event('toggle'))
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(body).toHaveBeenCalledWith('/home/me/.agents/skills/pptx-builder/SKILL.md')
    expect(host.textContent).toContain('Do the thing.')
  })

  it('survives a preload that predates the feature', async () => {
    // an app update can leave the old bridge in place for one launch
    window.aiOfficeIntegrations = {} as IntegrationsApi
    await render()
    expect(host.textContent).toBe('')
  })
})
