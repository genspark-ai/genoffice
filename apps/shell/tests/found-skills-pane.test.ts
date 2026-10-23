/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { IntegrationsApi } from '../src/shared/integrations-api'
import type { FoundSkill, RecommendedSkill } from '../src/shared/found-skill'
import { FoundSkills } from '../src/renderer/src/FoundSkills'
import { LocaleProvider } from '../src/renderer/src/locale'
import { strings } from '../src/renderer/src/strings'

/** the real English text, so assertions read like what a user sees */
const en = strings.en as Record<string, string>
const t = ((key: string, params?: Record<string, string>) =>
  (en[key] ?? key).replace(/\{(\w+)\}/g, (_, name) => params?.[name] ?? `{${name}}`)) as never

/**
 * The pane that lists the skills installed into GenOffice and lets the user
 * import more from a folder or a zip. The only thing that copies anything is
 * pressing one of those two buttons; the list itself is read-only.
 */

const actEnvironment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

let host: HTMLDivElement
let root: Root
let imported: string[]

const skill = (over: Partial<FoundSkill>): FoundSkill => ({
  name: 'pptx-builder',
  description: 'Build a pptx',
  path: '/app/skills/pptx-builder/SKILL.md',
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

function mountApi(
  list: FoundSkill[],
  body = 'Do the thing.\n',
  recommended: RecommendedSkill[] = [],
): void {
  window.aiOfficeIntegrations = {
    listSkills: async () => list,
    listRecommendedSkills: async () => recommended,
    installRecommendedSkill: async (name: string) => {
      imported.push(`recommended:${name}`)
      return skill({ name, path: `/app/skills/${name}/SKILL.md` })
    },
    skillBody: async (p: string) => {
      if (!p.endsWith('SKILL.md')) throw new Error('unknown skill')
      return body
    },
    importSkillFolder: async (title: string) => {
      imported.push(`folder:${title}`)
      return skill({ name: 'from-folder', path: '/app/skills/from-folder/SKILL.md' })
    },
    importSkillZip: async (title: string) => {
      imported.push(`zip:${title}`)
      return null
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
  it('renders nothing until the listing answers', async () => {
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

  it('imports from a folder only when the button is pressed, then relists', async () => {
    mountApi([skill({})])
    await render()
    expect(imported).toEqual([])
    ;(window.aiOfficeIntegrations as IntegrationsApi).listSkills = async () => [
      skill({}),
      skill({ name: 'from-folder', path: '/app/skills/from-folder/SKILL.md' }),
    ]

    await click(buttonWithText('Import folder…'))
    expect(imported).toEqual(['folder:Choose a skill folder'])
    expect(host.textContent).toContain('Imported from-folder')
    expect(host.textContent).toContain('from-folder')
  })

  it('says nothing when the zip dialog is cancelled', async () => {
    mountApi([skill({})])
    await render()
    await click(buttonWithText('Import skill package (zip / .md)…'))
    expect(imported).toEqual(['zip:Choose a skill zip or .md'])
    expect(host.textContent).not.toContain('Imported')
    expect(host.textContent).not.toContain('Import failed')
  })

  it('shows the reason when an import is refused', async () => {
    mountApi([])
    ;(window.aiOfficeIntegrations as IntegrationsApi).importSkillZip = async () => {
      throw new Error('the archive contains a symlink, and imports refuse those')
    }
    await render()
    await click(buttonWithText('Import skill package (zip / .md)…'))
    expect(host.textContent).toContain('Import failed: the archive contains a symlink')
  })

  it('shows a skill about another format like any other', async () => {
    mountApi([
      skill({
        description: 'Roll back a Kubernetes release',
        relevance: { relevant: false, formats: [] },
      }),
    ])
    await render()
    expect(host.textContent).toContain('Roll back a Kubernetes release')
    expect(host.textContent).not.toContain('GenOffice opens')
  })

  it('says so when there is nothing to show', async () => {
    mountApi([])
    await render()
    expect(host.textContent).toContain('No skills installed yet.')
    expect(buttonWithText('Import folder…')).toBeTruthy()
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
    expect(body).toHaveBeenCalledWith('/app/skills/pptx-builder/SKILL.md')
    expect(host.textContent).toContain('Do the thing.')
  })

  it('offers shipped skills with Install, Update or an installed tag, and installs by name', async () => {
    mountApi([], 'x', [
      {
        ...skill({ name: 'fresh', path: '/res/recommended/fresh/SKILL.md' }),
        category: 'finance',
        version: '1.0.0',
        state: 'missing',
      },
      {
        ...skill({ name: 'old', path: '/res/recommended/old/SKILL.md' }),
        category: 'office',
        version: '1.1.0',
        state: 'outdated',
      },
      {
        ...skill({ name: 'same', path: '/res/recommended/same/SKILL.md' }),
        category: 'made-up',
        version: '1.0.0',
        state: 'installed',
      },
    ])
    await render()
    expect(host.textContent).toContain('Recommended')
    // one tab per category in the fixed order; an unknown category lands under Other
    const tabs = () => [...host.querySelectorAll<HTMLButtonElement>('.set-intg-tab')]
    expect(tabs().map((e) => e.textContent)).toEqual(['Office', 'Finance', 'Other'])
    // the first tab is open: only its skill shows
    expect(buttonWithText('Update')).toBeTruthy()
    expect(buttonWithText('Install')).toBeFalsy()
    await click(tabs()[2])
    expect(host.textContent).toContain('Installed')
    await click(tabs()[1])
    expect(buttonWithText('Install')).toBeTruthy()
    await click(buttonWithText('Install'))
    expect(imported).toEqual(['recommended:fresh'])
    expect(host.textContent).toContain('Imported fresh')
  })

  it('keeps installed and recommended skills on separate sections', async () => {
    mountApi([skill({ name: 'mine-only' })], 'x', [
      {
        ...skill({ name: 'offered', path: '/res/recommended/offered/SKILL.md' }),
        category: 'office',
        version: '1.0.0',
        state: 'missing',
      },
    ])
    await render()
    const sections = () => [...host.querySelectorAll<HTMLButtonElement>('.set-intg-seg')]
    expect(sections().map((e) => e.textContent)).toEqual(['My skills (1)', 'Recommended'])
    // with something installed the list opens on it; the catalog is one click away
    expect(host.textContent).toContain('mine-only')
    expect(host.textContent).not.toContain('offered')
    await click(sections()[1])
    expect(host.textContent).toContain('offered')
    expect(host.textContent).not.toContain('mine-only')
  })

  it('opens on the catalog when nothing is installed and stays there after the first install', async () => {
    const offered = {
      ...skill({ name: 'offered', path: '/res/recommended/offered/SKILL.md' }),
      category: 'office' as const,
      version: '1.0.0',
      state: 'missing' as const,
    }
    mountApi([], 'x', [offered])
    await render()
    const sections = () => [...host.querySelectorAll<HTMLButtonElement>('.set-intg-seg')]
    expect(sections()[1]?.getAttribute('aria-selected')).toBe('true')
    ;(window.aiOfficeIntegrations as IntegrationsApi).listSkills = async () => [
      skill({ name: 'offered', path: '/app/skills/offered/SKILL.md' }),
    ]
    await click(buttonWithText('Install')!)
    expect(imported).toEqual(['recommended:offered'])
    // the relist found one installed skill; the catalog stays in view
    expect(sections().map((e) => e.textContent)).toEqual(['My skills (1)', 'Recommended'])
    expect(sections()[1]?.getAttribute('aria-selected')).toBe('true')
  })

  it('survives a preload that predates the feature', async () => {
    // an app update can leave the old bridge in place for one launch
    window.aiOfficeIntegrations = {} as IntegrationsApi
    await render()
    expect(host.textContent).toBe('')
  })
})
