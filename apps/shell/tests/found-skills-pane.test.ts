/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { HomeApi } from '../src/shared/home-api'
import type { IntegrationsApi } from '../src/shared/integrations-api'
import type { ForeignSkillCandidate, FoundSkill, RecommendedSkill } from '../src/shared/found-skill'
import { FoundSkills } from '../src/renderer/src/FoundSkills'
import { LocaleProvider } from '../src/renderer/src/locale'
import { en as enStrings } from '../src/renderer/src/strings/en'

/** the real English text, so assertions read like what a user sees */
const en = enStrings as Record<string, string>
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

/** one prompt file another agent product keeps on this machine */
const foreign = (over: Partial<ForeignSkillCandidate>): ForeignSkillCandidate => ({
  name: 'rollback-release',
  description: 'Roll back a release',
  source: 'claude',
  sourcePath: '/home/.claude/commands/rollback-release.md',
  alreadyImported: false,
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

/** a click that also drains the promise chains it starts: the foreign importer
 *  imports sequentially, one microtask per row, so a shallow flush would
 *  assert against a half-finished run */
async function clickDeep(el: Element | null | undefined): Promise<void> {
  expect(el).toBeTruthy()
  await act(async () => {
    ;(el as HTMLElement).click()
    for (let i = 0; i < 30; i++) await Promise.resolve()
  })
}

/** the importer's checkbox row that names this candidate */
const foreignBox = (name: string): HTMLInputElement | null | undefined =>
  [...host.querySelectorAll('label.set-intg-foreign-row')]
    .find((r) => r.querySelector('.set-intg-foreign-name')?.textContent === name)
    ?.querySelector('input') ?? undefined

/** one editor/creator field by its input id */
const field = (id: string): HTMLInputElement | HTMLTextAreaElement | null =>
  host.querySelector<HTMLInputElement | HTMLTextAreaElement>(`#${id}`)

/** type into a controlled field through the native setter, so React's own
 *  value tracker registers the change and onChange fires */
function typeInto(el: HTMLInputElement | HTMLTextAreaElement | null, value: string): void {
  expect(el).toBeTruthy()
  const proto =
    el!.tagName === 'TEXTAREA'
      ? window.HTMLTextAreaElement.prototype
      : window.HTMLInputElement.prototype
  Object.getOwnPropertyDescriptor(proto, 'value')?.set?.call(el, value)
  el!.dispatchEvent(new Event('input', { bubbles: true }))
}

/** the aiOffice half the editor and creator talk to (raw read + save + the
 *  model gate); pane tests replace it wholesale like they do aiOfficeIntegrations */
function mountHomeApi(over: Partial<HomeApi>): void {
  window.aiOffice = {
    aiGetSettings: async () => ({}) as never,
    aiGskStatus: async () => ({ loggedIn: false }),
    ...over,
  } as unknown as HomeApi
}

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

    await click(buttonWithText('Import'))
    await click(buttonWithText('Import folder…'))
    expect(imported).toEqual(['folder:Choose a skill folder'])
    expect(host.textContent).toContain('Imported from-folder')
    expect(host.textContent).toContain('from-folder')
  })

  it('says nothing when the zip dialog is cancelled', async () => {
    mountApi([skill({})])
    await render()
    await click(buttonWithText('Import'))
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
    await click(buttonWithText('Import'))
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
    await click(buttonWithText('Import'))
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

  it('creates a skill through the New skill editor', async () => {
    mountApi([])
    const saved: unknown[][] = []
    mountHomeApi({
      skillsSave: async (name, description, body, allowOverwrite) => {
        saved.push([name, description, body, allowOverwrite])
        return skill({ name, description })
      },
    })
    ;(window.aiOfficeIntegrations as IntegrationsApi).listSkills = async () => [
      skill({ name: 'fresh-skill' }),
    ]
    await render()
    await click(buttonWithText('New skill'))
    // the editor replaced the list, and Save stays off until the fields pass
    expect(buttonWithText('Save')?.disabled).toBe(true)
    await act(async () => {
      typeInto(field('skill-doc-name'), 'fresh-skill')
      typeInto(field('skill-doc-desc'), 'A docx thing')
      typeInto(field('skill-doc-body'), 'Step one.')
      await Promise.resolve()
      await Promise.resolve()
    })
    await click(buttonWithText('Save'))
    expect(saved).toEqual([['fresh-skill', 'A docx thing', 'Step one.', false]])
    expect(host.textContent).toContain('Saved fresh-skill')
    expect(host.textContent).toContain('fresh-skill')
  })

  it('refuses a bad name live and does not save', async () => {
    mountApi([])
    const saved: unknown[][] = []
    mountHomeApi({
      skillsSave: async (name, description, body) => {
        saved.push([name, description, body])
        return skill({ name, description })
      },
    })
    await render()
    await click(buttonWithText('New skill'))
    await act(async () => {
      typeInto(field('skill-doc-name'), 'not/a/name')
      typeInto(field('skill-doc-desc'), 'A docx thing')
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(host.textContent).toContain('This name cannot be used as a folder name.')
    expect(buttonWithText('Save')?.disabled).toBe(true)
    expect(saved).toEqual([])
  })

  it('loads the raw skill for editing and saves back over it', async () => {
    mountApi([skill({})])
    const rawText = vi.fn(async () =>
      [
        '---',
        'name: pptx-builder',
        'description: Build a pptx',
        '---',
        '',
        'Do the thing.',
        '',
      ].join('\n'),
    )
    const saved: unknown[][] = []
    mountHomeApi({
      skillsRawText: rawText,
      skillsSave: async (name, description, body, allowOverwrite) => {
        saved.push([name, description, body, allowOverwrite])
        return skill({ name, description })
      },
    })
    await render()
    await click(buttonWithText('Edit'))
    expect(rawText).toHaveBeenCalledWith('pptx-builder')
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    // the name is the folder on disk — present but locked
    const name = field('skill-doc-name') as HTMLInputElement
    expect(name.value).toBe('pptx-builder')
    expect(name.readOnly).toBe(true)
    expect((field('skill-doc-desc') as HTMLInputElement).value).toBe('Build a pptx')
    expect((field('skill-doc-body') as HTMLTextAreaElement).value).toBe('Do the thing.\n')
    await act(async () => {
      typeInto(field('skill-doc-body'), 'New body.')
      await Promise.resolve()
      await Promise.resolve()
    })
    await click(buttonWithText('Save'))
    expect(saved).toEqual([['pptx-builder', 'Build a pptx', 'New body.', true]])
    expect(host.textContent).toContain('Saved pptx-builder')
  })

  it('walks the AI interview and falls back to the template when no model is configured', async () => {
    mountApi([])
    const saved: unknown[][] = []
    mountHomeApi({
      skillsSave: async (name, description, body) => {
        saved.push([name, description, body])
        return skill({ name, description })
      },
    })
    await render()
    await click(buttonWithText('Create with AI'))
    const answer = async (text: string): Promise<void> => {
      await act(async () => {
        typeInto(field('skill-creator-answer'), text)
        await Promise.resolve()
        await Promise.resolve()
      })
    }
    await answer('Turn meeting notes into a report')
    await click(buttonWithText('Next'))
    await answer('when the user asks for a report from notes')
    await click(buttonWithText('Next'))
    await answer('1. Read the notes')
    await click(buttonWithText('Next'))
    // the last question (an example) is optional: Generate goes without it
    await click(buttonWithText('Generate'))
    expect(host.textContent).toContain('No AI model is configured')
    expect((field('skill-doc-desc') as HTMLInputElement).value).toBe(
      'when the user asks for a report from notes',
    )
    await click(buttonWithText('Save'))
    expect(saved).toEqual([
      [
        'turn-meeting-notes-into-a-report',
        'when the user asks for a report from notes',
        expect.any(String),
      ],
    ])
    expect(host.textContent).toContain('Saved turn-meeting-notes-into-a-report')
  })

  it('generates a draft from one model call and saves it', async () => {
    mountApi([])
    const saved: unknown[][] = []
    // the ai:stream bridge as the transport sees it: start records the request,
    // chunks come back through the listener the transport subscribed
    let streamListener: ((chunk: unknown) => void) | undefined
    const started: Array<{ requestId: string; system: string }> = []
    mountHomeApi({
      aiGskStatus: async () => ({ loggedIn: true }),
      onAiStream: (listener) => {
        streamListener = listener as (chunk: unknown) => void
        return () => undefined
      },
      aiStream: async (request) => {
        started.push(request as { requestId: string; system: string })
      },
      aiStreamCancel: async () => undefined,
      skillsSave: async (name, description, body) => {
        saved.push([name, description, body])
        return skill({ name, description })
      },
    })
    await render()
    await click(buttonWithText('Create with AI'))
    const answer = async (text: string): Promise<void> => {
      await act(async () => {
        typeInto(field('skill-creator-answer'), text)
        await Promise.resolve()
        await Promise.resolve()
      })
    }
    await answer('Turn meeting notes into a report')
    await click(buttonWithText('Next'))
    await answer('when the user asks for a report from notes')
    await click(buttonWithText('Next'))
    await answer('1. Read the notes')
    await click(buttonWithText('Next'))
    await click(buttonWithText('Generate'))
    expect(host.textContent).toContain('The AI is writing the skill…')

    // exactly one request went out, carrying the meta-prompt as its system
    expect(started).toHaveLength(1)
    expect(started[0]!.system).toContain('SKILL.md')

    // the reply streams back in deltas, prose and fence included
    const reply = [
      'Here you go:',
      '',
      '```markdown',
      '---',
      'name: notes-report',
      'description: Use when the user asks for a report from meeting notes.',
      '---',
      '',
      '1. Read the notes.',
      '```',
    ].join('\n')
    await act(async () => {
      streamListener?.({
        requestId: started[0]!.requestId,
        type: 'delta',
        text: reply.slice(0, 40),
      })
      streamListener?.({ requestId: started[0]!.requestId, type: 'delta', text: reply.slice(40) })
      streamListener?.({ requestId: started[0]!.requestId, type: 'done' })
      await Promise.resolve()
      await Promise.resolve()
      await Promise.resolve()
    })
    expect((field('skill-doc-name') as HTMLInputElement).value).toBe('notes-report')
    expect((field('skill-doc-desc') as HTMLInputElement).value).toBe(
      'Use when the user asks for a report from meeting notes.',
    )
    expect((field('skill-doc-body') as HTMLTextAreaElement).value).toContain('1. Read the notes.')
    await click(buttonWithText('Save'))
    expect(saved).toEqual([
      [
        'notes-report',
        'Use when the user asks for a report from meeting notes.',
        expect.any(String),
      ],
    ])
    expect(host.textContent).toContain('Saved notes-report')
  })

  it('previews the exact SKILL.md, live as the fields change', async () => {
    mountApi([])
    mountHomeApi({})
    await render()
    await click(buttonWithText('New skill'))
    const preview = () => host.querySelector<HTMLElement>('.set-intg-preview-code')
    expect(preview()).toBeTruthy()
    expect(preview()!.textContent).toBe(
      ['---', 'name: ', 'description: ', '---', '', ''].join('\n'),
    )
    await act(async () => {
      typeInto(field('skill-doc-name'), 'fresh-skill')
      typeInto(field('skill-doc-desc'), 'A docx thing')
      typeInto(field('skill-doc-body'), 'Step one.')
      await Promise.resolve()
      await Promise.resolve()
    })
    // frontmatter and body, exactly what Save will write to disk
    expect(preview()!.textContent).toBe(
      ['---', 'name: fresh-skill', 'description: A docx thing', '---', '', 'Step one.'].join('\n'),
    )
    await act(async () => {
      typeInto(field('skill-doc-body'), 'Step one.\nStep two.')
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(preview()!.textContent).toContain('Step two.')
    expect(host.textContent).toContain('SKILL.md preview')
    expect(host.textContent).toContain('The exact file Save writes to disk:')
    expect(host.textContent).toContain('$ARGUMENTS')
    expect(host.textContent).toContain('2 lines')
  })

  it('lists the parameter vocabulary in the insert menu and inserts at the caret', async () => {
    mountApi([])
    mountHomeApi({})
    await render()
    await click(buttonWithText('New skill'))
    await act(async () => {
      typeInto(field('skill-doc-body'), 'Step one.')
      await Promise.resolve()
      await Promise.resolve()
    })
    await click(buttonWithText('Insert parameter'))
    // the vocabulary expandSkillBody replaces, one entry per form, each glossed
    const items = () => [...host.querySelectorAll<HTMLElement>('.set-intg-param-item')]
    expect(items().map((el) => el.querySelector('.set-intg-param-token')?.textContent)).toEqual([
      '$ARGUMENTS',
      '$1',
      '${FILE_NAME}',
      '${SELECTION}',
      '${TODAY}',
      '${LANGUAGE}',
      '${APP}',
    ])
    for (const item of items()) {
      expect(
        (item.querySelector('.set-intg-param-desc')?.textContent ?? '').length,
      ).toBeGreaterThan(0)
    }
    // the caret the textarea was left at decides where the token lands
    const body = field('skill-doc-body') as HTMLTextAreaElement
    body.setSelectionRange(5, 5)
    await click(items()[0])
    expect(body.value).toBe('Step $ARGUMENTSone.')
    // the caret is restored right after the inserted token, and the menu closed
    expect(body.selectionStart).toBe(5 + '$ARGUMENTS'.length)
    expect(body.selectionEnd).toBe(5 + '$ARGUMENTS'.length)
    expect(host.querySelector('.set-intg-param-menu')).toBeFalsy()
  })

  it('inserts a context variable the same way, at the caret the last insert left', async () => {
    mountApi([])
    mountHomeApi({})
    await render()
    await click(buttonWithText('New skill'))
    await act(async () => {
      typeInto(field('skill-doc-body'), 'Report for ')
      await Promise.resolve()
      await Promise.resolve()
    })
    const body = field('skill-doc-body') as HTMLTextAreaElement
    body.setSelectionRange(11, 11) // the end of "Report for "
    await click(buttonWithText('Insert parameter'))
    const today = [...host.querySelectorAll<HTMLElement>('.set-intg-param-item')].find((el) =>
      el.querySelector('.set-intg-param-token')?.textContent?.includes('TODAY'),
    )
    await click(today)
    expect(body.value).toBe('Report for ${TODAY}')
    expect(body.selectionStart).toBe(11 + '${TODAY}'.length)
  })

  it('suggests a kebab-case name from the description', async () => {
    mountApi([])
    mountHomeApi({})
    await render()
    await click(buttonWithText('New skill'))
    // with no description there is nothing to derive from
    expect(buttonWithText('Generate from description')?.disabled).toBe(true)
    await act(async () => {
      typeInto(field('skill-doc-desc'), 'Turn meeting notes into a report')
      await Promise.resolve()
      await Promise.resolve()
    })
    await click(buttonWithText('Generate from description'))
    expect((field('skill-doc-name') as HTMLInputElement).value).toBe(
      'turn-meeting-notes-into-a-report',
    )
  })

  it('offers no name suggestion while editing, where the name is locked', async () => {
    mountApi([skill({})])
    mountHomeApi({
      skillsRawText: async () =>
        [
          '---',
          'name: pptx-builder',
          'description: Build a pptx',
          '---',
          '',
          'Do the thing.',
          '',
        ].join('\n'),
    })
    await render()
    await click(buttonWithText('Edit'))
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    expect((field('skill-doc-name') as HTMLInputElement).value).toBe('pptx-builder')
    expect(buttonWithText('Generate from description')?.disabled).toBe(true)
  })

  it('scans while looking, then lists what other products have', async () => {
    mountApi([skill({})])
    let resolveScan: (v: ForeignSkillCandidate[]) => void = () => {}
    mountHomeApi({
      skillsForeignList: () => new Promise<ForeignSkillCandidate[]>((r) => (resolveScan = r)),
    })
    await render()
    await click(buttonWithText('Import'))
    await click(buttonWithText('Import from other products…'))
    expect(host.textContent).toContain('Looking for skills from other products…')
    await act(async () => {
      resolveScan([foreign({})])
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(host.textContent).toContain('rollback-release')
  })

  it('groups foreign candidates by product and locks the already-imported ones', async () => {
    const from = (source: ForeignSkillCandidate['source']): ForeignSkillCandidate =>
      foreign({ name: `${source}-rule`, source, sourcePath: `/home/${source}-rule.md` })
    mountApi([skill({})])
    const importedPaths: string[] = []
    mountHomeApi({
      skillsForeignList: async () => [
        foreign({}),
        from('cursor'),
        from('windsurf'),
        foreign({
          name: 'already-here',
          sourcePath: '/home/already-here.md',
          alreadyImported: true,
        }),
      ],
      skillsImportForeign: async (p) => {
        importedPaths.push(p)
        return skill({ name: 'imported' })
      },
    })
    await render()
    await click(buttonWithText('Import'))
    await click(buttonWithText('Import from other products…'))
    // one heading per product, in the fixed order; a brand name, never translated
    const headings = () =>
      [...host.querySelectorAll('.set-intg-foreign-source')].map((e) => e.textContent)
    expect(headings()).toEqual(['Claude Code', 'Cursor', 'Windsurf'])
    expect(host.textContent).toContain('already-here')
    expect(host.textContent).toContain('Already imported')
    // the imported row cannot be checked and counts for nothing
    expect(foreignBox('already-here')!.disabled).toBe(true)
    await clickDeep(foreignBox('already-here'))
    expect(foreignBox('already-here')!.checked).toBe(false)
    expect(buttonWithText('Import selected (0)')?.disabled).toBe(true)
    expect(importedPaths).toEqual([])
  })

  it('imports the checked candidates one by one and lands on the refreshed list', async () => {
    mountApi([])
    const imported: string[] = []
    mountHomeApi({
      skillsForeignList: async () => [
        foreign({ name: 'one', sourcePath: '/home/one.md' }),
        foreign({ name: 'two', source: 'codex', sourcePath: '/home/two.md' }),
        foreign({ name: 'already', sourcePath: '/home/already.md', alreadyImported: true }),
      ],
      skillsImportForeign: async (p) => {
        imported.push(p)
        return skill({ name: 'imported-skill' })
      },
    })
    ;(window.aiOfficeIntegrations as IntegrationsApi).listSkills = async () => [
      skill({ name: 'one', path: '/app/skills/one/SKILL.md' }),
      skill({ name: 'two', path: '/app/skills/two/SKILL.md' }),
    ]
    await render()
    await click(buttonWithText('Import'))
    await click(buttonWithText('Import from other products…'))
    await clickDeep(foreignBox('one'))
    await clickDeep(foreignBox('two'))
    expect(buttonWithText('Import selected (2)')?.disabled).toBe(false)
    await clickDeep(buttonWithText('Import selected (2)'))
    // each checked row imported once, through its own sourcePath
    expect(imported).toEqual(['/home/one.md', '/home/two.md'])
    // the view closed onto the refreshed My skills list, with the count notice
    expect(host.textContent).toContain('Imported 2 skills')
    expect(host.textContent).not.toContain('Import from other products')
    expect(host.textContent).toContain('one')
    expect(host.textContent).toContain('two')
  })

  it('reports a failed import on its row and keeps the view for a retry', async () => {
    mountApi([])
    mountHomeApi({
      skillsForeignList: async () => [
        foreign({ name: 'good-one', sourcePath: '/home/good.md' }),
        foreign({ name: 'bad-one', sourcePath: '/home/bad.md' }),
      ],
      skillsImportForeign: async (p) => {
        if (p === '/home/bad.md') throw new Error('that file is a directory')
        return skill({ name: 'good-one' })
      },
    })
    await render()
    await click(buttonWithText('Import'))
    await click(buttonWithText('Import from other products…'))
    await clickDeep(foreignBox('good-one'))
    await clickDeep(foreignBox('bad-one'))
    await clickDeep(buttonWithText('Import selected (2)'))
    // the failure sits on its own row; the view stays so it can be retried
    expect(host.textContent).toContain('Import failed: that file is a directory')
    expect(host.textContent).toContain('Import from other products')
    // the row that landed now reads as already imported, and only the failed
    // one is still checked
    expect(foreignBox('good-one')!.disabled).toBe(true)
    expect(host.textContent).toContain('Already imported')
    expect(buttonWithText('Import selected (1)')).toBeTruthy()
  })

  it('says so when no other products have skills', async () => {
    mountApi([skill({})])
    mountHomeApi({ skillsForeignList: async () => [] })
    await render()
    await click(buttonWithText('Import'))
    await click(buttonWithText('Import from other products…'))
    expect(host.textContent).toContain('No skills from other products were found.')
    expect(host.querySelectorAll('.set-intg-foreign-row')).toHaveLength(0)
  })

  it('survives a preload without the foreign scanner', async () => {
    mountApi([skill({})])
    mountHomeApi({})
    await render()
    await click(buttonWithText('Import'))
    await click(buttonWithText('Import from other products…'))
    // an old bridge answers with nothing, which reads the same as an empty scan
    expect(host.textContent).toContain('No skills from other products were found.')
  })
})
