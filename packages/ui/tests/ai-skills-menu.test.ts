/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { AiSkillChips, AiSkillsMenu } from '../src/AiSkillsMenu'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

let host: HTMLDivElement
let root: ReturnType<typeof createRoot>

beforeEach(() => {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})
afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

const SKILLS = [
  { name: 'tables', description: 'Header rows and totals' },
  { name: 'tone', description: 'Write plainly' },
]

async function settle(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
  })
}

const click = (el: Element | null | undefined) => {
  expect(el).toBeTruthy()
  act(() => (el as HTMLElement).click())
}

function mount(over: Partial<Parameters<typeof AiSkillsMenu>[0]> = {}) {
  const onToggle = vi.fn()
  const listSkills = vi.fn(async () => SKILLS)
  act(() => {
    root.render(
      createElement(AiSkillsMenu, {
        lang: 'en',
        listSkills,
        selected: [],
        onToggle,
        ...over,
      }),
    )
  })
  return { onToggle, listSkills }
}

describe('AiSkillsMenu', () => {
  it('opens on a root menu with Add file and Skills when the app takes files', async () => {
    const onAddFile = vi.fn()
    const { listSkills } = mount({ onAddFile })
    click(host.querySelector('.ai-plus-btn'))
    const items = [...host.querySelectorAll('.ai-plus-item')].map((b) => b.textContent)
    expect(items).toEqual(['Add file', 'Skills'])
    // the list is not fetched until the skills page opens
    expect(listSkills).not.toHaveBeenCalled()
    click(host.querySelectorAll('.ai-plus-item')[0])
    expect(onAddFile).toHaveBeenCalledTimes(1)
    expect(host.querySelector('.ai-plus-pop')).toBeNull()
  })

  it('opens straight on the skills page when the app takes no files', async () => {
    const { listSkills } = mount()
    click(host.querySelector('.ai-plus-btn'))
    await settle()
    expect(listSkills).toHaveBeenCalledTimes(1)
    expect(host.querySelector('.ai-plus-back')).toBeNull()
    expect([...host.querySelectorAll('.ai-plus-skill-name')].map((n) => n.textContent)).toEqual([
      'tables',
      'tone',
    ])
  })

  it('filters by name or description and toggles a skill without closing', async () => {
    const { onToggle } = mount()
    click(host.querySelector('.ai-plus-btn'))
    await settle()
    const search = host.querySelector<HTMLInputElement>('.ai-plus-search')!
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
      setter.call(search, 'plain')
      search.dispatchEvent(new Event('input', { bubbles: true }))
    })
    expect([...host.querySelectorAll('.ai-plus-skill-name')].map((n) => n.textContent)).toEqual([
      'tone',
    ])
    click(host.querySelector('.ai-plus-skill'))
    expect(onToggle).toHaveBeenCalledWith('tone')
    expect(host.querySelector('.ai-plus-pop')).not.toBeNull()
  })

  it('marks selected skills and offers Manage only when a handler is given', async () => {
    const onManage = vi.fn()
    mount({ selected: ['tables'], onManage })
    click(host.querySelector('.ai-plus-btn'))
    await settle()
    expect(host.querySelector('.ai-plus-skill.on .ai-plus-skill-name')?.textContent).toBe('tables')
    click(host.querySelector('.ai-plus-manage'))
    expect(onManage).toHaveBeenCalledTimes(1)
    expect(host.querySelector('.ai-plus-pop')).toBeNull()
  })

  it('says when nothing is installed', async () => {
    mount({ listSkills: async () => [] })
    click(host.querySelector('.ai-plus-btn'))
    await settle()
    expect(host.querySelector('.ai-plus-empty')?.textContent).toBe('No skills installed yet')
  })
})

describe('AiSkillChips', () => {
  it('renders nothing for an empty list', () => {
    act(() => root.render(createElement(AiSkillChips, { lang: 'en', names: [] })))
    expect(host.textContent).toBe('')
  })

  it('is removable only when given a handler', () => {
    const onRemove = vi.fn()
    act(() => root.render(createElement(AiSkillChips, { lang: 'en', names: ['tables'], onRemove })))
    click(host.querySelector('.ai-skill-chip-remove'))
    expect(onRemove).toHaveBeenCalledWith('tables')
    act(() => root.render(createElement(AiSkillChips, { lang: 'en', names: ['tables'] })))
    expect(host.querySelector('.ai-skill-chip-remove')).toBeNull()
    expect(host.querySelector('.ai-skill-chips-sent')).not.toBeNull()
  })
})
