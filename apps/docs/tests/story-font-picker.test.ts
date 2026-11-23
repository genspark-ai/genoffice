/**
 * The story bar's font picker component: the curated faces pick directly,
 * the free-text field applies on Enter or Use, and the button carries the
 * active face's name.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { StoryFontPicker } from '../src/renderer/components/StoryFontPicker'
import { t } from '../src/renderer/i18n/locale'
import { STORY_FACE_CHOICES } from '../src/renderer/story-font'

describe('the story bar font picker', () => {
  let root: Root
  let container: HTMLElement
  let open = false
  let picked: string | undefined

  const renderPicker = (value: string) => {
    act(() =>
      root.render(
        createElement(StoryFontPicker, {
          value,
          open,
          onOpenChange: (v: boolean) => {
            open = v
            renderPicker(value)
          },
          onChange: (family: string) => {
            picked = family
            open = false
            renderPicker(family)
          },
        }),
      ),
    )
  }

  /** a controlled input only updates through the native value setter */
  const type = (input: HTMLInputElement, value: string) => {
    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')!.set!
    act(() => {
      setter.call(input, value)
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
  }

  beforeEach(() => {
    localStorage.clear()
    picked = undefined
    open = false
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
  })

  it('the closed button names the automatic default; a picked face shows on it', () => {
    renderPicker('')
    expect(container.querySelector('button')!.textContent).not.toContain('Aa ')
    renderPicker('IBM Plex Mono')
    expect(container.querySelector('button')!.textContent).toContain('Aa IBM Plex Mono')
  })

  it('open lists the automatic row, every curated face, and the free-text field', () => {
    renderPicker('')
    open = true
    renderPicker('')
    const items = [...container.querySelectorAll('.story-font-item')]
    expect(items[0]!.textContent).toContain(t('storyFontAuto'))
    expect(items.length).toBe(STORY_FACE_CHOICES.length + 1)
    expect(container.querySelectorAll('.story-font-dot')).toHaveLength(
      STORY_FACE_CHOICES.length - 2,
    )
    expect(container.querySelector('.story-font-bundled')).not.toBeNull()
    expect(container.querySelector('.story-font-any input')).not.toBeNull()
  })

  it('a curated pick reports the family and closes the panel', () => {
    renderPicker('')
    open = true
    renderPicker('')
    const plex = [...container.querySelectorAll<HTMLButtonElement>('.story-font-item')].find((el) =>
      el.textContent!.includes('IBM Plex Mono'),
    )!
    act(() => plex.click())
    expect(picked).toBe('IBM Plex Mono')
    expect(open).toBe(false)
  })

  it('the free-text field applies on Enter and Use; rule characters are stripped', () => {
    renderPicker('')
    open = true
    renderPicker('')
    const input = container.querySelector<HTMLInputElement>('.story-font-any input')!
    const use = [...container.querySelectorAll<HTMLButtonElement>('.story-font-any button')].at(-1)!

    type(input, 'Lora')
    expect(use.disabled).toBe(false)
    act(() => use.click())
    expect(picked).toBe('Lora')

    open = true
    renderPicker('')
    const input2 = container.querySelector<HTMLInputElement>('.story-font-any input')!
    type(input2, "Lora'; }")
    const enter = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })
    act(() => input2.dispatchEvent(enter))
    expect(picked).toBe('Lora')
  })
})
