import { useEffect, useLayoutEffect, useRef, useState, type ReactNode, type RefObject } from 'react'
import { flushSync } from 'react-dom'
import { useDismissablePopover } from '@genoffice/ui'
import { IconCaret } from './icons'
import { GROUP_MENU } from './ribbon-fit'

const OPEN = 'rb-group-open'
const PANEL = '[data-rb-panel]'
const FOCUSABLE = 'button:not(:disabled), input:not(:disabled), select:not(:disabled)'

/**
 * The button a group becomes when the window is too narrow for it (Word's last
 * rung). The group's own controls stay mounted and drop down from it, so they
 * run the same commands and keep their state. Until the group folds this is a
 * span of the same size, which the fit measures.
 */
export function RbGroupMenu({ name, children }: { name: string; children: ReactNode }) {
  const button = useRef<HTMLElement>(null)
  const [folded, setFolded] = useState(false)
  const [open, setOpen] = useState(false)
  const group = () => button.current?.closest<HTMLElement>('.ribbon-group') ?? null
  const close = () => setOpen(false)

  useLayoutEffect(() => {
    const el = group()
    if (!el) return
    // the fit writes the class straight to the group; in step with it before the next paint
    const watch = new MutationObserver(() =>
      flushSync(() => {
        const now = el.classList.contains(GROUP_MENU)
        setFolded(now)
        if (!now) setOpen(false)
      }),
    )
    watch.observe(el, { attributes: true, attributeFilter: ['class'] })
    return () => watch.disconnect()
  }, [])

  useLayoutEffect(() => {
    const el = group()
    el?.classList.toggle(OPEN, open)
    // CSS keeps the menu inside the window, and has to be told how wide it came out
    const items = el?.querySelector<HTMLElement>(':scope > .ribbon-group-items')
    if (open && items) items.style.setProperty('--rb-menu-width', `${items.offsetWidth}px`)
    return () => el?.classList.remove(OPEN)
  }, [open])

  useDismissablePopover(open, close, { inside: () => [group()] })

  useEffect(() => {
    const el = group()
    if (!open || !el) return
    // what a click did shows once React has drawn it; the next key or press must not get in first
    let settle: (() => void) | null = null
    const onKey = (event: KeyboardEvent) => {
      settle?.()
      // a dropdown open inside takes the first Escape; asked on capture, before that dropdown closes
      if (event.key !== 'Escape' || el.querySelector(PANEL)) return
      setOpen(false)
      if (el.contains(document.activeElement)) button.current?.focus()
    }
    // a command closes the menu; opening or closing a control's own dropdown does not
    let hadPanel = false
    const onPress = () => {
      settle?.()
      hadPanel = !!el.querySelector(PANEL)
    }
    const onClick = (event: MouseEvent) => {
      const target = event.target as Element
      const control = target.closest('button, [role=menuitem], [role=option]')
      if (!control || control === button.current || !el.contains(control)) return
      const picked = !!target.closest(PANEL)
      const before = hadPanel
      const mine = () => {
        if (settle !== mine) return
        settle = null
        if (!el.querySelector(PANEL) && (picked || !before)) setOpen(false)
      }
      settle = mine
      window.setTimeout(mine, 0)
    }
    window.addEventListener('keydown', onKey, true)
    el.addEventListener('pointerdown', onPress, true)
    el.addEventListener('click', onClick)
    return () => {
      window.removeEventListener('keydown', onKey, true)
      el.removeEventListener('pointerdown', onPress, true)
      el.removeEventListener('click', onClick)
    }
  }, [open])

  const icon = (
    <span className="rb-big-icon">
      {children}
      <IconCaret />
    </span>
  )
  // the name is drawn by CSS here, so a search for the text finds only real controls
  if (!folded)
    return (
      <span ref={button} className="rb-group-menu-button" data-name={name} aria-hidden="true">
        {icon}
      </span>
    )
  return (
    <button
      ref={button as RefObject<HTMLButtonElement>}
      className={`rb-big rb-group-menu-button${open ? ' active' : ''}`}
      aria-haspopup="true"
      aria-expanded={open}
      data-tip={name}
      // command-exempt: opens the group's own menu
      onClick={(event) => {
        const byKeyboard = event.detail === 0
        setOpen((was) => !was)
        if (open || !byKeyboard) return
        const items = group()?.querySelector<HTMLElement>('.ribbon-group-items')
        // the controls show only once the group carries the open class
        requestAnimationFrame(() => items?.querySelector<HTMLElement>(FOCUSABLE)?.focus())
      }}
    >
      {icon}
      <span className="rb-group-menu-name">{name}</span>
    </button>
  )
}
