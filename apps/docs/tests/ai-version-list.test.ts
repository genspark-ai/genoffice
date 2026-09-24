import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { PmNode } from '../src/renderer/editor/convert'
import { AiVersionList } from '../src/renderer/ai/AiVersionList'
import type { DocVersion } from '../src/renderer/ai/version-history'
import { t } from '../src/renderer/i18n/locale'

function doc(tag: string): PmNode {
  return { type: 'doc', attrs: { tag } }
}

function version(over: Partial<DocVersion> & { id: number }): DocVersion {
  return {
    label: `turn ${over.id}`,
    time: `10:0${over.id}`,
    doc: doc(`before-${over.id}`),
    ...over,
  }
}

let roots: Array<{ root: Root; container: HTMLElement }> = []

// React 19 needs this flag for `act()` to flush; without it every render warns.
vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)

function mount(element: React.ReactElement) {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  roots.push({ root, container })
  return container
}

afterEach(() => {
  for (const { root, container } of roots) {
    act(() => root.unmount())
    container.remove()
  }
  roots = []
})

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

function rows(container: HTMLElement) {
  return Array.from(container.querySelectorAll('.ai-version-row'))
}

describe('AiVersionList', () => {
  it('renders nothing while the document has no versions yet', () => {
    const container = mount(
      createElement(AiVersionList, {
        versions: [],
        busy: false,
        onRollback: () => {},
        onUndo: () => {},
      }),
    )
    expect(container.querySelector('.ai-versions')).toBeNull()
  })

  it('lists every turn with its time and instruction', () => {
    const container = mount(
      createElement(AiVersionList, {
        versions: [version({ id: 1 }), version({ id: 2 })],
        busy: false,
        onRollback: () => {},
        onUndo: () => {},
      }),
    )
    expect(container.querySelector('.ai-versions-title')?.textContent).toContain(
      t('aiSnapshotsTitle'),
    )
    expect(rows(container)).toHaveLength(2)
    expect(rows(container)[0].textContent).toContain('10:01')
    expect(rows(container)[0].textContent).toContain('turn 1')
  })

  it('restores a version through onRollback', () => {
    const onRollback = vi.fn()
    const container = mount(
      createElement(AiVersionList, {
        versions: [version({ id: 7 })],
        busy: false,
        onRollback,
        onUndo: () => {},
      }),
    )
    const button = container.querySelector<HTMLButtonElement>('.ai-version-rollback')!
    expect(button.textContent).toBe(t('aiRollback'))
    act(() => button.click())
    expect(onRollback).toHaveBeenCalledWith(7)
  })

  it('flips a rolled-back version to its own undo', () => {
    const onUndo = vi.fn()
    const container = mount(
      createElement(AiVersionList, {
        versions: [version({ id: 3, rolledBack: true, rolledBackFrom: doc('as-it-was') })],
        busy: false,
        onRollback: () => {},
        onUndo,
      }),
    )
    const button = container.querySelector<HTMLButtonElement>('.ai-version-rollback')!
    expect(button.textContent).toBe(t('aiRollbackUndo'))
    expect(rows(container)[0].className).toContain('ai-version-rolled-back')
    act(() => button.click())
    expect(onUndo).toHaveBeenCalledWith(3)
  })

  it('keeps discarded versions listed but not restorable', () => {
    const onRollback = vi.fn()
    const container = mount(
      createElement(AiVersionList, {
        versions: [
          version({ id: 1, rolledBack: true, rolledBackFrom: doc('as-it-was') }),
          version({ id: 2, discarded: true }),
        ],
        busy: false,
        onRollback,
        onUndo: () => {},
      }),
    )
    expect(rows(container)).toHaveLength(2)
    expect(rows(container)[1].className).toContain('ai-version-discarded')
    // only the rolled-back row carries a button, and it undoes
    expect(container.querySelectorAll('.ai-version-rollback')).toHaveLength(1)
    act(() => container.querySelector<HTMLButtonElement>('.ai-version-rollback')!.click())
    expect(onRollback).not.toHaveBeenCalled()
  })

  it('says an expired version is expired instead of hiding its action', () => {
    // the stored snapshot is gone and this session never loaded it: the row is
    // part of the record and stays put, with the reason where its button was
    const container = mount(
      createElement(AiVersionList, {
        versions: [version({ id: 1, snapshotId: 'snap-1', expired: true, doc: undefined })],
        busy: false,
        onRollback: () => {},
        onUndo: () => {},
      }),
    )
    expect(rows(container)).toHaveLength(1)
    expect(rows(container)[0].className).toContain('ai-version-expired')
    expect(container.querySelector('.ai-version-rollback')).toBeNull()
    expect(container.querySelector('.ai-version-note')?.textContent).toBe(t('aiVersionExpired'))
  })

  it('still offers a version read back from disk once its document has been loaded', () => {
    const container = mount(
      createElement(AiVersionList, {
        versions: [version({ id: 1, snapshotId: 'snap-1', expired: true, doc: doc('loaded') })],
        busy: false,
        onRollback: () => {},
        onUndo: () => {},
      }),
    )
    expect(container.querySelector('.ai-version-rollback')?.textContent).toBe(t('aiRollback'))
    expect(container.querySelector('.ai-version-note')).toBeNull()
  })

  it('disables its buttons while a run is in flight', () => {
    const container = mount(
      createElement(AiVersionList, {
        versions: [version({ id: 1 })],
        busy: true,
        onRollback: () => {},
        onUndo: () => {},
      }),
    )
    expect(container.querySelector<HTMLButtonElement>('.ai-version-rollback')!.disabled).toBe(true)
  })
})
