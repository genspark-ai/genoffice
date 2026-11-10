import { describe, expect, it } from 'vitest'
import {
  createSessionHost,
  createSessionTools,
  type FamilyDriver,
} from '../../src/main/mcp/tools/session-tools'

/**
 * Unit tests for the shared visible-session host.
 *
 * The three families fold their create/save lifecycle into one
 * create_session / save_session pair; the host is what keeps the active tab
 * unambiguous. The family tool modules and the e2e specs cover the routing end
 * to end, so these stay on the host's own contract.
 */

function driver(family: 'docx' | 'pptx' | 'xlsx', wcId: number): FamilyDriver {
  return {
    family,
    openBlankTab: async () => wcId,
    save: async (_wcId, path) => ({ ok: true, path }),
  }
}

describe('SessionHost', () => {
  it('requires create_session before any content tool runs', () => {
    const host = createSessionHost()
    expect(host.current()).toBeNull()
    expect(() => host.require('docx')).toThrow(/no session is open — call create_session first/)
  })

  it('refuses a content tool aimed at a different family than the active session', () => {
    const host = createSessionHost()
    host.begin('pptx', 7)
    expect(() => host.require('docx')).toThrow(/active session is a presentation/)
    expect(() => host.require('xlsx')).toThrow(/active session is a presentation/)
    expect(host.require('pptx')).toBe(7)
  })

  it('a new session replaces the previous one', () => {
    const host = createSessionHost()
    host.begin('docx', 1)
    host.begin('xlsx', 2)
    expect(host.current()?.family).toBe('xlsx')
    expect(host.require('xlsx')).toBe(2)
    expect(() => host.require('docx')).toThrow(/active session is a spreadsheet/)
  })

  it('end() clears the session so edits fall back to the guided error', () => {
    const host = createSessionHost()
    host.begin('docx', 1)
    host.end()
    expect(host.current()).toBeNull()
    expect(() => host.require('docx')).toThrow(/no session is open/)
  })
})

describe('createSessionTools', () => {
  it('registers nothing when no family is wired (headless runs)', () => {
    expect(createSessionTools([], createSessionHost())).toEqual([])
  })

  it('begins a session and saves it to the active family driver, then ends it', async () => {
    const host = createSessionHost()
    const saved: Array<{ wcId: number; path: string; overwrite: boolean }> = []
    const pptx: FamilyDriver = {
      ...driver('pptx', 11),
      save: async (wcId, path, overwrite) => {
        saved.push({ wcId, path, overwrite })
        return { ok: true, path }
      },
    }
    const [create, save] = createSessionTools([pptx], host)

    const created = (await create!.handler({ family: 'pptx' })) as { sessionId: number }
    expect(created.sessionId).toBe(11)
    expect(host.current()?.family).toBe('pptx')

    const result = await save!.handler({ path: '/tmp/out.pptx', overwrite: true })
    expect(result).toEqual({ ok: true, path: '/tmp/out.pptx' })
    expect(saved).toEqual([{ wcId: 11, path: '/tmp/out.pptx', overwrite: true }])
    expect(host.current()).toBeNull()
  })

  it('rejects a relative save path and keeps the session open', async () => {
    const host = createSessionHost()
    const [create, save] = createSessionTools([driver('docx', 3)], host)
    await create!.handler({ family: 'docx' })
    await expect(save!.handler({ path: 'relative.docx' })).rejects.toThrow(/path must be absolute/)
    expect(host.current()?.family).toBe('docx')
  })

  it('rejects save_session with no active session', async () => {
    const host = createSessionHost()
    const [, save] = createSessionTools([driver('docx', 3)], host)
    await expect(save!.handler({ path: '/tmp/out.docx' })).rejects.toThrow(/no session is open/)
  })

  it('refuses a save path outside the family format, keeping the session open', async () => {
    const host = createSessionHost()
    const saved: string[] = []
    const docx: FamilyDriver = {
      ...driver('docx', 5),
      save: async (_wcId, path) => {
        saved.push(path)
        return { ok: true, path }
      },
    }
    const [create, save] = createSessionTools([docx], host)
    await create!.handler({ family: 'docx' })
    await expect(save!.handler({ path: '/tmp/report.pdf' })).rejects.toThrow(
      /must be saved as \.docx \(got "\.pdf"\)/,
    )
    // still active: the agent can retry with the right extension
    expect(host.current()?.family).toBe('docx')
    await expect(save!.handler({ path: '/tmp/report.docx' })).resolves.toMatchObject({ ok: true })
  })

  it('appends the family extension when save_session gets an extensionless path', async () => {
    const host = createSessionHost()
    const saved: string[] = []
    const docx: FamilyDriver = {
      ...driver('docx', 9),
      save: async (_wcId, path) => {
        saved.push(path)
        return { ok: true, path }
      },
    }
    const [create, save] = createSessionTools([docx], host)
    await create!.handler({ family: 'docx' })
    await save!.handler({ path: '/tmp/notes' })
    expect(saved).toEqual(['/tmp/notes.docx'])
  })

  // Drivers disagree on the success shape (docs returns {ok,path}, slides just
  // {path}). An agent checking `ok` across families would read the slides save
  // as a failure, so the shared tool normalizes it.
  it('reports ok:true even when the driver omits it', async () => {
    const host = createSessionHost()
    const pptx: FamilyDriver = {
      ...driver('pptx', 11),
      save: async (_wcId, path) => ({ path }),
    }
    const [create, save] = createSessionTools([pptx], host)
    await create!.handler({ family: 'pptx' })
    await expect(save!.handler({ path: '/tmp/deck.pptx' })).resolves.toMatchObject({
      ok: true,
      path: '/tmp/deck.pptx',
    })
  })

  // A driver may report a failed write as data instead of throwing — sheets
  // forwards the renderer's SaveOutcome `{ok:false}`. Returning that verbatim
  // and ending the session told the agent the file was written, then left it
  // with no session to retry against.
  it('raises a tool error when the driver reports a failed save, keeping the session', async () => {
    const host = createSessionHost()
    const xlsx: FamilyDriver = {
      ...driver('xlsx', 12),
      save: async () => ({ ok: false, reason: 'the workbook could not be written' }),
    }
    const [create, save] = createSessionTools([xlsx], host)
    await create!.handler({ family: 'xlsx' })

    await expect(save!.handler({ path: '/tmp/out.xlsx' })).rejects.toThrow(
      /could not be saved: the workbook could not be written/,
    )
    // the edits are still in the tab, so the caller can retry
    expect(host.current()?.family).toBe('xlsx')
  })

  it('reports a bare failed save without inventing a reason', async () => {
    const host = createSessionHost()
    const xlsx: FamilyDriver = { ...driver('xlsx', 13), save: async () => ({ ok: false }) }
    const [create, save] = createSessionTools([xlsx], host)
    await create!.handler({ family: 'xlsx' })
    await expect(save!.handler({ path: '/tmp/out.xlsx' })).rejects.toThrow(
      /^the file could not be saved$/,
    )
  })
})
