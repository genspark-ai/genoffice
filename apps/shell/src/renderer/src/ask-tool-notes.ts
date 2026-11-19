import {
  CREATE_DOCUMENT_TOOL,
  LOAD_GUIDE_TOOL,
  READ_ATTACHMENT_TOOL,
  RUN_GENOFFICE_TOOL,
  type AgentToolCall,
} from '@genoffice/agent-core'

/**
 * The activity chips of one Ask AI answer: one per tool call, labelled with
 * what the call actually did (the CLI command and the file it touches, the
 * knowledge-base file read), and folded with the previous chip when a turn
 * repeats the same action on the same target.
 */

export type ToolNoteKind = 'search' | 'read' | 'create' | 'cli'

export interface ToolNote {
  label: string
  kind: ToolNoteKind
  /** the command / file behind the generic label, when the call names one */
  detail?: string
  /** consecutive calls folded into this chip */
  count: number
}

const FILE_ARG = /\.(docx|xlsx|pptx|pdf|md|html|csv|json|txt)$/i

function baseName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path
}

export function toolNoteKind(call: AgentToolCall): ToolNoteKind {
  if (call.name === 'search_knowledge_base') return 'search'
  if (call.name === CREATE_DOCUMENT_TOOL) return 'create'
  if (call.name === RUN_GENOFFICE_TOOL || call.name === LOAD_GUIDE_TOOL) return 'cli'
  return 'read'
}

/**
 * `slides create · deck.pptx` for a CLI run, `guide slides` for a guide load,
 * the file name for a read or create; undefined when the call carries none.
 */
export function toolNoteDetail(
  call: AgentToolCall,
  readName?: (index: number) => string | undefined,
  attachmentName?: (index: number) => string | undefined,
): string | undefined {
  if (call.name === LOAD_GUIDE_TOOL) {
    const domain = typeof call.input.domain === 'string' ? call.input.domain : ''
    return domain ? `guide ${domain}` : undefined
  }
  if (call.name === RUN_GENOFFICE_TOOL) {
    const args = Array.isArray(call.input.args)
      ? (call.input.args as unknown[]).filter((a): a is string => typeof a === 'string')
      : []
    const verb = args.slice(0, 2).filter((a) => !a.startsWith('-') && !FILE_ARG.test(a))
    const at = args.indexOf('--out')
    const target = at >= 0 ? args[at + 1] : args.find((a) => FILE_ARG.test(a))
    const file = target ? baseName(target) : args.includes('--help') ? '--help' : ''
    return [verb.join(' '), file].filter(Boolean).join(' · ') || undefined
  }
  if (call.name === CREATE_DOCUMENT_TOOL) {
    const title = typeof call.input.title === 'string' ? call.input.title.trim() : ''
    return title || undefined
  }
  if (call.name === 'read_knowledge_file' || call.name === READ_ATTACHMENT_TOOL) {
    const index = Number(call.input.index)
    const lookup = call.name === READ_ATTACHMENT_TOOL ? attachmentName : readName
    const name = Number.isInteger(index) ? lookup?.(index) : undefined
    return name ? baseName(name) : undefined
  }
  return undefined
}

export function appendToolNote(notes: readonly ToolNote[], note: ToolNote): ToolNote[] {
  const last = notes[notes.length - 1]
  if (last && last.kind === note.kind && last.detail === note.detail) {
    return [...notes.slice(0, -1), { ...last, count: last.count + note.count }]
  }
  return [...notes, note]
}
