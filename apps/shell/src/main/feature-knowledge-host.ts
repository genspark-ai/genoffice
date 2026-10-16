import { readdirSync, statSync } from 'node:fs'
import { basename, join, resolve, sep } from 'node:path'
import {
  filterKnowledgeEntries,
  type FeatureKnowledge,
  type KnowledgeEntry,
} from '@genoffice/agent-core'
import { isReadableDocument, readDocumentText } from './feature-documents'

/**
 * The knowledge kind's host side (K7): list what a source offers, and read one
 * entry out of it.
 *
 * The rules that keep it from becoming "read whatever you like":
 * an entry is addressed by **name**, and only a name the listing offered resolves;
 * a folder entry is resolved inside the folder and refused if it escapes (or is
 * reached through a link out); reading goes through the same reader and the same
 * bounds as read_document. The listing itself is bounded, so a huge folder cannot
 * turn into a huge answer. A source's include/exclude globs narrow what the
 * listing offers — and with it what the corpus holds and what a name resolves to.
 */

export const KNOWLEDGE_MAX_ENTRIES = 50

/** What the source offers: entry names and sizes, never paths. */
export function listKnowledgeEntries(knowledge: FeatureKnowledge): KnowledgeEntry[] {
  if (knowledge.source === 'file') {
    const stat = statSync(knowledge.path)
    if (!stat.isFile()) throw new Error('that knowledge source is not a file')
    return [{ name: basename(knowledge.path), bytes: stat.size }]
  }
  const stat = statSync(knowledge.path)
  if (!stat.isDirectory()) throw new Error('that knowledge source is not a folder')
  const entries = readdirSync(knowledge.path, { withFileTypes: true })
    .filter((entry) => entry.isFile() && isReadableDocument(entry.name))
    .map((entry) => ({
      name: entry.name,
      bytes: statSync(join(knowledge.path, entry.name)).size,
    }))
  // the source's own file rules first, then the cap: 50 *matching* entries fit
  // even in a folder with hundreds of files
  return filterKnowledgeEntries(knowledge, entries).slice(0, KNOWLEDGE_MAX_ENTRIES)
}

/**
 * The absolute path of one entry — and only a path inside the declared source.
 * Exported because this is the security-relevant step, so it is unit-tested.
 */
export function pathForEntry(knowledge: FeatureKnowledge, entry: string): string {
  const offered = listKnowledgeEntries(knowledge)
  if (!offered.some((one) => one.name === entry)) {
    throw new Error('the entry is not in that source: ' + entry)
  }
  if (knowledge.source === 'file') return knowledge.path
  const root = resolve(knowledge.path)
  const target = resolve(root, entry)
  if (target !== root && !target.startsWith(root + sep)) {
    throw new Error('that entry is outside the source')
  }
  return target
}

/** Read one entry, bounded by the same reader read_document uses. */
export async function readKnowledgeEntry(
  knowledge: FeatureKnowledge,
  entry: string,
): Promise<{ name: string; text: string; chars: number }> {
  return readDocumentText(pathForEntry(knowledge, entry))
}
