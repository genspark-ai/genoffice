import { basename } from 'node:path'
import type { ScriptRunOutcome } from '@genoffice/agent-core'
import { readInput } from '@genoffice/cli/headless-docs'
import { pdfText, PDF_PAGE_PREVIEW_CHARS, type PdfText } from '@genoffice/cli/headless-formats'
import {
  PDF_SNAPSHOT_PAGE_CAP,
  ScriptRunError,
  type PdfSnapshot,
  type SandboxOutcome,
} from './script-shared'

/**
 * The pdf half of the script kind — read-side only, on purpose. A pdf script
 * scans pages ($.pdf.info(), $.pdf.pages(from, to, maxChars?)) and returns a
 * report, which the run hands back as its result; the win is condensing a long
 * document into a short answer without pulling every page into a model's
 * context. Writes would need a shared headless pdf ops/save module and a bounded
 * batch input, which do not exist yet (mods-implementation-report.md §8.2.6), so
 * there is no apply path and the sandbox mounts no propose.
 */

export const pdfHost = {
  async snapshot(documentPath: string): Promise<PdfSnapshot> {
    let info: PdfText
    try {
      info = await pdfText(readInput(documentPath), {
        maxChars: PDF_PAGE_PREVIEW_CHARS,
        select: (pageCount) => [1, Math.min(pageCount, PDF_SNAPSHOT_PAGE_CAP)],
      })
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err)
      throw new ScriptRunError(
        basename(documentPath) + ': ' + message,
        /encrypt|password/i.test(message)
          ? 'pdf scripts cannot open password-protected files'
          : 'check that the file is a readable pdf',
      )
    }
    return {
      format: 'pdf',
      pdf: {
        pages: info.pages,
        encrypted: info.encrypted,
        pagesRead: info.pages_read.length,
        pageCap: PDF_SNAPSHOT_PAGE_CAP,
        ...(info.title === undefined ? {} : { title: info.title }),
        ...(info.author === undefined ? {} : { author: info.author }),
        ...(info.subject === undefined ? {} : { subject: info.subject }),
        ...(info.creator === undefined ? {} : { creator: info.creator }),
        ...(info.producer === undefined ? {} : { producer: info.producer }),
        pageTexts: info.pages_read.map((page) => ({
          page: page.page,
          width_pt: page.width_pt,
          height_pt: page.height_pt,
          text: page.text,
          truncated: page.truncated,
        })),
      },
    }
  },
  plan({ outcome }: { outcome: SandboxOutcome }): ScriptRunOutcome {
    return {
      applied: 0,
      plan: [],
      failures: [],
      log: outcome.log,
      ms: outcome.ms,
      ...(outcome.result === undefined ? {} : { result: outcome.result }),
    }
  },
}
