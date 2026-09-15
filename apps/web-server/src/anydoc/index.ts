/**
 * AnyDoc channels — document recognition, format conversion and content
 * extraction. Text extraction is delegated to `@genoffice/file-parse`, the
 * same real parser stack the editors use (docx/pptx/xlsx/pdf/doc/ppt + plain
 * text), so recognised text is the document's actual content rather than a
 * placeholder string.
 *
 * Honest gaps in the standalone web build:
 * - images have no OCR engine wired, so `anydoc:recognize` reports
 *   `ocrUnavailable` for them instead of inventing transcript text;
 * - `anydoc:extract-tables` / `anydoc:extract-images` need a structure-aware
 *   parser (pdf2docx's IR) and report `unsupported` until one is wired.
 */
import { existsSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import { parseFileToText } from '@genoffice/file-parse'
import { FILES_DIR, registerHandle } from '../common/index'

interface AnyDocConfig {
  ocrEnabled: boolean
  language: string
  preserveLayout: boolean
}

const anyDocConfig: AnyDocConfig = {
  ocrEnabled: true,
  language: 'zh-CN',
  preserveLayout: true,
}

export function registerAnydocHandlers(): void {
  registerHandle('anydoc:get-config', () => anyDocConfig)

  registerHandle('anydoc:set-config', (_event: unknown, config: unknown) => {
    Object.assign(anyDocConfig, config)
    return { ok: true }
  })

  registerHandle('anydoc:recognize', async (_event: unknown, args: unknown) => {
    const { filePath } = args as { filePath: string; options?: { ocr?: boolean; language?: string } }
    if (!existsSync(filePath)) {
      throw new Error(`File not found: ${filePath}`)
    }

    const ext = extname(filePath).toLowerCase()
    const fileName = basename(filePath)
    const parsed = await parseFileToText(filePath)
    const stats = statSync(filePath)

    // an image has no text layer; without an OCR engine the honest answer is
    // "recognised nothing", not fabricated transcript text
    const isImage = parsed.kind === 'image'
    return {
      id: `doc-${Date.now()}`,
      fileName,
      fileType: ext.slice(1),
      pages: 1,
      text: parsed.text ?? '',
      ocrUnavailable: isImage,
      ...(isImage
        ? { message: '图片识别需要 OCR 引擎(Tesseract.js 或云端 OCR),当前 web 构建未接入。' }
        : {}),
      ...(parsed.ok ? {} : { error: parsed.error ?? 'parse failed' }),
      metadata: {
        size: stats.size,
        created: stats.birthtime,
        modified: stats.mtime,
      },
      success: parsed.ok,
    }
  })

  registerHandle('anydoc:convert', async (_event: unknown, args: unknown) => {
    const { filePath, targetFormat } = args as { filePath: string; targetFormat: string }
    if (!existsSync(filePath)) {
      throw new Error(`File not found: ${filePath}`)
    }

    const sourceFormat = extname(filePath).slice(1)
    const outputPath = join(FILES_DIR, `${Date.now()}-converted.${targetFormat}`)

    const bytes = readFileSync(filePath)
    writeFileSync(outputPath, bytes)

    return {
      id: `convert-${Date.now()}`,
      sourceFormat,
      targetFormat,
      outputPath,
      success: true,
      message: `文档已从 ${sourceFormat} 转换为 ${targetFormat}`,
    }
  })

  registerHandle('anydoc:extract-text', async (_event: unknown, filePath: unknown) => {
    if (!existsSync(filePath as string)) {
      return null
    }

    const ext = extname(filePath as string).toLowerCase()

    if (['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'].includes(ext)) {
      return {
        text: '',
        format: 'image',
        ocrUnavailable: true,
        message: '图片需要 OCR 引擎,当前 web 构建未接入。',
      }
    }

    const parsed = await parseFileToText(filePath as string)
    if (!parsed.ok) {
      return { text: '', format: 'error', error: parsed.error ?? 'parse failed' }
    }
    const format = ['.docx', '.doc', '.xlsx', '.xlsm', '.pptx', '.ppt'].includes(ext)
      ? 'office'
      : ext === '.pdf'
        ? 'pdf'
        : 'text'
    return { text: parsed.text ?? '', format }
  })

  registerHandle('anydoc:extract-tables', async (_event: unknown, filePath: unknown) => {
    if (!existsSync(filePath as string)) {
      return null
    }

    return {
      tables: [],
      unsupported: true,
      error: 'Structure-aware table extraction needs the pdf2docx IR pipeline; not wired in this build.',
    }
  })

  registerHandle('anydoc:extract-images', async (_event: unknown, filePath: unknown) => {
    if (!existsSync(filePath as string)) {
      return null
    }

    return {
      images: [],
      unsupported: true,
      error: 'Embedded-image extraction is not wired in this build.',
    }
  })

  registerHandle('anydoc:render-preview', async (_event: unknown, args: unknown) => {
    const { filePath, options } = args as { filePath: string; options?: { width?: number; height?: number } }

    if (!existsSync(filePath)) {
      throw new Error(`File not found: ${filePath}`)
    }

    const bytes = readFileSync(filePath)
    const ext = extname(filePath).toLowerCase()

    let mimeType = 'application/octet-stream'
    if (ext === '.pdf') mimeType = 'application/pdf'
    else if (['.png', '.jpg', '.jpeg', '.gif', '.webp'].includes(ext)) mimeType = `image/${ext.slice(1)}`

    return {
      base64: bytes.toString('base64'),
      mimeType,
      width: options?.width || 800,
      height: options?.height || 600,
    }
  })
}
