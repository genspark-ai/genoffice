/**
 * AnyDoc channels — generic document recognition, format conversion and
 * content extraction. Real OCR/PDF parsing lands via `packages/file-parse`
 * and `packages/pdf2docx`; for now these endpoints return placeholder
 * payloads that match the legacy single-file implementation.
 */
import { existsSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import { FILES_DIR, registerHandle } from '../common/index.js'

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

    return {
      id: `doc-${Date.now()}`,
      fileName,
      fileType: ext.slice(1),
      pages: 1,
      text: `这是从 ${fileName} 提取的文本内容。\n完整的 OCR 识别需要集成 Tesseract.js 或云端 OCR 服务。`,
      metadata: {
        size: statSync(filePath).size,
        created: statSync(filePath).birthtime,
        modified: statSync(filePath).mtime,
      },
      success: true,
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
    const bytes = readFileSync(filePath as string)

    if (['.txt', '.md', '.json', '.xml', '.html', '.csv'].includes(ext)) {
      return { text: bytes.toString('utf-8'), format: 'text' }
    } else if (['.docx', '.xlsx', '.pptx'].includes(ext)) {
      return { text: `Office 文档内容 (${ext})\n需要集成 mammoth.js 或专业解析库`, format: 'office' }
    } else if (ext === '.pdf') {
      return { text: `PDF 文档内容\n需要集成 pdf-parse 或 pdf.js`, format: 'pdf' }
    } else if (['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'].includes(ext)) {
      return { text: `图片内容\n需要 OCR 识别 (Tesseract.js)`, format: 'image' }
    }

    return { text: '未知文件格式', format: 'unknown' }
  })

  registerHandle('anydoc:extract-tables', async (_event: unknown, filePath: unknown) => {
    if (!existsSync(filePath as string)) {
      return null
    }

    return {
      tables: [],
      message: '表格提取需要专业解析库支持',
    }
  })

  registerHandle('anydoc:extract-images', async (_event: unknown, filePath: unknown) => {
    if (!existsSync(filePath as string)) {
      return null
    }

    return {
      images: [],
      message: '图片提取功能需要实现',
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
