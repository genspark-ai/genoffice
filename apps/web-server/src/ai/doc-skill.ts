/**
 * Doc-AI skill handlers — write continue/expand/shrink/rewrite, tone adjust,
 * format suggest/apply. Placeholder implementations carried over verbatim
 * from the legacy `index.ts` so behaviour stays identical at this phase.
 */
import { registerHandle } from '../common/index.js'

export function registerDocAiSkillHandlers(): void {
  registerHandle('ai:doc-write-continue', async (_event: unknown, args: unknown) => {
    const { docId, content, cursor, length } = args as {
      docId: string
      content: string
      cursor: number
      length?: number
    }

    const targetLength = length || 200
    return {
      text: `根据上文续写的内容，关于"${content.slice(0, 50)}..."的详细展开说明...`,
      insertedAt: cursor,
      length: targetLength,
      style: 'continuation',
    }
  })

  registerHandle('ai:doc-write-expand', async (_event: unknown, args: unknown) => {
    const { content, targetLength } = args as { content: string; targetLength?: number }

    return {
      text: `${content}\n\n详细说明：此处内容已被扩展，包含更多细节和示例。`,
      originalLength: content.length,
      expandedLength: (targetLength || content.length * 2),
      changes: ['added_examples', 'added_details', 'added_explanations'],
    }
  })

  registerHandle('ai:doc-write-shrink', async (_event: unknown, args: unknown) => {
    const { content, ratio } = args as { content: string; ratio?: number }
    const shrinkRatio = ratio || 0.5
    const targetLength = Math.floor(content.length * shrinkRatio)

    return {
      text: content.slice(0, targetLength) + '...',
      originalLength: content.length,
      shrunkLength: targetLength,
      ratio: shrinkRatio,
      summary: '原文核心内容已压缩',
    }
  })

  registerHandle('ai:doc-write-rewrite', async (_event: unknown, args: unknown) => {
    const { content, style } = args as { content: string; style?: 'formal' | 'casual' | 'simple' }

    const styleMap = {
      formal: '正式',
      casual: '轻松',
      simple: '简洁',
    }

    return {
      text: `[${styleMap[style || 'formal']}风格改写]${content}`,
      originalStyle: 'original',
      newStyle: style || 'formal',
      changes: ['rephrased', 'tone_adjusted'],
    }
  })

  registerHandle('ai:doc-tone-adjust', async (_event: unknown, args: unknown) => {
    const { content, tone } = args as {
      content: string
      tone: 'professional' | 'friendly' | 'persuasive' | 'academic'
    }

    const toneDescriptions = {
      professional: '专业正式的语调',
      friendly: '友好亲切的语调',
      persuasive: '有说服力的语调',
      academic: '学术严谨的语调',
    }

    return {
      text: `[${toneDescriptions[tone]}调整后的内容]${content}`,
      originalTone: 'neutral',
      newTone: tone,
    }
  })

  registerHandle('ai:doc-format-suggest', async (_event: unknown, args: unknown) => {
    const { content } = args as { content: string }

    return {
      suggestions: [
        { type: 'heading', level: 1, text: '建议添加标题' },
        { type: 'list', style: 'bullet', items: ['要点1', '要点2', '要点3'] },
        { type: 'spacing', before: 12, after: 6 },
        { type: 'font', name: '微软雅黑', size: 12 },
      ],
      score: 0.85,
    }
  })

  registerHandle('ai:doc-format-apply', async (_event: unknown, args: unknown) => {
    const { suggestions } = args as { suggestions: Array<{ type: string; [key: string]: unknown }> }

    return {
      applied: suggestions.length,
      changes: suggestions.map(s => ({ type: s.type, applied: true })),
      preview: true,
    }
  })
}
