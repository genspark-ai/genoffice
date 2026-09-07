/**
 * Core AI handlers — settings, login stub, chat/stream, web/image search,
 * sentiment, smart-summary, translation, language detection, generate.
 *
 * Each handler is registered at module load via the shared registry. The
 * function bodies are intentionally left as the original placeholders so
 * Phase 1.2 / 1.3 can swap them for real provider-backed logic without
 * touching the wiring.
 */
import { AI_STREAMS, registerHandle } from '../common/index.js'
import { callMiniMax, generateAIResponse } from './minimax.js'

let aiSettings = {
  provider: 'genspark',
  model: 'auto',
  temperature: 0.7,
  maxTokens: 4096,
  streaming: true,
}

export function registerAiCoreHandlers(): void {
  registerHandle('ai:get-settings', () => aiSettings)
  registerHandle('ai:set-settings', (_event: unknown, settings: unknown) => {
    aiSettings = { ...aiSettings, ...(settings as Record<string, unknown>) }
    return { ok: true }
  })

  registerHandle('ai:gsk-login', () => ({
    loggedIn: true,
    email: 'web-user@genoffice.ai',
    credits: 1000,
  }))

  registerHandle('ai:chat', async (_event: unknown, request: unknown) => {
    const req = request as { message?: string; system?: string; sessionId?: string; context?: unknown }
    const message = req.message || ''
    const context = req.context

    const minimaxKey = process.env.MINIMAX_API_KEY
    const systemPrompt = req.system || '你是一个专业的办公助手，帮助用户处理文档、表格和幻灯片。'

    try {
      if (minimaxKey) {
        const result = await callMiniMax(
          minimaxKey,
          [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: message },
          ],
          aiSettings.model || 'MiniMax-M3'
        )

        return {
          id: result.id,
          role: 'assistant',
          content: result.content,
          createdAt: Date.now(),
          metadata: {
            model: aiSettings.model || 'MiniMax-M3',
            provider: 'minimax',
            usage: result.usage,
          },
        }
      }
    } catch (error) {
      console.error('MiniMax API error, falling back to mock:', error)
    }

    let content = generateAIResponse(message)
    if (context) {
      content = `基于您提供的文档内容，我来帮您分析：\n\n${content}\n\n如需进一步帮助，请告诉我具体问题。`
    }

    return {
      id: `chat-${Date.now()}`,
      role: 'assistant',
      content,
      createdAt: Date.now(),
      metadata: {
        model: aiSettings.model || 'mock',
        provider: minimaxKey ? 'minimax' : 'mock',
      },
    }
  })

  registerHandle('ai:stream', async (event: unknown, request: unknown) => {
    const req = request as { message?: string; sessionId?: string }
    const sessionId = req.sessionId || `stream-${Date.now()}`

    const abort = new AbortController()
    AI_STREAMS.set(sessionId, { chunks: [], abort })

    const messages = [
      '正在处理您的请求',
      '分析文档结构',
      '生成内容',
      '完成',
    ]

    const sender = (event as { sender?: { send?: (ch: string, ...args: unknown[]) => void } })?.sender
    if (sender?.send) {
      for (const msg of messages) {
        await new Promise(r => setTimeout(r, 500))
        sender.send('ai:stream-chunk', { sessionId, chunk: msg, done: false })
      }
      sender.send('ai:stream-chunk', { sessionId, chunk: '', done: true })
    }

    AI_STREAMS.delete(sessionId)
    return { id: sessionId }
  })

  registerHandle('ai:stream-cancel', (_event: unknown, sessionId: unknown) => {
    const stream = AI_STREAMS.get(sessionId as string)
    if (stream) {
      stream.abort.abort()
      AI_STREAMS.delete(sessionId as string)
    }
    return { ok: true }
  })

  registerHandle('ai:web-search', async (_event: unknown, query: unknown, maxResults = 5) => {
    return [
      { title: `${query} - 结果 1`, url: 'https://example.com/1', snippet: '' },
      { title: `${query} - 结果 2`, url: 'https://example.com/2', snippet: '需要配置 Tavily API' },
    ].slice(0, maxResults as number)
  })

  registerHandle('ai:image-search', async (_event: unknown, query: unknown, maxResults = 5) => {
    return [
      { url: `https://picsum.photos/200?random=${Date.now()}`, title: `${query} 图片 1` },
      { url: `https://picsum.photos/200?random=${Date.now() + 1}`, title: `${query} 图片 2` },
    ].slice(0, maxResults as number)
  })

  registerHandle('ai:log-run-failure', () => ({ ok: true }))

  registerHandle('ai:generate-content', async (_event: unknown, request: unknown) => {
    const req = request as { type?: string; topic?: string; length?: number; style?: string }
    const { type = 'paragraph', topic = '', length = 200, style = 'formal' } = req

    const templates: Record<string, string> = {
      paragraph: `关于"${topic}"的段落内容。`,
      summary: `以下是关于"${topic}"的摘要总结。`,
      outline: `# ${topic}大纲\n\n1. 介绍\n2. 主要内容\n3. 结论`,
      introduction: `欢迎阅读关于"${topic}"的介绍。`,
      conclusion: `总结以上内容，关于"${topic}"的主要观点是...`,
    }

    return {
      id: `gen-${Date.now()}`,
      type,
      content: templates[type] || templates.paragraph,
      tokens: Math.floor(length / 4),
    }
  })

  registerHandle('ai:translate', async (_event: unknown, request: unknown) => {
    const req = request as { text?: string; from?: string; to?: string }
    return {
      id: `trans-${Date.now()}`,
      original: req.text || '',
      translated: `[${req.to || 'en'}] ${req.text || ''}`,
      from: req.from || 'auto',
      to: req.to || 'en',
    }
  })

  registerHandle('ai:summarize', async (_event: unknown, request: unknown) => {
    const req = request as { text?: string; maxLength?: number }
    const text = req.text || ''
    const maxLength = req.maxLength || 100

    return {
      id: `sum-${Date.now()}`,
      originalLength: text.length,
      summary: text.slice(0, maxLength) + (text.length > maxLength ? '...' : ''),
      keyPoints: ['要点1', '要点2', '要点3'],
    }
  })

  registerHandle('ai:qa', async (_event: unknown, request: unknown) => {
    const req = request as { question?: string; context?: string }
    return {
      id: `qa-${Date.now()}`,
      question: req.question || '',
      answer: `基于提供的内容，关于"${req.question}"的回答是...`,
      confidence: 0.85,
    }
  })

  registerHandle('ai:grammar-check', async (_event: unknown, text: unknown) => {
    return {
      id: `grammar-${Date.now()}`,
      original: text,
      corrected: text,
      errors: [],
      suggestions: [],
    }
  })

  registerHandle('ai:extract-keywords', async (_event: unknown, text: unknown) => {
    return {
      id: `kw-${Date.now()}`,
      keywords: ['关键词1', '关键词2', '关键词3'],
      score: [0.9, 0.7, 0.5],
    }
  })

  registerHandle('ai:smart-summary', async (_event: unknown, args: unknown) => {
    const { text, maxLength, type } = args as { text: string; maxLength?: number; type?: 'brief' | 'detailed' | 'bullets' }

    const length = maxLength || 200
    const summaryType = type || 'brief'

    if (summaryType === 'brief') {
      return {
        summary: text.slice(0, length) + (text.length > length ? '...' : ''),
        keyPoints: ['要点1', '要点2'],
        wordCount: text.length,
      }
    } else if (summaryType === 'bullets') {
      return {
        bullets: [
          '• 第一个要点',
          '• 第二个要点',
          '• 第三个要点',
        ],
        wordCount: text.length,
      }
    }

    return {
      summary: text.slice(0, length) + (text.length > length ? '...' : ''),
      keyPoints: ['要点1', '要点2', '要点3'],
      wordCount: text.length,
    }
  })

  registerHandle('ai:auto-translate', async (_event: unknown, args: unknown) => {
    const { text, from, to } = args as { text: string; from?: string; to: string }

    const langMap: Record<string, string> = {
      'zh': '中文', 'en': 'English', 'ja': '日本語', 'ko': '한국어',
      'fr': 'Français', 'de': 'Deutsch', 'es': 'Español', 'ru': 'Русский',
    }

    return {
      original: text,
      translated: `[${to}] ${text}`,
      from: from || 'auto',
      to,
      fromLang: langMap[from || 'auto'] || '自动检测',
      toLang: langMap[to] || to,
      confidence: 0.95,
    }
  })

  registerHandle('ai:detect-language', async (_event: unknown, text: unknown) => {
    const str = text as string
    const hasChinese = /[\u4e00-\u9fff]/.test(str)
    const hasJapanese = /[\u3040-\u309f\u30a0-\u30ff]/.test(str)
    const hasKorean = /[\uac00-\ud7af]/.test(str)

    if (hasChinese) return { language: 'zh', confidence: 0.98 }
    if (hasJapanese) return { language: 'ja', confidence: 0.90 }
    if (hasKorean) return { language: 'ko', confidence: 0.90 }

    return { language: 'en', confidence: 0.85 }
  })

  // The original index.ts registered two `ai:sentiment` handlers. The second
  // (more capable) one wins via `Map.set`; we keep that semantic.
  registerHandle('ai:sentiment', async (_event: unknown, args: unknown) => {
    const { text } = args as { text: string }

    const positiveWords = ['好', '棒', '优', '赞', '满意', '喜欢', 'good', 'great', 'excellent', 'amazing']
    const negativeWords = ['差', '坏', '糟', '不满', '讨厌', 'bad', 'poor', 'terrible', 'awful']

    const textLower = text.toLowerCase()
    const positiveCount = positiveWords.filter(w => textLower.includes(w)).length
    const negativeCount = negativeWords.filter(w => textLower.includes(w)).length

    let sentiment = 'neutral'
    let score = 0.5

    if (positiveCount > negativeCount) {
      sentiment = 'positive'
      score = Math.min(0.9, 0.5 + positiveCount * 0.1)
    } else if (negativeCount > positiveCount) {
      sentiment = 'negative'
      score = Math.max(0.1, 0.5 - negativeCount * 0.1)
    }

    return {
      sentiment,
      score,
      confidence: 0.85,
      keywords: positiveCount > negativeCount ? ['positive'] : negativeCount > positiveCount ? ['negative'] : [],
      emotions: {
        joy: sentiment === 'positive' ? 0.6 : 0.1,
        sadness: sentiment === 'negative' ? 0.5 : 0.1,
        anger: sentiment === 'negative' ? 0.3 : 0.05,
        fear: 0.05,
        surprise: 0.1,
      },
    }
  })
}
