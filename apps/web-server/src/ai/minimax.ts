/**
 * MiniMax provider client used by `ai:chat` and other channels.
 *
 * This is a placeholder provider kept here so we don't lose the surface
 * during the Phase 1.1 refactor. Phase 1.3 (separate issue) will route
 * these calls through `packages/ai-provider`.
 */

const MINIMAX_API_URL = 'https://api.minimax.chat/v1'

export interface MiniMaxMessage {
  role: 'user' | 'assistant' | 'system'
  content: string
}

interface MiniMaxResponse {
  id: string
  choices: Array<{
    message: {
      role: string
      content: string
    }
    finish_reason: string
  }>
  usage?: {
    prompt_tokens: number
    completion_tokens: number
    total_tokens: number
  }
}

export async function callMiniMax(
  apiKey: string,
  messages: MiniMaxMessage[],
  model = 'MiniMax-M3',
  stream = false,
): Promise<{ content: string; id: string; usage?: MiniMaxResponse['usage'] }> {
  if (!apiKey) {
    throw new Error('MiniMax API key not configured')
  }

  const response = await fetch(`${MINIMAX_API_URL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model,
      messages,
      stream,
      max_tokens: 2048,
      temperature: 0.7,
    }),
  })

  if (!response.ok) {
    const error = await response.text()
    throw new Error(`MiniMax API error: ${response.status} - ${error}`)
  }

  const data = (await response.json()) as MiniMaxResponse
  return {
    content: data.choices[0]?.message?.content || '',
    id: data.id,
    usage: data.usage,
  }
}

export function generateAIResponse(message: string): string {
  if (!message) return '请输入内容'

  const trimmed = message.trim()
  const lower = trimmed.toLowerCase()

  /* ── Pattern-aware mock responses ─────────────────────────────────
   * The web build ships without a configured MiniMax API key, so the
   * AI chat handler falls back to a deterministic local response.
   * Detect common intent patterns and return a contextually relevant
   * acknowledgement so the chat feels functional instead of generic. */
  if (/^(translate|翻译|翻譯)/i.test(trimmed)) {
    return `好的，我来帮您翻译。原文: "${trimmed.replace(/^(translate|翻译|翻譯)\s*[:：]?\s*/i, '')}"\n\n请选择目标语言:\n- 中文（简体）\n- English\n- 日本語\n- 한국어\n\n提示：连接真实的 MiniMax API 后会返回高质量翻译结果。`
  }
  if (/^(summarize|总结|摘要)/i.test(trimmed)) {
    return `我来帮您总结以下内容:\n\n${trimmed.replace(/^(summarize|总结|摘要)\s*[:：]?\s*/i, '').slice(0, 200)}\n\n【摘要】这是一段示例文本的关键内容。\n【关键词】核心要点 1, 核心要点 2, 核心要点 3\n\n需要更详细的摘要吗？`
  }
  if (/^(explain|解释|说明)/i.test(trimmed)) {
    return `让我为您解释 "${trimmed.replace(/^(explain|解释|说明)\s*[:：]?\s*/i, '').slice(0, 60)}"：\n\n这是一个概念或主题。可以从以下几个层面展开:\n1. 基本定义\n2. 核心原理\n3. 实际应用\n4. 相关案例\n\n请问您希望深入了解哪个方面？`
  }
  if (/^(write|生成|写|创作)/i.test(trimmed)) {
    return `我来为您创作内容:\n\n主题: ${trimmed.replace(/^(write|生成|写|创作)\s*[:：]?\s*/i, '').slice(0, 60)}\n\n【生成结果】\n这是一个示例内容。实际接入 MiniMax API 后会根据您的需求生成完整的文档、邮件、报告或其他文本内容。\n\n是否需要调整风格、长度或格式？`
  }
  if (
    lower === 'hi' ||
    lower === 'hello' ||
    lower === '你好' ||
    /^(hi|hello|你好)[!！.。\s]*$/.test(lower)
  ) {
    return '你好！我是 GenOffice AI 助手，可以帮您处理文档、表格、演示、PDF 等任务。请问今天需要什么帮助？'
  }
  if (/[?？]/.test(trimmed)) {
    return `您问的问题很有价值。让我来分析:\n\n问题: ${trimmed.slice(0, 100)}\n\n根据当前可用的本地分析，这是一个需要进一步处理的问题。\n\n建议的答案方向:\n- 考虑相关上下文\n- 检查文档关联\n- 必要时查询资料\n\n您希望我以哪种方式回应？`
  }

  /* Default contextual echo */
  return `收到您的消息: "${trimmed.slice(0, 80)}"\n\n我可以帮您:\n1. 翻译（请用"翻译: ..."开头）\n2. 总结（请用"总结: ..."开头）\n3. 解释（请用"解释: ..."开头）\n4. 创作（请用"生成: ..."开头）\n\n也可以直接提问，我会尽力回答。`
}

export function generateAgentResponse(messages: unknown[]): string {
  const lastMsg = messages[messages.length - 1] as { text?: string } | undefined
  const userMessage = lastMsg?.text || 'Hello'

  const responses = [
    `我收到了您的消息: "${userMessage}"。`,
    `正在分析您的请求...`,
    `根据我的理解，您需要帮助处理这个任务。`,
    `我可以帮助您完成文档编辑、表格处理、幻灯片制作等工作。`,
    `请问还有什么其他需要帮助的吗？`,
  ]

  return responses.join(' ')
}
