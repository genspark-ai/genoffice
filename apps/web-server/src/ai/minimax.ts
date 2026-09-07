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
  stream = false
): Promise<{ content: string; id: string; usage?: MiniMaxResponse['usage'] }> {
  if (!apiKey) {
    throw new Error('MiniMax API key not configured')
  }

  const response = await fetch(`${MINIMAX_API_URL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
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

  const data = await response.json() as MiniMaxResponse
  return {
    content: data.choices[0]?.message?.content || '',
    id: data.id,
    usage: data.usage,
  }
}

export function generateAIResponse(message: string): string {
  if (!message) return '请输入内容'
  return `这是 AI 助手的回复。您发送的消息是: "${message}"。\n\n我可以帮助您:\n1. 编辑和格式化文档\n2. 创建表格和幻灯片\n3. 回答问题和提供建议\n4. 搜索和整理信息\n\n请告诉我您需要什么帮助?`
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
