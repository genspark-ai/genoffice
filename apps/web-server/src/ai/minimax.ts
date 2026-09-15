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
