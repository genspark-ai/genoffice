import { readFileSync, writeFileSync } from 'node:fs'

function patch(path, from, to) {
  const src = readFileSync(path, 'utf8')
  if (src.includes(to)) return
  if (!src.includes(from)) throw new Error(`Patch anchor not found in ${path}`)
  writeFileSync(path, src.replace(from, to))
}

// Keep a curated list of currently useful hosted NVIDIA free endpoints for GenOffice.
// These are text/agent models rather than image/video-only endpoints.
patch('packages/ai-provider/src/providers.ts',
  `models: ['nvidia/nemotron-3-ultra-550b-a55b', 'nvidia/nemotron-3-super-120b-a12b', 'nvidia/nemotron-3-nano-30b-a3b'],`,
  `models: [\n      'moonshotai/kimi-k2.6',\n      'deepseek-ai/deepseek-v4-flash',\n      'deepseek-ai/deepseek-v4-pro',\n      'qwen/qwen3.5-397b-a17b',\n      'qwen/qwen3-next-80b-a3b-instruct',\n      'minimaxai/minimax-m3',\n      'nvidia/nemotron-3-ultra-550b-a55b',\n      'nvidia/nemotron-3-super-120b-a12b',\n      'nvidia/nemotron-3-nano-30b-a3b',\n    ],`)

patch('apps/shell/src/renderer/src/Home.tsx',
  `<><option value="nvidia/nemotron-3-ultra-550b-a55b">Nemotron 3 Ultra</option><option value="nvidia/nemotron-3-super-120b-a12b">Nemotron 3 Super</option><option value="nvidia/nemotron-3-nano-30b-a3b">Nemotron 3 Nano</option></>`,
  `<><option value="moonshotai/kimi-k2.6">Kimi K2.6 — Agent / Long Tasks</option><option value="deepseek-ai/deepseek-v4-flash">DeepSeek V4 Flash — Fast Agent</option><option value="deepseek-ai/deepseek-v4-pro">DeepSeek V4 Pro — Reasoning / Agent</option><option value="qwen/qwen3.5-397b-a17b">Qwen 3.5 397B — Agent / Multimodal</option><option value="qwen/qwen3-next-80b-a3b-instruct">Qwen3 Next 80B — Fast Long Context</option><option value="minimaxai/minimax-m3">MiniMax M3 — Tool Calling</option><option value="nvidia/nemotron-3-ultra-550b-a55b">Nemotron 3 Ultra</option><option value="nvidia/nemotron-3-super-120b-a12b">Nemotron 3 Super</option><option value="nvidia/nemotron-3-nano-30b-a3b">Nemotron 3 Nano</option></>`)

console.log('Expanded NVIDIA NIM model selector with current generous/free agent-capable endpoints.')
