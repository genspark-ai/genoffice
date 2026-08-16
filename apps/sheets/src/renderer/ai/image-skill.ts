import type { AgentSkill } from '@genoffice/agent-core'
import { t } from '../i18n/locale'

/**
 * Image acquisition AgentSkill: image_search (shared main-process channel, same
 * source as docs/slides) and generate_image (sheets-owned Genspark channel).
 * Both return a URL; placement happens through the normal propose_operations
 * add_image path, which downloads the URL in the main process on apply.
 */

const IMAGES_SYSTEM_PROMPT = `## Images
- image_search finds real web images (returns direct imageUrl entries); generate_image creates an illustration with AI when no suitable real image exists or the user explicitly wants generated art.
- To place an image on a sheet, pass the URL to propose_operations {op:"add_image", sheetId, path:"<https url>", anchorCell} — field details in guide charts. The image anchors at that cell and is written into the file on save (imported xlsx only).
- Only insert images the user asked for; data correctness always outranks decoration.`

export function createImageSkill(): AgentSkill {
  return {
    id: 'images',
    systemPrompt: IMAGES_SYSTEM_PROMPT,
    tools: [
      {
        name: 'image_search',
        description:
          'Search the web for images. Returns a numbered list of direct imageUrl entries with pixel sizes; ' +
          'pick one and insert it with propose_operations add_image (path = the URL).',
        inputSchema: {
          type: 'object',
          properties: {
            query: { type: 'string', description: 'Image search keywords (English works better)' },
            maxResults: { type: 'integer', description: 'Maximum number of results, default 8' },
          },
          required: ['query'],
        },
      },
      {
        name: 'generate_image',
        description:
          'Generate an image with AI from a text prompt (Genspark account required). Returns a URL to insert ' +
          'with propose_operations add_image. Use for illustrations/decorative art; prefer image_search for real-world subjects.',
        inputSchema: {
          type: 'object',
          properties: {
            prompt: {
              type: 'string',
              description: 'What to draw — subject, style, composition (English works better)',
            },
            aspectRatio: {
              type: 'string',
              description: 'Aspect ratio like "1:1", "16:9", "4:3"; default 1:1',
            },
          },
          required: ['prompt'],
        },
      },
    ],
    executeTool: async (call) => {
      if (call.name === 'image_search') {
        const query = String(call.input.query ?? '').trim()
        if (!query) {
          return {
            output: 'query must not be empty',
            isError: true,
            summary: t('aiToolImageSearch'),
          }
        }
        const result = await window.desktopApi.imageSearch(
          query,
          Number(call.input.maxResults) || 8,
        )
        // A backend failure must not read as an empty gallery — the model
        // would fabricate image choices
        if (result.method === 'error') {
          return {
            output: `image search failed (service error, not an empty result — you may retry): ${result.error ?? 'unknown error'}`,
            isError: true,
            summary: t('aiToolImageSearch'),
          }
        }
        const lines = result.images.map(
          (image, index) =>
            `${index + 1}. ${image.title || '(untitled)'} [${image.width ?? '?'}x${image.height ?? '?'}]\n   ${image.imageUrl}`,
        )
        return {
          output: lines.join('\n') || '(no images)',
          mutated: false,
          summary: t('aiToolImageSearchDone', { query, count: result.images.length }),
        }
      }
      if (call.name === 'generate_image') {
        const prompt = String(call.input.prompt ?? '').trim()
        if (!prompt) {
          return { output: 'prompt must not be empty', isError: true, summary: t('aiToolGenImage') }
        }
        const aspectRatio = String(call.input.aspectRatio ?? '').trim()
        const result = await window.desktopApi.generateImage({
          prompt,
          ...(aspectRatio ? { aspectRatio } : {}),
        })
        if (!result.url) {
          return {
            output: `image generation failed: ${result.error ?? 'unknown error'}`,
            isError: true,
            summary: t('aiToolGenImage'),
          }
        }
        return {
          output: `Image generated: ${result.url}\nInsert it with propose_operations {op:"add_image", path:"${result.url}", ...}.`,
          mutated: false,
          summary: t('aiToolGenImageDone'),
        }
      }
      return { output: `Unknown tool: ${call.name}`, isError: true, summary: call.name }
    },
  }
}
