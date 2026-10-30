import { z } from 'zod'
import type { DesktopControlRequest } from '@genoffice/cli/control-protocol'
import type { McpToolDefinition } from '../mcp-server'
import { runDesktopAction, type DesktopControlDeps } from '../../desktop-control'

/**
 * The desktop domain as an in-app MCP tool: the same actions the genoffice
 * CLI reaches over the control socket (`genoffice desktop …`), backed by the
 * same live stores, so the shell agent and an external script never disagree
 * about the desktop's state.
 */

export function createDesktopTools(control: DesktopControlDeps): McpToolDefinition[] {
  return [
    {
      name: 'desktop',
      description:
        'Read and drive the GenOffice Home desktop. Actions: "get-state" returns the current ' +
        'wallpaper, the icon layout summary, the todo list and the open windows/tabs in one ' +
        'snapshot; "set-wallpaper" (wallpaper: "default", a builtin name like "teal-deep", or an ' +
        'absolute path to a png/jpg/webp image, which is imported first) changes the desktop ' +
        'wallpaper; "todo-add" (text) adds one item to the Home todo widget; "todo-list" returns ' +
        'the widget\'s items; "windows-list" enumerates the open editor tabs and detached windows.',
      inputSchema: {
        action: z
          .enum(['get-state', 'set-wallpaper', 'todo-add', 'todo-list', 'windows-list'])
          .describe('what to do'),
        wallpaper: z
          .string()
          .optional()
          .describe('"set-wallpaper" only: "default", a builtin wallpaper name, or an image path'),
        text: z.string().optional().describe('"todo-add" only: the todo item\'s text'),
      },
      handler: async (args) => {
        const action = String(args.action ?? '')
        let request: DesktopControlRequest
        switch (action) {
          case 'get-state':
          case 'todo-list':
          case 'windows-list':
            request = { cmd: 'desktop', action }
            break
          case 'set-wallpaper': {
            const wallpaper = typeof args.wallpaper === 'string' ? args.wallpaper.trim() : ''
            if (!wallpaper) throw new Error('"set-wallpaper" needs a wallpaper argument')
            request = { cmd: 'desktop', action, wallpaper }
            break
          }
          case 'todo-add': {
            const text = typeof args.text === 'string' ? args.text.trim() : ''
            if (!text) throw new Error('"todo-add" needs a text argument')
            request = { cmd: 'desktop', action, text }
            break
          }
          default:
            throw new Error(
              `unknown action "${action}" — use "get-state", "set-wallpaper", "todo-add", "todo-list" or "windows-list"`,
            )
        }
        return runDesktopAction(control, request)
      },
    },
  ]
}
