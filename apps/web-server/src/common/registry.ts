/**
 * Single IPC handler registry shared by every capability module.
 *
 * `apps/web-server/src/index.ts` instantiates one registry at startup; each
 * module imports `registerHandle` and pushes its handlers at module load
 * time. The HTTP layer then walks `handlers` to dispatch `/api/ipc/:channel`
 * requests and the boot banner reports `handlers.size`.
 */

export type IpcHandler = (event: unknown, ...args: unknown[]) => unknown

const handlers = new Map<string, IpcHandler>()

export function registerHandle(channel: string, handler: IpcHandler): void {
  handlers.set(channel, handler)
}

export function getHandler(channel: string): IpcHandler | undefined {
  return handlers.get(channel)
}

export function listChannels(): string[] {
  return [...handlers.keys()].sort()
}

export function handlerCount(): number {
  return handlers.size
}
