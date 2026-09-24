/** The Univer facade the renderer owns, shared with the script host (issue #815). */
let univerAPI: unknown = null

export function setUniverAPI(api: unknown): void {
  univerAPI = api
}

export function getUniverAPI(): unknown {
  return univerAPI
}
