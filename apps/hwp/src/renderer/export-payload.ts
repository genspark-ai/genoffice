export interface HmlBlocker {
  code: string
  xmlPath: string
  message: string
}

export interface StudioExport {
  exportHwp(): Promise<Uint8Array>
  exportHwpx(): Promise<Uint8Array>
  exportHml(): Promise<Uint8Array>
  getHmlSaveState(): Promise<{ hmlSavable: boolean; blockers: HmlBlocker[] }>
}

export interface StudioPayload {
  hwp: Uint8Array
  hwpx: Uint8Array
  hml?: Uint8Array
  hmlBlockers?: HmlBlocker[]
}

/** Prefer getHmlSaveState() over catching exportHml() — SDK documents that. */
export async function exportStudioPayload(studio: StudioExport): Promise<StudioPayload> {
  const [hwp, hwpx] = await Promise.all([studio.exportHwp(), studio.exportHwpx()])
  let state: { hmlSavable: boolean; blockers: HmlBlocker[] }
  try {
    state = await studio.getHmlSaveState()
  } catch {
    return { hwp, hwpx }
  }
  if (!state.hmlSavable) return { hwp, hwpx, hmlBlockers: state.blockers }
  try {
    return { hwp, hwpx, hml: await studio.exportHml() }
  } catch {
    return { hwp, hwpx, hmlBlockers: state.blockers }
  }
}

export function hmlUnavailableMessage(blockers?: HmlBlocker[]): string {
  const detail = blockers
    ?.map((item) => item.message)
    .filter(Boolean)
    .join(' ')
  return detail
    ? `This document cannot be saved as HML. ${detail}`
    : 'This document cannot be saved as HML.'
}
