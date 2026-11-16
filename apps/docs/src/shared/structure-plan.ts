/** A bounded explanation of one provisional edit, never an authorization grant. */
export interface DocsStructurePlan {
  required: string[]
  missing: string[]
  extra: string[]
  wouldRemove: Array<{ id: string; kind: string; from: number; to: number }>
  wouldPreserve: Array<{ id: string; kind: string; from: number; to: number }>
  complete: boolean
  counts: {
    required: number
    missing: number
    extra: number
    wouldRemove: number
    wouldPreserve: number
  }
}

export function bindStructurePlan(
  plan: DocsStructurePlan,
  documentId: string,
  revision: number,
): DocsStructurePlan {
  const bind = (id: string) => `structure:${documentId}:${revision}:${id}`
  return {
    ...plan,
    required: plan.required.map(bind),
    missing: plan.missing.map(bind),
    extra: plan.extra.map(bind),
    wouldRemove: plan.wouldRemove.map((item) => ({ ...item, id: bind(item.id) })),
    wouldPreserve: plan.wouldPreserve.map((item) => ({ ...item, id: bind(item.id) })),
  }
}
