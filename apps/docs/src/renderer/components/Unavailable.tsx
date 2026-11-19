import type { ReactNode } from 'react'
import { isUnavailable } from '../../shared/unavailable-entries'

/** leaves its content out while `id` is listed in `shared/unavailable-entries.ts` */
export function Unavailable({ id, children }: { id: string; children: ReactNode }) {
  return isUnavailable(id) ? null : <>{children}</>
}
