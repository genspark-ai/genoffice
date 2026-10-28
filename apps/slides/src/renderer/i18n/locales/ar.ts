import { ar as app } from '../app/ar'
import { ar as ribbon } from '../ribbon/ar'
import { ar as panes } from '../panes/ar'
import { ar as ai } from '../ai/ar'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const ar: Dict = { ...app, ...ribbon, ...panes, ...ai }
