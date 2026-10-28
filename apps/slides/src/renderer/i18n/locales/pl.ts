import { pl as app } from '../app/pl'
import { pl as ribbon } from '../ribbon/pl'
import { pl as panes } from '../panes/pl'
import { pl as ai } from '../ai/pl'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const pl: Dict = { ...app, ...ribbon, ...panes, ...ai }
