import { pt as app } from '../app/pt'
import { pt as ribbon } from '../ribbon/pt'
import { pt as panes } from '../panes/pt'
import { pt as ai } from '../ai/pt'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const pt: Dict = { ...app, ...ribbon, ...panes, ...ai }
