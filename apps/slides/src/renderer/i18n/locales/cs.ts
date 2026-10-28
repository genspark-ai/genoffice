import { cs as app } from '../app/cs'
import { cs as ribbon } from '../ribbon/cs'
import { cs as panes } from '../panes/cs'
import { cs as ai } from '../ai/cs'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const cs: Dict = { ...app, ...ribbon, ...panes, ...ai }
