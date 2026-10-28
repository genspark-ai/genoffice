import { id as app } from '../app/id'
import { id as ribbon } from '../ribbon/id'
import { id as panes } from '../panes/id'
import { id as ai } from '../ai/id'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const id: Dict = { ...app, ...ribbon, ...panes, ...ai }
