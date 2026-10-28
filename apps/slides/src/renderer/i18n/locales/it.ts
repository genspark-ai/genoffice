import { it as app } from '../app/it'
import { it as ribbon } from '../ribbon/it'
import { it as panes } from '../panes/it'
import { it as ai } from '../ai/it'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const it: Dict = { ...app, ...ribbon, ...panes, ...ai }
