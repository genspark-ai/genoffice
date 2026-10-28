import { ms as app } from '../app/ms'
import { ms as ribbon } from '../ribbon/ms'
import { ms as panes } from '../panes/ms'
import { ms as ai } from '../ai/ms'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const ms: Dict = { ...app, ...ribbon, ...panes, ...ai }
