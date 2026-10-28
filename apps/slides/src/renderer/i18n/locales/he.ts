import { he as app } from '../app/he'
import { he as ribbon } from '../ribbon/he'
import { he as panes } from '../panes/he'
import { he as ai } from '../ai/he'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const he: Dict = { ...app, ...ribbon, ...panes, ...ai }
