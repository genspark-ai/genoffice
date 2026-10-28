import { ru as app } from '../app/ru'
import { ru as ribbon } from '../ribbon/ru'
import { ru as panes } from '../panes/ru'
import { ru as ai } from '../ai/ru'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const ru: Dict = { ...app, ...ribbon, ...panes, ...ai }
