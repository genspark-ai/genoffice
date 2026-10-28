import { ja as app } from '../app/ja'
import { ja as ribbon } from '../ribbon/ja'
import { ja as panes } from '../panes/ja'
import { ja as ai } from '../ai/ja'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const ja: Dict = { ...app, ...ribbon, ...panes, ...ai }
