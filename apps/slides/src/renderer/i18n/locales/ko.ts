import { ko as app } from '../app/ko'
import { ko as ribbon } from '../ribbon/ko'
import { ko as panes } from '../panes/ko'
import { ko as ai } from '../ai/ko'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const ko: Dict = { ...app, ...ribbon, ...panes, ...ai }
