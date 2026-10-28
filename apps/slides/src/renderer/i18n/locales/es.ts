import { es as app } from '../app/es'
import { es as ribbon } from '../ribbon/es'
import { es as panes } from '../panes/es'
import { es as ai } from '../ai/es'
import type { Dict } from '../strings'

/** one lazy chunk per locale: the four domain shards merged into the UI dictionary */
export const es: Dict = { ...app, ...ribbon, ...panes, ...ai }
