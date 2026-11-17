import type { ClipboardSuggestApi } from '../clipboard-suggest-api'
import type { ClipboardHistoryApi } from '../clipboard-history-api'
import type { AgyOcrApi } from './agy-ocr'
import type { DocumentIndexApi } from './document-index-api'
import type { HomeChatApi } from './home-chat-types'
import type { IndexingModeApi } from './indexing-mode'
import type { AiInstructionsApi } from './ai-instructions-meta'
import type { StorageStartupApi } from './storage-startup'
import type { LiveBackgroundApi } from './live-background'
import type { LocalVlmApi } from './local-vlm'

/** Every fork-only method on window.aiOffice; HomeApi extends this so upstream stays untouched. */
export interface ForkHomeApi
  extends
    ClipboardSuggestApi,
    ClipboardHistoryApi,
    HomeChatApi,
    DocumentIndexApi,
    IndexingModeApi,
    AgyOcrApi,
    AiInstructionsApi,
    StorageStartupApi,
    LiveBackgroundApi,
    LocalVlmApi {}
