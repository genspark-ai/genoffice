/**
 * @genoffice/mail-protocols: pure-TS mail protocol layer. No Electron,
 * no native deps. Driver lanes (IMAP/SMTP, Gmail, Graph) add pure-TS
 * MailProvider implementations importing these interfaces.
 */
export type {
  MailFolder,
  MailMessageRef,
  MailFlags,
  MailEnvelope,
  MailBody,
  StagedAttachment,
  SendRequest,
  MailQuery,
  MailEvent,
  MailProvider,
  MailIntent,
  SyncState,
  MailStore,
} from './types'
export {
  parseAddress,
  parseAddressList,
  formatAddress,
  formatAddressList,
  decodeDisplayName,
  toEnvelopeStrings,
  type MailAddress,
} from './addresses'
export {
  parseMimeMessage,
  makeSnippet,
  singleLine,
  stripTags,
  detectRemoteRefs,
  type ParsedMail,
  type ParsedAttachment,
} from './mime'
export {
  OUTGOING_WARN_BYTES,
  INLINE_PREVIEW_CAP_BYTES,
  safeFilename,
  contentAddress,
  checkAttachmentSize,
  stageAttachment,
  type StageInput,
} from './attachments'
export {
  decideSync,
  nextIntent,
  resolveIntentAttempt,
  intentKey,
  type SyncDecision,
  type SyncSignals,
  type IntentAttempt,
  type IntentStep,
  type IntentResolution,
} from './sync'
export { matchesQuery } from './query.js'
export * from './imap-smtp-driver.js'
export * from './imap-transport.js'
export * from './imap-fakes.js'
export {
  GmailDriver,
  GmailAuthExpiredError,
  GmailSendLimitError,
  GMAIL_MODIFY_SCOPE,
  GMAIL_MAX_MESSAGE_BYTES,
  GMAIL_WARN_BYTES,
  GMAIL_MAX_RECIPIENTS_PER_MESSAGE,
  isHistoryExpiredError,
  isInvalidGrantError,
  mapGmailError,
  countSendRecipients,
  estimateSendBytes,
  checkSendLimits,
  gmailLabelToFolder,
  gmailLabelsToFlags,
  decodeGmailBody,
  extractGmailBody,
  gmailMessageToEnvelope,
  buildGmailQuery,
  buildRawMessage,
  encodeRawMessage,
  parseSentRaw,
  type GmailHeader,
  type GmailPayload,
  type GmailMessage,
  type GmailLabel,
  type GmailAuthProvider,
  type GmailTransport,
  type GmailSyncHooks,
  type GmailRawBody,
  type SendLimitVerdict,
  type RawAttachmentInput,
  type RawMessageInput,
  type GmailDriverOptions,
} from './gmail'
export { createGmailTransport } from './gmail-transport'
export {
  GRAPH_MAIL_SCOPES,
  GraphServiceError,
  toGraphError,
  isDeltaExpiredError,
  isAuthFailure,
  isInvalidGrant,
  createGraphTransport,
  type GraphAuthProvider,
  type GraphRequestInit,
  type GraphTransport,
} from './graph-transport'
export {
  GraphDriver,
  GraphDriverError,
  DEFAULT_POLL_INTERVAL_MS as GRAPH_DEFAULT_POLL_INTERVAL_MS,
  MAX_POLL_BACKOFF_MS as GRAPH_MAX_POLL_BACKOFF_MS,
  GRAPH_WARN_TOTAL_BYTES,
  GRAPH_MAX_TOTAL_BYTES,
  GRAPH_MAX_ATTACHMENT_BYTES,
  GRAPH_MAX_RECIPIENTS,
  type GraphDriverErrorCode,
  type GraphSendCaps,
  type GraphDriverConfig,
  type GraphMessage,
  type GraphFolder,
  type DeltaSyncResult,
  type SendSizeVerdict,
} from './graph-driver'
