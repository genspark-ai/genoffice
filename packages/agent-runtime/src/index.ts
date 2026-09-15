/**
 * @genoffice/agent-runtime — public surface
 */

export {
  ReactUIAdapter,
  type DialogRequest,
  type DialogKind,
  type NotificationItem,
  type NotificationKind,
  type SelectDialogRequest,
  type ConfirmDialogRequest,
  type InputDialogRequest,
} from "./ui-adapter";

export {
  createOfficeSession,
  type OfficeSession,
  type OfficeSessionOptions,
} from "./session";

export {
  PiSessionProvider,
  useOfficeSession,
  usePiSession,
  usePiAgentSession,
  useUiAdapter,
  usePiDialogs,
  usePiNotifications,
  usePiStatuses,
  type PiSessionProviderProps,
} from "./provider";

export {
  PiDialogHost,
  NotificationToaster,
  PiStatusBar,
  type PiDialogHostProps,
  type NotificationToasterProps,
} from "./components";

export {
	createBenchmark,
	createResponseCache,
	DEFAULT_CACHE_TTL_MS,
	DEFAULT_CACHE_MAX_ENTRIES,
	PERFORMANCE_TARGETS,
	recordTiming,
	summarizeBenchmark,
	type Benchmark,
	type BenchmarkSummary,
	type CacheEntry,
	type ResponseCache,
	type ResponseCacheOptions,
	type TimingSample,
} from "./performance";
