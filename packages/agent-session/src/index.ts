export {
	createElectronSessionBackend,
	DEFAULT_DATABASE_FILENAME,
	DEFAULT_USER_HOME_DIR,
	resolveDatabasePath,
	type ElectronSessionBackend,
	type ElectronSessionBackendOptions,
} from "./sqlite";

export {
	createWebSessionBackend,
	DEFAULT_DATABASE_NAME,
	DEFAULT_STORE_NAME,
	fromJsonl,
	toJsonl,
	type JsonlSessionEntry,
	type WebSessionBackend,
	type WebSessionBackendOptions,
	type WebSessionMetadata,
} from "./indexeddb";
