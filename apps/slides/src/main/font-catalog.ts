/**
 * The suite-wide downloadable font catalog lives in @genoffice/electron-utils so
 * every app downloads from the same mirror; slides re-exports it for its store.
 */
export * from '@genoffice/electron-utils/font-catalog'
