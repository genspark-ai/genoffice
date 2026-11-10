export declare const CLI_VERSION_ENV: string
export declare const CLI_BUNDLE: string
export declare const CLI_VERSION_BANNER_PREFIX: string
export declare function resolveCliVersion(env: NodeJS.ProcessEnv, packageVersion: string): string
export declare function cliVersionBanner(cliVersion: string): string
export declare function readBundledCliVersion(bundleText: string): string | null
export declare function findElectronRequires(bundleText: string): string[]
