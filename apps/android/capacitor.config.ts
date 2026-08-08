import type { CapacitorConfig } from '@capacitor/cli'

const config: CapacitorConfig = {
  appId: 'com.genoffice.mobile',
  appName: 'GenOffice',
  webDir: 'dist',
  bundledWebRuntime: false,
  android: {
    backgroundColor: '#ffffff',
    allowMixedContent: false,
  },
  server: {
    androidScheme: 'https',
  },
}

export default config
