/**
 * Preload for the settings window. Single-file bundle by necessity (a shared
 * runtime chunk cannot be loaded from a preload), so it only imports types and
 * the channel constants.
 */
import { contextBridge, ipcRenderer } from 'electron'
import type {
  AgentRules,
  AiModelSettings,
  AiProviderTestResult,
  SettingsApi,
  SettingsSnapshot,
  UserSkill,
} from '../shared/settings-api'
import { SETTINGS_CHANNELS } from '../shared/settings-api'

const api: SettingsApi = {
  async load() {
    return (await ipcRenderer.invoke(SETTINGS_CHANNELS.load)) as SettingsSnapshot
  },
  async saveAi(settings) {
    return (await ipcRenderer.invoke(SETTINGS_CHANNELS.saveAi, settings)) as AiModelSettings
  },
  async testProvider(settings) {
    const result: unknown = await ipcRenderer.invoke(SETTINGS_CHANNELS.testProvider, {
      ...settings,
      mode: 'custom',
    })
    const value = result as AiProviderTestResult | undefined
    return value?.ok === true
      ? { ok: true }
      : { ok: false, error: value?.error ?? 'Request failed' }
  },
  async saveRules(rules) {
    return (await ipcRenderer.invoke(SETTINGS_CHANNELS.saveRules, rules)) as AgentRules
  },
  async saveSkill(skill) {
    return (await ipcRenderer.invoke(SETTINGS_CHANNELS.saveSkill, skill)) as UserSkill
  },
  async deleteSkill(id) {
    await ipcRenderer.invoke(SETTINGS_CHANNELS.deleteSkill, id)
  },
  async deleteMemory(id) {
    await ipcRenderer.invoke(SETTINGS_CHANNELS.deleteMemory, id)
  },
  async importSkillFiles() {
    const result: unknown = await ipcRenderer.invoke(SETTINGS_CHANNELS.importSkillFiles)
    return Array.isArray(result) ? (result as UserSkill[]) : []
  },
  async importSkillContents(files) {
    const result: unknown = await ipcRenderer.invoke(SETTINGS_CHANNELS.importSkillContents, files)
    return Array.isArray(result) ? (result as UserSkill[]) : []
  },
  async setLanguage(lang) {
    await ipcRenderer.invoke(SETTINGS_CHANNELS.setLanguage, lang)
  },
  async setUpdateChannel(channel) {
    if (channel !== 'stable' && channel !== 'beta') throw new Error('Invalid update channel.')
    await ipcRenderer.invoke(SETTINGS_CHANNELS.setUpdateChannel, channel)
  },
  async accountLogin() {
    return (await ipcRenderer.invoke(SETTINGS_CHANNELS.accountLogin)) === true
  },
  async accountLogout() {
    await ipcRenderer.invoke(SETTINGS_CHANNELS.accountLogout)
  },
  async close() {
    await ipcRenderer.invoke(SETTINGS_CHANNELS.close)
  },
}

contextBridge.exposeInMainWorld('genofficeSettings', api)

/** main pushes this when something changed outside the window (e.g. sign-in) */
contextBridge.exposeInMainWorld('genofficeSettingsEvents', {
  onChanged(handler: () => void) {
    const listener = () => handler()
    ipcRenderer.on(SETTINGS_CHANNELS.changed, listener)
    return () => ipcRenderer.removeListener(SETTINGS_CHANNELS.changed, listener)
  },
})
