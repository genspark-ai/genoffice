/// Local replacement for the `createUniver` helper from `@univerjs/presets`.
///
/// The `@univerjs/presets` meta-package bundles every preset — including the
/// advanced/collaboration ones, which transitively pull 27 proprietary
/// `@univerjs-pro/*` packages (npm license "None") into the lockfile and
/// node_modules. We only ever use the free Apache-2.0 presets, so we depend
/// on the individual `@univerjs/preset-sheets-*` packages and reproduce the
/// small composition helper here (collaboration mode intentionally omitted).
import { LogLevel, Univer } from '@univerjs/core'
import type { DependencyOverride, IUniverConfig, Plugin, PluginCtor } from '@univerjs/core'
import { FUniver } from '@univerjs/core/lib/facade'

type PluginEntry = PluginCtor<Plugin> | [PluginCtor<Plugin>, ConstructorParameters<PluginCtor<Plugin>>[0]]

/** A collection of plugins and their default configs (same shape the preset packages export). */
export interface IPreset {
  plugins: PluginEntry[]
}

export interface CreateUniverOptions extends Partial<IUniverConfig> {
  presets: Array<IPreset | [IPreset, { lazy?: boolean }]>
  plugins?: PluginEntry[]
  override?: DependencyOverride
}

export function createUniver(options: CreateUniverOptions): { univer: Univer; univerAPI: FUniver } {
  const { presets, plugins, override = [], ...univerConfig } = options
  const univer = new Univer({ logLevel: LogLevel.WARN, ...univerConfig, override })

  // Later registrations win across presets (same dedupe semantics as the
  // upstream helper); explicitly listed plugins must not collide with any
  // preset-provided plugin.
  const registry = new Map<string, { plugin: PluginCtor<Plugin>; options: unknown }>()
  for (const entry of presets) {
    const preset = Array.isArray(entry) ? entry[0] : entry
    for (const pluginEntry of preset.plugins) {
      const [plugin, pluginOptions] = Array.isArray(pluginEntry) ? pluginEntry : [pluginEntry, undefined]
      registry.delete(plugin.pluginName)
      registry.set(plugin.pluginName, { plugin, options: pluginOptions })
    }
  }
  for (const pluginEntry of plugins ?? []) {
    const [plugin, pluginOptions] = Array.isArray(pluginEntry) ? pluginEntry : [pluginEntry, undefined]
    if (registry.has(plugin.pluginName)) {
      throw new Error(
        `Plugin ${plugin.pluginName} already registered by presets or other ways! `
        + 'Repeated registration may cause potential problems, please check your code.',
      )
    }
    registry.set(plugin.pluginName, { plugin, options: pluginOptions })
  }
  for (const { plugin, options: pluginOptions } of registry.values()) {
    univer.registerPlugin(plugin, pluginOptions)
  }

  return { univer, univerAPI: FUniver.newAPI(univer) }
}
