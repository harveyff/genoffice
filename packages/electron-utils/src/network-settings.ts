/**
 * The user's network settings — outbound proxy and Tavily key — pushed into
 * every consumer that has to be told about them separately.
 *
 * Docs, sheets and slides all read `ai-settings.json` for the model selection,
 * but the network half of that file used to be applied by the shell alone. A
 * standalone sheets or slides window therefore ignored the configured proxy
 * entirely, and the gsk CLI children it spawns — PPT generation, image
 * generation, media analysis — dialled genspark.ai direct on the one kind of
 * network where nothing but the proxy can reach it.
 */
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'

import { session } from 'electron'

import {
  defaultAiSettings,
  normalizeProxyUrl,
  resolveAiSettings,
  type AiSettings,
  type LegacyAiSettings,
} from '@genoffice/ai-provider'
import { setGskProxyUrl, setTavilyApiKey } from '@genoffice/ai-search'

const SETTINGS_FILE = 'ai-settings.json'

/** The network-facing subset of the settings file; everything else is the provider's business. */
export interface NetworkSettings {
  proxyUrl: string
  tavilyApiKey: string
}

function readJson<T>(path: string, fallback: T): T {
  try {
    if (existsSync(path)) return JSON.parse(readFileSync(path, 'utf-8')) as T
  } catch {
    /* corrupted state file: fall back to defaults */
  }
  return fallback
}

/** Read the network fields out of `<userData>/ai-settings.json`, migrations applied. */
export function readNetworkSettings(userDataPath: string): NetworkSettings {
  const stored = readJson<Partial<AiSettings> & LegacyAiSettings>(
    join(userDataPath, SETTINGS_FILE),
    {},
  )
  const settings = resolveAiSettings(stored, defaultAiSettings())
  return { proxyUrl: settings.proxyUrl ?? '', tavilyApiKey: settings.tavilyApiKey ?? '' }
}

/** last proxy handed to undici/Chromium, so a no-op save does not churn them */
let appliedProxyUrl: string | null = null

/**
 * Route outbound traffic through the user's proxy.
 *
 * Three consumers need telling separately, which is why this is not one call:
 * main-process `fetch` runs on undici and ignores the system proxy entirely;
 * Chromium sessions carry the renderer, sign-in window and the agent browser;
 * and the gsk CLI is a child process that only sees environment variables.
 *
 * An empty url restores direct connections, so clearing the field in the
 * dialog actually turns the proxy off rather than leaving the old one wired.
 */
export async function applyProxy(rawUrl: string): Promise<void> {
  const proxyUrl = normalizeProxyUrl(rawUrl)
  if (proxyUrl === appliedProxyUrl) return
  appliedProxyUrl = proxyUrl
  setGskProxyUrl(proxyUrl)
  try {
    const { ProxyAgent, getGlobalDispatcher, setGlobalDispatcher, Agent } = await import('undici')
    if (proxyUrl) {
      setGlobalDispatcher(new ProxyAgent(proxyUrl))
    } else if (getGlobalDispatcher() instanceof ProxyAgent) {
      setGlobalDispatcher(new Agent())
    }
  } catch (err) {
    console.warn('[proxy] failed to set undici dispatcher:', err)
  }
  try {
    // Empty proxy = direct. On Olares/Kasm the Chromium "system" proxy often
    // points at cluster egress (e.g. *.frp.olares.com) and breaks local LLMs.
    await session.defaultSession.setProxy(proxyUrl ? { proxyRules: proxyUrl } : { mode: 'direct' })
  } catch (err) {
    console.warn('[proxy] failed to set session proxy:', err)
  }
  console.log(
    proxyUrl
      ? // strip user:pass before logging
        `[proxy] outbound via ${proxyUrl.replace(/\/\/[^@/]*@/, '//***@')}`
      : '[proxy] direct',
  )
}

/** current proxy, for callers that need to pass it on (e.g. the agent browser) */
export function currentProxyUrl(): string {
  return appliedProxyUrl ?? ''
}

/**
 * Push the network-facing settings into the modules that hold them as process
 * state. Called on load and after every save so a settings change takes effect
 * without a restart, the same way the provider switch does.
 */
export function applyNetworkSettings(settings: Partial<NetworkSettings>): void {
  setTavilyApiKey(settings.tavilyApiKey ?? '')
  void applyProxy(settings.proxyUrl ?? '')
}

/**
 * Load the persisted network settings at startup. Returns true when
 * `ai-settings.json` exists, which tells app bootstraps to skip env-var /
 * system-proxy detection: a saved file (even with an empty proxy field) is an
 * explicit "use direct" choice on machines whose system proxy would otherwise
 * hijack local LLM endpoints.
 */
export function bootstrapNetworkSettings(userDataPath: string): boolean {
  const settingsPath = join(userDataPath, SETTINGS_FILE)
  const settings = readNetworkSettings(userDataPath)
  applyNetworkSettings(settings)
  return existsSync(settingsPath)
}
