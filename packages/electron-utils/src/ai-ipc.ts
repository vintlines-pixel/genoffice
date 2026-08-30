/**
 * Shared registration for the generic AI IPC channels ('ai:*'): settings
 * persistence, the streaming/chat proxy (main process avoids renderer CORS),
 * Genspark login status, web/image search, and remote image fetch.
 *
 * One implementation, three call sites: docs-main registers it shell-wide
 * (covering docs/pdf/markdown renderers), slides and sheets register the same
 * set when running standalone. App-owned channels (per-app generate-image,
 * slides media/style channels) stay with their apps and reuse
 * generateImageWithOwnApi for the BYOK-first image path.
 *
 * Per-caller variation is expressed as options, not forks:
 * - errors: localized error text providers
 * - streamStopReason: include the normalized stop reason on the done chunk
 * - onInvoke: sender gate (sheets validates the webContents session)
 * - parseSettings: validate/transform settings before persisting (sheets zod)
 * - healLegacyGskTools: one-time rewrite of a legacy genspark+tools-off file
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname } from 'node:path'
import { ipcMain, net, shell, type IpcMainInvokeEvent } from 'electron'
import {
  AiCreditsError,
  AiTimeoutError,
  activeProvider,
  chatForProvider,
  cloudToolsEnabled,
  defaultAiSettings,
  hasImageApiConfig,
  isAiNetworkError,
  resolveAiSettings,
  setRescueFetch,
  streamForProvider,
  type AiChatRequest,
  type AiSettings,
  type AiStreamChunk,
  type AiStreamRequest,
  type GenSparkAccountStatus,
  type LegacyAiSettings,
} from '@genoffice/ai-provider'
import {
  ensureGenofficeLogin,
  gskApiKey,
  gskLoginInfo,
  hasGskAuth,
  imageSearch,
  openaiCompatibleGenerateImage,
  takeStashedImage,
  webSearch,
} from '@genoffice/ai-search'
import { fetchRemoteImage } from './remote-image'

export interface AiIpcErrorTexts {
  /** provider === 'genspark' but no gsk key is available */
  gskNotLoggedIn(): string
  noApiKey(provider: string): string
  noModel(): string
}

export interface SharedAiIpcOptions {
  settingsPath(): string
  errors: AiIpcErrorTexts
  /** include the normalized stop reason on the done chunk (docs parity) */
  streamStopReason?: boolean
  /** called before handling session-scoped invokes; may throw to reject (sheets sender gate) */
  onInvoke?(event: IpcMainInvokeEvent): void
  /** validate/transform settings before persisting (sheets zod schema) */
  parseSettings?(input: unknown): AiSettings
  /** one-time heal of a legacy stored genspark+tools-off file (default true) */
  healLegacyGskTools?: boolean
}

function readJson<T>(path: string, fallback: T): T {
  try {
    if (existsSync(path)) return JSON.parse(readFileSync(path, 'utf-8')) as T
  } catch {
    /* corrupted state file: fall back to defaults */
  }
  return fallback
}

function writeJson(path: string, value: unknown): void {
  mkdirSync(dirname(path), { recursive: true })
  writeFileSync(path, JSON.stringify(value, null, 2))
}

/** Same ceiling as the apps' local image insert limits */
const MAX_REMOTE_IMAGE_BYTES = 20 * 1024 * 1024

/**
 * Resolve the request's provider config, injecting the gsk key for the
 * genspark provider (its key never lands in the settings file — it is read
 * from the gsk login state per request).
 */
function resolveRequestConfig(
  request: Pick<AiStreamRequest, 'settings'>,
  errors: AiIpcErrorTexts,
):
  | { config: null; error: string; provider: string }
  | { config: { apiKey: string; model: string }; error?: undefined; provider: string } {
  const provider = request.settings.provider
  let config = request.settings.providers?.[provider]
  if (provider === 'genspark' && config && !config.apiKey) {
    config = { ...config, apiKey: gskApiKey() }
  }
  if (!config?.apiKey) {
    return {
      config: null,
      error: provider === 'genspark' ? errors.gskNotLoggedIn() : errors.noApiKey(provider),
      provider,
    }
  }
  if (!config.model) return { config: null, error: errors.noModel(), provider }
  return { config, provider }
}

export function registerSharedAiIpc(options: SharedAiIpcOptions): void {
  const { settingsPath, errors } = options
  const healLegacyGskTools = options.healLegacyGskTools !== false

  // Node fetch (undici) direct connections get reset under VPN/tun setups; retry over Chromium's stack
  setRescueFetch((url, init) => net.fetch(url, init))

  ipcMain.handle('ai:get-settings', (): AiSettings => {
    const stored = readJson<Partial<AiSettings> & LegacyAiSettings>(settingsPath(), {})
    // pre-toggle legacy file: genspark selected with cloud tools opted out. The
    // settings UI locks the tools switch on with genspark and apps read this
    // file live, so heal the stored flag once. Judged on the *stored* provider
    // — never the activeProvider fallback below, which must not leak into the
    // file and clobber a saved (half-configured) BYOK selection.
    if (healLegacyGskTools && (stored.provider ?? 'genspark') === 'genspark' && stored.gskToolsEnabled === false) {
      stored.gskToolsEnabled = true
      writeJson(settingsPath(), stored)
    }
    const settings = resolveAiSettings(stored, defaultAiSettings())
    // a stored BYOK provider is honored as-is; unusable configs error at request time
    settings.provider = activeProvider(settings)
    return settings
  })

  // Genspark account (gsk login state): the auth source for its cloud features;
  // the frontend uses it to guide sign-in when logged out
  ipcMain.handle(
    'ai:gsk-status',
    async (_event, withEmail?: boolean): Promise<GenSparkAccountStatus> => {
      if (!hasGskAuth()) return { loggedIn: false }
      if (!withEmail) return { loggedIn: true }
      const info = await gskLoginInfo()
      return info?.email ? { loggedIn: true, email: info.email } : { loggedIn: true }
    },
  )

  ipcMain.handle('ai:gsk-login', () => {
    ensureGenofficeLogin((url) => void shell.openExternal(url))
  })

  ipcMain.handle('ai:set-settings', (event, input: AiSettings) => {
    options.onInvoke?.(event)
    const settings = options.parseSettings ? options.parseSettings(input) : input
    writeJson(settingsPath(), settings)
  })

  const activeAiStreams = new Map<string, AbortController>()

  ipcMain.handle('ai:stream', async (event, request: AiStreamRequest) => {
    options.onInvoke?.(event)
    const { requestId, system, messages } = request
    const tools = request.tools ?? []
    const maxTokens = request.maxTokens ?? 8192
    const { config, error } = resolveRequestConfig(request, errors)
    const send = (chunk: AiStreamChunk) => {
      if (!event.sender.isDestroyed()) event.sender.send('ai:stream-chunk', chunk)
    }
    if (!config) {
      send({ requestId, type: 'error', error: error ?? 'AI request rejected' })
      return
    }
    const provider = request.settings.provider
    const controller = new AbortController()
    activeAiStreams.set(requestId, controller)
    // wire-activity keepalive: lets the renderer's silence watchdog tell a slow turn from a dead one
    let lastPing = 0
    const ping = () => {
      const now = Date.now()
      if (now - lastPing < 5_000) return
      lastPing = now
      send({ requestId, type: 'ping' })
    }
    try {
      let stopReason: string | undefined
      await streamForProvider(provider, config, system, messages, tools, maxTokens, {
        signal: controller.signal,
        onDelta: (text) => send({ requestId, type: 'delta', text }),
        onReasoningDelta: (text) => send({ requestId, type: 'reasoning', text }),
        onToolCall: (toolCall) => send({ requestId, type: 'tool-call', toolCall }),
        onActivity: ping,
        ...(options.streamStopReason
          ? { onStopReason: (reason: string) => (stopReason = reason) }
          : {}),
      })
      send({
        requestId,
        type: 'done',
        ...(options.streamStopReason && stopReason !== undefined ? { stopReason } : {}),
      })
    } catch (err) {
      if (controller.signal.aborted) {
        send({ requestId, type: 'done' })
      } else {
        send({
          requestId,
          type: 'error',
          error: err instanceof Error ? err.message : String(err),
          ...(err instanceof AiTimeoutError
            ? { errorCode: 'timeout' as const }
            : err instanceof AiCreditsError
              ? { errorCode: 'credits' as const }
              : isAiNetworkError(err)
                ? { errorCode: 'network' as const }
                : {}),
        })
      }
    } finally {
      activeAiStreams.delete(requestId)
    }
  })

  ipcMain.handle('ai:stream-cancel', (event, requestId: string) => {
    options.onInvoke?.(event)
    activeAiStreams.get(requestId)?.abort()
  })

  ipcMain.handle('ai:chat', async (event, request: AiChatRequest) => {
    options.onInvoke?.(event)
    const { config, error } = resolveRequestConfig(request, errors)
    if (!config) return { ok: false, error }
    try {
      return await chatForProvider(request.settings.provider, config, request.system, request.user)
    } catch (err) {
      return { ok: false, error: String(err) }
    }
  })

  // shared search tools (content + images): Serper with DuckDuckGo fallback
  // (gsk first when the cloud tools are on and the user is signed in)
  ipcMain.handle('ai:web-search', async (_event, query: string, maxResults?: number) => {
    try {
      const live = readJson<Partial<AiSettings>>(settingsPath(), {})
      return await webSearch(
        String(query),
        typeof maxResults === 'number' ? maxResults : 6,
        cloudToolsEnabled(live),
        live.serperApiKey,
      )
    } catch (err) {
      return { results: [], method: 'error', error: String(err) }
    }
  })

  ipcMain.handle('ai:image-search', async (_event, query: string, maxResults?: number) => {
    try {
      const live = readJson<Partial<AiSettings>>(settingsPath(), {})
      return await imageSearch(
        String(query),
        typeof maxResults === 'number' ? maxResults : 8,
        cloudToolsEnabled(live),
        live.serperApiKey,
      )
    } catch (err) {
      return { images: [], method: 'error', error: String(err) }
    }
  })

  // download image from URL → base64+mime (download in the main process avoids CORS)
  ipcMain.handle(
    'ai:fetch-image',
    async (_event, url: string): Promise<{ base64: string; mime: string } | null> => {
      try {
        // bytes generated by our own image API ride the in-memory genimage://
        // stash — resolve them directly (no network, no file I/O)
        const stashed = takeStashedImage(String(url))
        if (stashed) return stashed
        // the URL originates from AI tool calls (prompt-injectable via web search
        // results), so refuse non-http schemes and private/link-local targets;
        // redirects are followed manually so every hop is validated too.
        const resp = await fetchRemoteImage(String(url))
        if (!resp || !resp.ok || !resp.body) return null
        const declared = Number(resp.headers.get('content-length') ?? 0)
        if (declared > MAX_REMOTE_IMAGE_BYTES) return null
        // stream with a running cap: a missing/understated Content-Length must
        // not let a prompt-injected URL buffer unbounded bytes
        const reader = resp.body.getReader()
        const chunks: Buffer[] = []
        let received = 0
        for (;;) {
          const { done, value } = await reader.read()
          if (done) break
          received += value.byteLength
          if (received > MAX_REMOTE_IMAGE_BYTES) {
            await reader.cancel()
            return null
          }
          chunks.push(Buffer.from(value))
        }
        const buf = Buffer.concat(chunks)
        const ct = resp.headers.get('content-type') ?? ''
        const mime = ct.includes('png')
          ? 'image/png'
          : ct.includes('gif')
            ? 'image/gif'
            : 'image/jpeg'
        return { base64: buf.toString('base64'), mime }
      } catch {
        return null
      }
    },
  )
}

/**
 * BYOK-first image generation: when the settings carry a usable
 * OpenAI-compatible image endpoint, generate through it (no Genspark login
 * required) and return the genimage:// URL. Returns null when the image API
 * is not configured so callers fall back to their gsk path.
 */
export async function generateImageWithOwnApi(
  settings: Partial<AiSettings> | undefined,
  prompt: string,
): Promise<{ url: string } | null> {
  if (!hasImageApiConfig(settings?.imageGeneration)) return null
  const r = await openaiCompatibleGenerateImage(settings!.imageGeneration!, prompt)
  return { url: r.url }
}
