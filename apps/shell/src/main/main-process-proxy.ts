/**
 * Main-process network bootstrap: on mainland-China networks Node's fetch
 * (undici) bypasses the system proxy, so direct calls to overseas LLM /
 * image-search APIs time out or get region-blocked (403). Prefer proxy env
 * vars (terminal launch); a packaged app launched from Finder inherits no
 * shell env vars, so fall back to the system HTTP proxy resolved by
 * Chromium. The renderer uses Chromium's system proxy and is unaffected.
 * Same bootstrap as slides-main startSlidesStandalone.
 */
import { session } from 'electron'
import { setGskProxyUrl } from '@genoffice/ai-search'

/** awaited by login IPC so the first status probe / login click cannot race the proxy resolution */
let proxyBootstrap: Promise<void> = Promise.resolve()

export function getProxyBootstrap(): Promise<void> {
  return proxyBootstrap
}

export async function installMainProcessProxy(): Promise<void> {
  let proxyUrl = [
    process.env.HTTPS_PROXY,
    process.env.https_proxy,
    process.env.HTTP_PROXY,
    process.env.http_proxy,
    process.env.ALL_PROXY,
    process.env.all_proxy,
  ].find((v) => v && /^https?:\/\//.test(v))
  if (!proxyUrl) {
    try {
      // PAC/rule proxies answer per-host: probe the host the login flow, the
      // Genspark LLM proxy and the gsk CLI actually target
      const resolved = await session.defaultSession.resolveProxy('https://www.genspark.ai/')
      const m = /PROXY\s+([^;\s]+)/.exec(resolved)
      if (m) proxyUrl = `http://${m[1]}`
    } catch {
      /* no system proxy */
    }
  }
  if (!proxyUrl) return
  // spawned gsk CLI children (login/search/…) do their own fetch and never see
  // the dispatcher below — forward the proxy to them via env
  setGskProxyUrl(proxyUrl)
  try {
    const { ProxyAgent, setGlobalDispatcher } = await import('undici')
    setGlobalDispatcher(new ProxyAgent(proxyUrl))
    // strip user:pass credentials before logging
    console.log('[proxy] main-process fetch via', proxyUrl.replace(/\/\/[^@/]*@/, '//***@'))
  } catch (e) {
    console.warn('[proxy] failed to set ProxyAgent:', e)
  }
}

/** Called once at whenReady: starts the bootstrap and records its promise. */
export function startMainProcessProxy(): void {
  proxyBootstrap = installMainProcessProxy()
}
