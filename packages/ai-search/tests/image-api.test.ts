import { describe, it, expect, vi, afterEach } from 'vitest'
import {
  openaiCompatibleGenerateImage,
  stashGeneratedImage,
  takeStashedImage,
} from '../src/image-api'

const realFetch = globalThis.fetch
afterEach(() => {
  globalThis.fetch = realFetch
})

const cfg = { baseUrl: 'https://img.example/v1', apiKey: 'sk-img', model: 'gpt-image-1' }

describe('openaiCompatibleGenerateImage', () => {
  it('returns stashed bytes from a b64_json response', async () => {
    globalThis.fetch = vi.fn(async () => ({ ok: true, status: 200, text: async () =>
      JSON.stringify({ data: [{ b64_json: 'QUJD' }] }),
    }) as unknown as Response)
    const r = await openaiCompatibleGenerateImage(cfg, 'a cat')
    expect(r.mime).toBe('image/png')
    expect(r.url).toMatch(/^genimage:\/\//)
    expect(takeStashedImage(r.url)).toEqual({ base64: 'QUJD', mime: 'image/png' })
  })

  it('follows a url response and caches the downloaded bytes', async () => {
    const calls: string[] = []
    globalThis.fetch = vi.fn(async (url: any) => {
      calls.push(String(url))
      if (String(url).endsWith('/images/generations')) {
        return { ok: true, status: 200, text: async () =>
          JSON.stringify({ data: [{ url: 'https://cdn.example/img.png' }] }),
        } as unknown as Response
      }
      return {
        ok: true,
        status: 200,
        headers: new Map([['content-type', 'image/png']]),
        arrayBuffer: async () => new ArrayBuffer(4),
      } as unknown as Response
    })
    const r = await openaiCompatibleGenerateImage(cfg, 'a cat')
    expect(calls[0]).toBe('https://img.example/v1/images/generations')
    expect(calls[1]).toBe('https://cdn.example/img.png')
    expect(r.base64).toBeTruthy()
    expect(takeStashedImage(r.url)?.base64).toBe(r.base64)
  })

  it('surfaces API errors', async () => {
    globalThis.fetch = vi.fn(async () => ({ ok: false, status: 402, text: async () => 'no credits' }) as unknown as Response)
    await expect(openaiCompatibleGenerateImage(cfg, 'x')).rejects.toThrow(/402/)
  })
})

describe('genimage stash', () => {
  it('unknown ids and foreign schemes resolve to null', () => {
    expect(takeStashedImage('genimage://nope')).toBeNull()
    expect(takeStashedImage('https://evil.example/genimage://x')).toBeNull()
    const url = stashGeneratedImage('QQ==', 'image/png')
    expect(takeStashedImage(url)).toEqual({ base64: 'QQ==', mime: 'image/png' })
  })
})
