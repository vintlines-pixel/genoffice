import { describe, expect, it } from 'vitest'
import {
  AI_PROVIDERS,
  activeProvider,
  cloudToolsEnabled,
  defaultAiSettings,
  resolveAiSettings,
} from '../src/providers'
import type { AiProviderId } from '../src/types'

describe('defaultAiSettings', () => {
  it('gives every provider its default model and an empty key by default', () => {
    const settings = defaultAiSettings()
    // BYOK default: the custom (OpenAI-compatible endpoint) slot is selected,
    // Genspark login is not required, and gsk cloud tools start off
    expect(settings.provider).toBe('custom')
    expect(settings.gskToolsEnabled).toBe(false)
    for (const meta of AI_PROVIDERS) {
      expect(settings.providers[meta.id].apiKey).toBe('')
      expect(settings.providers[meta.id].model).toBe(meta.defaultModel)
    }
    expect(settings.providers.custom.baseUrl).toBe('')
    expect(settings.providers.anthropic.baseUrl).toBeUndefined()
  })

  it('applies caller-supplied default keys only to the listed providers', () => {
    const settings = defaultAiSettings({ anthropic: 'sk-ant-preset' })
    expect(settings.providers.anthropic.apiKey).toBe('sk-ant-preset')
    expect(settings.providers.gemini.apiKey).toBe('')
  })
})

describe('provider model catalog', () => {
  it('offers DeepSeek Vision Exp only through the direct BYOK provider', () => {
    const genspark = AI_PROVIDERS.find((provider) => provider.id === 'genspark')!
    const deepseek = AI_PROVIDERS.find((provider) => provider.id === 'deepseek')!

    expect(deepseek.models).toContain('deepseek-v4-flash-vision-exp')
    expect(genspark.models).not.toContain('deep-seek-v4-flash')
    expect(genspark.models).not.toContain('deep-seek-v4-flash-vision-exp-openrouter')
  })
})

describe('resolveAiSettings', () => {
  it('returns fresh defaults when nothing is stored', () => {
    const defaults = defaultAiSettings({ anthropic: 'sk-ant-preset' })
    expect(resolveAiSettings({}, defaults)).toEqual(defaults)
  })

  it('migrates the pre-provider single-endpoint shape into the custom provider', () => {
    const defaults = defaultAiSettings()
    const resolved = resolveAiSettings(
      { apiKey: 'legacy-key', model: 'legacy-model', baseUrl: 'https://legacy.example.com/v1' },
      defaults,
    )
    expect(resolved.providers.custom).toEqual({
      apiKey: 'legacy-key',
      model: 'legacy-model',
      baseUrl: 'https://legacy.example.com/v1',
    })
    // untouched providers keep their defaults
    expect(resolved.providers.anthropic).toEqual(defaults.providers.anthropic)
  })

  it('defaults the legacy base URL to the OpenAI endpoint when omitted', () => {
    const resolved = resolveAiSettings({ apiKey: 'legacy-key' }, defaultAiSettings())
    expect(resolved.providers.custom.baseUrl).toBe('https://api.openai.com/v1')
  })

  it('passes through the optional serperApiKey (trimmed)', () => {
    const resolved = resolveAiSettings(
      { providers: {} as never, serperApiKey: ' sk-test ' },
      defaultAiSettings(),
    )
    expect(resolved.serperApiKey).toBe('sk-test')
    // absent stays absent
    expect(resolveAiSettings({ providers: {} as never }, defaultAiSettings()).serperApiKey).toBeUndefined()
  })

  it('keeps imageGeneration only when fully configured, trimmed', () => {
    const full = resolveAiSettings(
      {
        providers: {} as never,
        imageGeneration: { baseUrl: ' https://img.example/v1 ', apiKey: ' sk ', model: ' gpt-image-1 ' },
      },
      defaultAiSettings(),
    )
    expect(full.imageGeneration).toEqual({
      baseUrl: 'https://img.example/v1',
      apiKey: 'sk',
      model: 'gpt-image-1',
    })
    // half-filled = not configured
    const half = resolveAiSettings(
      { providers: {} as never, imageGeneration: { baseUrl: 'https://x/v1', apiKey: '', model: '' } },
      defaultAiSettings(),
    )
    expect(half.imageGeneration).toBeUndefined()
  })

  it('merges stored multi-provider settings over the defaults, provider by provider', () => {
    const defaults = defaultAiSettings({ anthropic: 'preset-key' })
    const resolved = resolveAiSettings(
      {
        provider: 'gemini',
        providers: {
          gemini: { apiKey: 'stored-gemini-key', model: 'gemini-2.5-pro' },
        } as never,
      },
      defaults,
    )
    expect(resolved.provider).toBe('gemini')
    expect(resolved.providers.gemini).toEqual({
      apiKey: 'stored-gemini-key',
      model: 'gemini-2.5-pro',
    })
    // provider not mentioned in stored.providers keeps the computed default
    expect(resolved.providers.anthropic.apiKey).toBe('preset-key')
  })

  it('rewrites a stored model id the vendor has retired', () => {
    const resolved = resolveAiSettings(
      {
        providers: {
          deepseek: { apiKey: 'sk-user', model: 'deepseek-reasoner' },
        } as never,
      },
      defaultAiSettings(),
    )
    expect(resolved.providers.deepseek).toEqual({ apiKey: 'sk-user', model: 'deepseek-v4-flash' })
  })

  it('leaves a still-supported model id alone', () => {
    const resolved = resolveAiSettings(
      {
        providers: {
          deepseek: { apiKey: 'sk-user', model: 'deepseek-v4-pro' },
        } as never,
      },
      defaultAiSettings(),
    )
    expect(resolved.providers.deepseek.model).toBe('deepseek-v4-pro')
  })

  it('trims whitespace pasted around stored keys and base URLs', () => {
    const resolved = resolveAiSettings(
      {
        providers: {
          deepseek: { apiKey: ' sk-user\n', model: 'deepseek-v4-pro' },
          custom: { apiKey: 'k', model: 'm', baseUrl: ' http://localhost:1234/v1 ' },
        } as never,
      },
      defaultAiSettings(),
    )
    expect(resolved.providers.deepseek.apiKey).toBe('sk-user')
    expect(resolved.providers.deepseek.baseUrl).toBeUndefined()
    expect(resolved.providers.custom.baseUrl).toBe('http://localhost:1234/v1')
  })

  it('trims the legacy single-endpoint key and base URL too', () => {
    const resolved = resolveAiSettings(
      { apiKey: ' legacy-key ', baseUrl: ' https://legacy.example.com/v1 ' },
      defaultAiSettings(),
    )
    expect(resolved.providers.custom.apiKey).toBe('legacy-key')
    expect(resolved.providers.custom.baseUrl).toBe('https://legacy.example.com/v1')
  })
})

describe('activeProvider', () => {
  it('honors the stored provider selection; unusable configs error at request time', () => {
    const settings = defaultAiSettings()
    expect(settings.provider).toBe('custom')
    // unconfigured custom is honored — the ai:stream handler replies with a
    // "configure your API key" error instead of silently rerouting
    expect(activeProvider(settings)).toBe('custom')

    settings.provider = 'kimi'
    expect(activeProvider(settings)).toBe('kimi') // no key yet — same semantics
    settings.providers.kimi.apiKey = 'sk-user'
    expect(activeProvider(settings)).toBe('kimi')
  })

  it('genspark stays genspark (its key is injected from the gsk login at request time)', () => {
    const settings = defaultAiSettings()
    settings.provider = 'genspark'
    expect(activeProvider(settings)).toBe('genspark')
  })

  it('falls back to the default provider for unknown ids from a hand-edited settings file', () => {
    const settings = defaultAiSettings()
    settings.provider = 'nonsense' as AiProviderId
    expect(activeProvider(settings)).toBe('custom')
  })
})

describe('gskToolsEnabled', () => {
  it('defaults off for fresh installs', () => {
    expect(cloudToolsEnabled(defaultAiSettings())).toBe(false)
  })

  it('keeps a pre-toggle settings file (multi-provider shape, flag absent) on', () => {
    const legacy = resolveAiSettings({ providers: {} as never }, defaultAiSettings())
    expect(cloudToolsEnabled(legacy)).toBe(true)
  })

  it('only an explicit false turns it off', () => {
    const off = resolveAiSettings(
      { providers: {} as never, gskToolsEnabled: false },
      defaultAiSettings(),
    )
    expect(off.gskToolsEnabled).toBe(false)
    expect(cloudToolsEnabled(off)).toBe(false)
    const on = resolveAiSettings(
      { providers: {} as never, gskToolsEnabled: true },
      defaultAiSettings(),
    )
    expect(cloudToolsEnabled(on)).toBe(true)
  })
})
