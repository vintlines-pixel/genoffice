import { describe, expect, it } from 'vitest'

import { fitImageDim } from '@genoffice/ui'

describe('fitImageDim', () => {
  it('keeps small images unchanged (never upscales)', () => {
    expect(fitImageDim(600, 400, 1280)).toEqual({ width: 600, height: 400 })
    expect(fitImageDim(1280, 1280, 1280)).toEqual({ width: 1280, height: 1280 })
  })

  it('scales the long edge to maxDim, keeping the aspect ratio', () => {
    expect(fitImageDim(4000, 3000, 1280)).toEqual({ width: 1280, height: 960 })
    expect(fitImageDim(3000, 4000, 1280)).toEqual({ width: 960, height: 1280 })
  })

  it('handles extreme aspect ratios and degenerate input', () => {
    expect(fitImageDim(20000, 100, 1280)).toEqual({ width: 1280, height: 6 })
    expect(fitImageDim(0, 0, 1280)).toEqual({ width: 0, height: 0 })
  })
})
