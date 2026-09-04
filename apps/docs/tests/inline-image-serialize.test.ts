import { describe, expect, it } from 'vitest'

import { inlineToRuns } from '../src/renderer/editor/convert'

describe('inlineToRuns docInlineImage', () => {
  it('maps a new inline image (no stored xml) to a run, so the engine can synthesize', () => {
    const runs = inlineToRuns([
      { type: 'text', text: 'before', marks: [] },
      {
        type: 'docInlineImage',
        attrs: {
          dataUrl: 'data:image/png;base64,QUJD',
          widthPx: 120,
          heightPx: 60,
          xml: '',
        },
      },
      { type: 'text', text: 'after', marks: [] },
    ])
    expect(runs[0]).toMatchObject({ text: 'before' })
    expect(runs[1]).toMatchObject({
      text: '',
      image: { dataUrl: 'data:image/png;base64,QUJD', xml: '', widthPx: 120, heightPx: 60 },
    })
    expect(runs[2]).toMatchObject({ text: 'after' })
  })

  it('carries floating/wrap state through to the run image', () => {
    const runs = inlineToRuns([
      {
        type: 'docInlineImage',
        attrs: {
          dataUrl: 'data:image/png;base64,QUJD',
          widthPx: 120,
          heightPx: 60,
          xml: '',
          wrap: 'square-right',
          offsetXEmu: 9525,
          offsetYEmu: 0,
          border: { color: 'FF0000', widthPt: 1 },
        },
      },
    ])
    expect(runs[0]).toMatchObject({
      text: '',
      image: {
        dataUrl: 'data:image/png;base64,QUJD',
        wrap: 'square-right',
        offsetXEmu: 9525,
        offsetYEmu: 0,
        border: { color: 'FF0000', widthPt: 1 },
      },
    })
  })
})
