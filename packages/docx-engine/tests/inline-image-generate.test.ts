import { describe, expect, it } from 'vitest'

import { generateParagraphXml, inlineRunsXml, type GenerateContext } from '../src/generate'
import type { NewImage } from '../src/types'

function ctxBase(embedImage?: (image: NewImage) => string): GenerateContext {
  return {
    headingStyleIds: new Map(),
    allocateHyperlinkRel: () => 'rIdX',
    ...(embedImage ? { embedImage } : {}),
  }
}

describe('generateParagraphXml inline image', () => {
  it('synthesizes a run-inner drawing for a new inline image (no stored xml)', () => {
    let received: NewImage | null = null
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        runs: [
          { text: 'before ' },
          {
            text: '',
            image: {
              dataUrl: 'data:image/png;base64,QUJD',
              xml: '',
              widthPx: 100,
              heightPx: 50,
            },
          },
          { text: ' after' },
        ],
      },
      ctxBase((image) => {
        received = image
        return '<w:drawing>NEW</w:drawing>'
      }),
    )
    expect(received).toMatchObject({ base64: 'QUJD', mime: 'image/png', widthPx: 100, heightPx: 50 })
    // the image sits inline between the text runs (no <w:p> split)
    expect(xml).toContain('before ')
    expect(xml).toContain('<w:r><w:drawing>NEW</w:drawing></w:r>')
    expect(xml).toContain('after')
  })

  it('keeps a stored drawing xml verbatim and never calls embedImage', () => {
    let called = false
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        runs: [{ text: '', image: { dataUrl: 'data:image/png;base64,QUJD', xml: '<w:drawing>ORIG</w:drawing>' } }],
      },
      ctxBase(() => {
        called = true
        return '<w:drawing>NEW</w:drawing>'
      }),
    )
    expect(xml).toContain('<w:drawing>ORIG</w:drawing>')
    expect(called).toBe(false)
  })

  it('degrades to an empty run when there is no embedImage callback', () => {
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        runs: [{ text: '', image: { dataUrl: 'data:image/png;base64,QUJD', xml: '' } }],
      },
      ctxBase(),
    )
    expect(xml).toContain('<w:r></w:r>')
  })

  it('rejects an unsupported data mime instead of synthesizing', () => {
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        runs: [{ text: '', image: { dataUrl: 'data:application/zip;base64,QUJD', xml: '' } }],
      },
      ctxBase(() => '<w:drawing>NEW</w:drawing>'),
    )
    expect(xml).not.toContain('NEW')
  })

  it('inlineRunsXml (no context) leaves a stored drawing intact', () => {
    expect(inlineRunsXml([{ text: '', image: { dataUrl: 'd', xml: '<w:drawing>KEEP</w:drawing>' } }])).toContain(
      '<w:drawing>KEEP</w:drawing>',
    )
  })

  it('passes wrap + position offsets to embedImage so the float persists', () => {
    let received: NewImage | null = null
    generateParagraphXml(
      {
        type: 'paragraph',
        runs: [
          {
            text: '',
            image: {
              dataUrl: 'data:image/png;base64,QUJD',
              xml: '',
              widthPx: 100,
              heightPx: 50,
              wrap: 'square-left',
              offsetXEmu: 9525,
              offsetYEmu: 0,
            },
          },
        ],
      },
      ctxBase((image) => {
        received = image
        return '<w:drawing>NEW</w:drawing>'
      }),
    )
    expect(received).toMatchObject({ wrap: 'square-left', posOffsetEmu: { x: 9525, y: 0 } })
  })
})
