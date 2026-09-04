import { describe, expect, it } from 'vitest'

import { buildBlankDocx, parseDocx, type Block } from '@genoffice/docx-engine'

/** plain text of a paragraph-like block; empty for protected/passthrough ones */
function blockText(block: Block): string {
  return (block.runs ?? []).map((run) => run.text).join('')
}

describe('buildBlankDocx paragraphs', () => {
  // parseDocx yields the seeded paragraphs followed by the trailing section
  // paragraph (w:sectPr holder) — compare the leading seeded ones only.
  it('seeds the body with the given paragraphs', async () => {
    const bytes = await buildBlankDocx({ paragraphs: ['Line one', 'Line two'] })
    const doc = await parseDocx(bytes)
    expect(doc.blocks.slice(0, 2).map(blockText)).toEqual(['Line one', 'Line two'])
  })

  it('keeps formatting markup out of the imported text (XML-escaped)', async () => {
    const bytes = await buildBlankDocx({ paragraphs: ['A <B> & "quote"'] })
    const doc = await parseDocx(bytes)
    expect(doc.blocks.slice(0, 1).map(blockText)).toEqual(['A <B> & "quote"'])
  })

  it('defaults to a single empty paragraph', async () => {
    const bytes = await buildBlankDocx()
    const doc = await parseDocx(bytes)
    expect(doc.blocks.slice(0, 1).map(blockText)).toEqual([''])
  })
})
