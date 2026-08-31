import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { SectionSettings } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeTool, type AiSectionAccess } from '../src/renderer/ai/tools'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const para = (t: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [{ type: 'text', text: t }],
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function createEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [para('Body paragraph.')] },
  })
  editors.add(editor)
  return editor
}

const NUM_IDS = { bullet: null, ordered: null }
const CM = (cm: number) => Math.round((cm * 1440) / 2.54)

// A4 portrait with 1-inch margins
function a4(): SectionSettings {
  return {
    pageWidth: CM(21.0),
    pageHeight: CM(29.7),
    orientation: 'portrait',
    marginTop: CM(2.54),
    marginRight: CM(2.54),
    marginBottom: CM(2.54),
    marginLeft: CM(2.54),
    pageBorder: false,
    columns: 1,
  }
}

function run(
  access: AiSectionAccess,
  input: Record<string, unknown>,
): { mutated: boolean; isError: boolean; output: string } {
  const result = executeTool(
    createEditor(),
    { id: 't', name: 'set_page_setup', input },
    NUM_IDS,
    undefined,
    undefined,
    null,
    undefined,
    undefined,
    access,
  ) as { mutated: boolean; isError: boolean; output: string }
  return result
}

describe('set_page_setup tool', () => {
  it('applies the "narrow" preset to all four margins', () => {
    const current = a4()
    const writes: SectionSettings[] = []
    const r = run(
      { read: () => current, set: (next) => (writes.push(next), null) },
      { margins: 'narrow' },
    )
    expect(r.isError).toBeFalsy()
    expect(writes).toHaveLength(1)
    const next = writes[0]!
    expect(next.marginTop).toBe(CM(1.27))
    expect(next.marginRight).toBe(CM(1.27))
    expect(next.marginBottom).toBe(CM(1.27))
    expect(next.marginLeft).toBe(CM(1.27))
    // page size + orientation untouched
    expect(next.pageWidth).toBe(current.pageWidth)
    expect(next.orientation).toBe('portrait')
  })

  it('sets a single margin without touching the others', () => {
    const current = a4()
    const writes: SectionSettings[] = []
    const r = run(
      { read: () => current, set: (next) => (writes.push(next), null) },
      { marginTop: 3 },
    )
    expect(r.isError).toBeFalsy()
    const next = writes[0]!
    expect(next.marginTop).toBe(CM(3))
    expect(next.marginRight).toBe(current.marginRight)
    expect(next.marginBottom).toBe(current.marginBottom)
    expect(next.marginLeft).toBe(current.marginLeft)
  })

  it('swaps width/height for landscape orientation', () => {
    const current = a4()
    const writes: SectionSettings[] = []
    const r = run(
      { read: () => current, set: (next) => (writes.push(next), null) },
      { orientation: 'landscape' },
    )
    expect(r.isError).toBeFalsy()
    const next = writes[0]!
    expect(next.orientation).toBe('landscape')
    expect(next.pageWidth).toBe(current.pageHeight)
    expect(next.pageHeight).toBe(current.pageWidth)
  })

  it('rejects an out-of-range margin', () => {
    const r = run({ read: () => a4(), set: () => null }, { marginLeft: 99 })
    expect(r.isError).toBe(true)
    expect(r.output).toContain('marginLeft must be between')
  })

  it('reports unavailable when no section accessor is provided', () => {
    const result = executeTool(
      createEditor(),
      { id: 't', name: 'set_page_setup', input: { margins: 'narrow' } },
      NUM_IDS,
      undefined,
      undefined,
      null,
    ) as { isError: boolean; output: string }
    expect(result.isError).toBe(true)
    expect(result.output).toContain('not available')
  })
})
