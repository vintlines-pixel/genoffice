import { describe, expect, it } from 'vitest'

import { inlineStyleEditorCommand } from '../src/renderer/editor-inline-style'

describe('inlineStyleEditorCommand', () => {
  it('maps the rich-run-capable style commands to Univer’s range wrappers', () => {
    expect(inlineStyleEditorCommand('bold', '')).toEqual({
      id: 'sheet.command.set-range-bold',
      params: {},
    })
    expect(inlineStyleEditorCommand('italic', '')).toEqual({
      id: 'sheet.command.set-range-italic',
      params: {},
    })
    expect(inlineStyleEditorCommand('underline', '')).toEqual({
      id: 'sheet.command.set-range-underline',
      params: {},
    })
    expect(inlineStyleEditorCommand('strike', '')).toEqual({
      id: 'sheet.command.set-range-stroke',
      params: {},
    })
  })

  it('carries values for size, family and color', () => {
    expect(inlineStyleEditorCommand('font-size', '14')).toEqual({
      id: 'sheet.command.set-range-fontsize',
      params: { value: 14 },
    })
    expect(inlineStyleEditorCommand('font-family', 'Comic Sans MS')).toEqual({
      id: 'sheet.command.set-range-font-family',
      params: { value: 'Comic Sans MS' },
    })
    expect(inlineStyleEditorCommand('font-color', '#ff0000')).toEqual({
      id: 'sheet.command.set-range-text-color',
      params: { value: '#ff0000' },
    })
    // 'auto' is Excel's automatic (black) color inside the editor.
    expect(inlineStyleEditorCommand('font-color', 'auto')).toEqual({
      id: 'sheet.command.set-range-text-color',
      params: { value: '#000000' },
    })
  })

  it('keeps whole-cell formats out of the editor path', () => {
    // Double underline has no rich-run equivalent.
    expect(inlineStyleEditorCommand('underline', 'double')).toBeNull()
    expect(inlineStyleEditorCommand('fill', '#ffff00')).toBeNull()
    expect(inlineStyleEditorCommand('border', 'all')).toBeNull()
    expect(inlineStyleEditorCommand('wrap', 'on')).toBeNull()
  })

  it('rejects malformed values instead of dispatching', () => {
    expect(inlineStyleEditorCommand('font-size', 'abc')).toBeNull()
    expect(inlineStyleEditorCommand('font-size', '0')).toBeNull()
    expect(inlineStyleEditorCommand('font-size', '')).toBeNull()
    expect(inlineStyleEditorCommand('font-family', '')).toBeNull()
  })
})
