import { describe, expect, it } from 'vitest'
import { PAGE_MARK, TOTAL_PAGES_MARK, type HeaderFooter } from '@genoffice/docx-engine'
import { cssColorToHex, hfEditDomToValue, hfToEditHtml } from '../src/renderer/editor/hf-rich'

function dom(html: string): HTMLElement {
  const el = document.createElement('div')
  el.innerHTML = html
  return el
}

describe('hfToEditHtml', () => {
  it('renders runs as styled spans with field tokens visible', () => {
    const value: HeaderFooter = {
      text: '',
      paras: [
        {
          align: 'center',
          runs: [
            { text: 'Draft — ', bold: true },
            { text: `Page ${PAGE_MARK} of ${TOTAL_PAGES_MARK}`, color: 'FF0000', sizeHalfPoints: 24 },
          ],
        },
      ],
    }
    const html = hfToEditHtml(value)
    expect(html).toContain('text-align:center')
    expect(html).toContain('font-weight:600')
    expect(html).toContain('font-size:12pt')
    expect(html).toContain('color:#FF0000')
    expect(html).toContain('{PAGE} of {NUMPAGES}')
    expect(html).not.toContain(PAGE_MARK)
  })

  it('skips layout-table rows', () => {
    const value: HeaderFooter = {
      text: '',
      paras: [
        { runs: [], cells: [{ paras: [[{ text: 'cell' }]] }] },
        { align: 'right', runs: [{ text: 'visible' }] },
      ],
    }
    expect(hfToEditHtml(value)).toBe('<div class="page-hf-edit-para" style="text-align:right">visible</div>')
  })

  it('escapes html-special characters', () => {
    const value: HeaderFooter = { text: '', paras: [{ runs: [{ text: 'a < b & c' }] }] }
    expect(hfToEditHtml(value)).toContain('a &lt; b &amp; c')
  })
})

describe('hfEditDomToValue', () => {
  it('parses per-run styles back to runs', () => {
    const value: HeaderFooter = { text: '', paras: [{ align: 'left', runs: [{ text: 'old' }] }] }
    const root = dom(
      '<div style="text-align:center">plain <b>bold</b> <i>italic</i> <span style="color:#FF0000">red</span> <span style="font-size:14pt">big</span></div>',
    )
    const next = hfEditDomToValue(value, root)
    const runs = next.paras![0]!.runs
    expect(next.paras![0]!.align).toBe('center')
    expect(runs[0]).toMatchObject({ text: 'plain ' })
    expect(runs[1]).toMatchObject({ text: 'bold', bold: true })
    expect(runs[2]).toMatchObject({ text: ' ' })
    expect(runs.find((r) => r.text === 'red')).toMatchObject({ color: 'FF0000' })
    expect(runs.find((r) => r.text === 'big')).toMatchObject({ sizeHalfPoints: 28 })
  })

  it('converts tokens back to field sentinels', () => {
    const next = hfEditDomToValue(null, dom('<div>Page {PAGE} of {NUMPAGES}</div>'))
    expect(next.paras![0]!.runs[0]!.text).toBe(`Page ${PAGE_MARK} of ${TOTAL_PAGES_MARK}`)
  })

  it('inherits unset styles from the template first run', () => {
    const value: HeaderFooter = {
      text: '',
      paras: [{ align: 'right', runs: [{ text: 'styled', italic: true, color: '0000FF' }] }],
    }
    const next = hfEditDomToValue(value, dom('<div>plain <b>bold</b></div>'))
    const runs = next.paras![0]!.runs
    expect(runs[0]).toMatchObject({ text: 'plain ', italic: true, color: '0000FF' })
    expect(runs[1]).toMatchObject({ text: 'bold', bold: true, italic: true, color: '0000FF' })
  })

  it('splices cells rows back and keeps other paragraph format', () => {
    const value: HeaderFooter = {
      text: '',
      paras: [
        { runs: [], cells: [{ paras: [[{ text: 'logo' }]] }] },
        { align: 'center', runs: [{ text: 'title' }] },
      ],
    }
    const next = hfEditDomToValue(value, dom('<div>new title</div>'))
    expect(next.paras![0]!.cells).toBeTruthy()
    expect(next.paras![1]!.align).toBe('center')
    expect(next.paras![1]!.runs[0]!.text).toBe('new title')
    expect(next.text).toBe('new title')
  })

  it('splits soft line breaks into separate paragraphs', () => {
    const value: HeaderFooter = { text: '', paras: [{ align: 'left', runs: [{ text: 'a' }] }] }
    const next = hfEditDomToValue(value, dom('<div>one<br>two</div>'))
    expect(next.paras).toHaveLength(2)
    expect(next.paras![0]!.runs[0]!.text).toBe('one')
    expect(next.paras![1]!.runs[0]!.text).toBe('two')
  })

  it('empty surface clears the part', () => {
    const value: HeaderFooter = { text: 'x', paras: [{ runs: [{ text: 'x' }] }] }
    const next = hfEditDomToValue(value, dom(''))
    expect(next.text).toBe('')
    expect(next.paras).toEqual([])
  })
})

describe('cssColorToHex', () => {
  it('converts hex, rgb and named colors', () => {
    expect(cssColorToHex('#ff0000')).toBe('FF0000')
    expect(cssColorToHex('#f00')).toBe('FF0000')
    expect(cssColorToHex('rgb(255, 0, 0)')).toBe('FF0000')
    expect(cssColorToHex('red')).toBe('FF0000')
    expect(cssColorToHex('weird')).toBeUndefined()
  })
})
