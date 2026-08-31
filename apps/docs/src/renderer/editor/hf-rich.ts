/**
 * Rich-text view of a header/footer part for the in-place editing surface.
 *
 * Where hf-text.ts edits the part as plain lines (each line inherits the first
 * run's style), this maps the part to/from styled HTML so the on-canvas editor
 * can carry per-run formatting (bold / italic / underline / color / size /
 * font) and per-paragraph alignment — Word-style, without flattening.
 *
 * Layout-table rows (cells) stay out of the editing surface and are spliced
 * back at their original positions on commit, exactly like applyHfText.
 */
import {
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
  type HfParagraph,
  type Run,
} from '@genoffice/docx-engine'
import { hfParasOf } from './hf-text'

export const PAGE_TOKEN = '{PAGE}'
export const TOTAL_TOKEN = '{NUMPAGES}'

/** font sizes offered by the header/footer formatting toolbar (points) */
export const HF_FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 36]

const escapeHtml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/** substitute the private-use field sentinels with their editable tokens */
const displayText = (text: string): string =>
  text.replaceAll(PAGE_MARK, PAGE_TOKEN).replaceAll(TOTAL_PAGES_MARK, TOTAL_TOKEN)

function runStyleCss(run: Run): string {
  const parts: string[] = []
  if (run.bold) parts.push('font-weight:600')
  if (run.italic) parts.push('font-style:italic')
  if (run.underline) parts.push('text-decoration:underline')
  if (run.strike) parts.push('text-decoration:line-through')
  if (run.color) parts.push(`color:#${run.color}`)
  if (run.sizeHalfPoints) parts.push(`font-size:${run.sizeHalfPoints / 2}pt`)
  if (run.font) parts.push(`font-family:${JSON.stringify(run.font)}`)
  return parts.join(';')
}

function runToHtml(run: Run): string {
  const text = displayText(run.text)
  const style = runStyleCss(run)
  return style ? `<span style="${escapeHtml(style)}">${escapeHtml(text)}</span>` : escapeHtml(text)
}

/** HTML for one paragraph's runs (no wrapping block) */
const runsToHtml = (runs: Run[]): string => runs.map(runToHtml).join('')

/** align of a header paragraph as a CSS text-align value */
const alignCss = (align: HfParagraph['align']): string | undefined =>
  align === 'left' || align === 'center' || align === 'right' ? align : undefined

/**
 * Editable HTML of the part: one <div> per text paragraph (cells rows omitted),
 * each run a styled <span>. Field sentinels become visible {PAGE}/{NUMPAGES}.
 */
export function hfToEditHtml(value: HeaderFooter): string {
  const blocks = hfParasOf(value)
    .filter((p) => !p.cells)
    .map((p) => {
      const align = alignCss(p.align)
      return `<div class="page-hf-edit-para"${
        align ? ` style="text-align:${align}"` : ''
      }>${runsToHtml(p.runs)}</div>`
    })
    .join('')
  return blocks
}

/** one flattened inline segment collected while walking the edited DOM */
interface Segment {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
  sizeHalfPoints?: number
  font?: string
}

interface WalkStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
  sizeHalfPoints?: number
  font?: string
}

const NAMED_COLORS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  gray: '808080',
  grey: '808080',
}

/** inline style color -> hex without '#', uppercased (undefined when unknown) */
export function cssColorToHex(color: string): string | undefined {
  const c = color.trim()
  if (!c) return undefined
  if (c[0] === '#') {
    const hex = c.slice(1)
    if (hex.length === 3) return (hex[0]! + hex[0]! + hex[1]! + hex[1]! + hex[2]! + hex[2]!).toUpperCase()
    if (hex.length === 6) return hex.toUpperCase()
    return undefined
  }
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(c)
  if (rgb) {
    const to = (n: string) => Number(n).toString(16).padStart(2, '0')
    return (to(rgb[1]!) + to(rgb[2]!) + to(rgb[3]!)).toUpperCase()
  }
  return NAMED_COLORS[c.toLowerCase()]
}

const isBoldTag = (t: string) => t === 'B' || t === 'STRONG'
const isItalicTag = (t: string) => t === 'I' || t === 'EM'
const isUnderlineTag = (t: string) => t === 'U'

function mergeWalkStyle(el: HTMLElement, prev: WalkStyle): WalkStyle {
  const s: WalkStyle = { ...prev }
  const st = el.style
  const weight = st.fontWeight
  if (isBoldTag(el.tagName) || weight === '600' || weight === '700' || weight === 'bold' || weight === 'bolder') {
    s.bold = true
  } else if (weight === 'normal' || weight === '400') {
    s.bold = false
  }
  const fs = st.fontStyle
  if (isItalicTag(el.tagName) || fs === 'italic' || fs === 'oblique') s.italic = true
  else if (fs === 'normal') s.italic = false
  const deco = st.textDecorationLine ?? st.textDecoration ?? ''
  if (isUnderlineTag(el.tagName) || deco.includes('underline')) s.underline = true
  if (st.color) {
    const hex = cssColorToHex(st.color)
    if (hex) s.color = hex
  }
  if (st.fontSize) {
    const pt = parseFloat(st.fontSize)
    if (Number.isFinite(pt) && pt > 0) s.sizeHalfPoints = Math.round(pt * 2)
  }
  if (st.fontFamily) {
    const family = st.fontFamily.split(',')[0]!.replace(/^['"]|['"]$/g, '').trim()
    if (family) s.font = family
  }
  return s
}

/** walk a block element into flat inline segments (style-resolved, tokens decoded) */
function collectSegments(block: Element): Segment[] {
  const segs: Segment[] = []
  const walk = (node: Node, style: WalkStyle): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = (node.textContent ?? '')
        .replaceAll(PAGE_TOKEN, PAGE_MARK)
        .replaceAll(TOTAL_TOKEN, TOTAL_PAGES_MARK)
      if (text) segs.push({ text, ...style })
      return
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return
    const el = node as HTMLElement
    if (el.tagName === 'BR') {
      segs.push({ text: '\n', ...style })
      return
    }
    const next = mergeWalkStyle(el, style)
    for (const child of Array.from(el.childNodes)) walk(child, next)
  }
  walk(block, {})
  return segs
}

/** collapse adjacent same-style segments into one */
function mergeSegments(segs: Segment[]): Segment[] {
  const out: Segment[] = []
  for (const seg of segs) {
    const last = out[out.length - 1]
    if (
      last &&
      last.bold === seg.bold &&
      last.italic === seg.italic &&
      last.underline === seg.underline &&
      last.color === seg.color &&
      last.sizeHalfPoints === seg.sizeHalfPoints &&
      last.font === seg.font
    ) {
      last.text += seg.text
    } else {
      out.push({ ...seg })
    }
  }
  return out
}

/** build a Run from a segment, inheriting unset styles from the template run */
function segmentToRun(seg: Segment, base: Run): Run {
  const run: Run = { text: seg.text }
  const inherit = <K extends 'bold' | 'italic' | 'underline'>(
    key: K,
    value: boolean | undefined,
  ): void => {
    if (value !== undefined) run[key] = value
    else if (base[key] !== undefined) run[key] = base[key]
  }
  inherit('bold', seg.bold)
  inherit('italic', seg.italic)
  inherit('underline', seg.underline)
  if (seg.color) run.color = seg.color
  else if (base.color) run.color = base.color
  if (seg.sizeHalfPoints) run.sizeHalfPoints = seg.sizeHalfPoints
  else if (base.sizeHalfPoints) run.sizeHalfPoints = base.sizeHalfPoints
  if (seg.font) run.font = seg.font
  else if (base.font || base.fontAscii) {
    run.font = base.font
    run.fontAscii = base.fontAscii
  }
  return run
}

function blockAlignOf(el: Element): HfParagraph['align'] | undefined {
  const a = (el as HTMLElement).style.textAlign
  if (a === 'left' || a === 'center' || a === 'right') return a
  return undefined
}

/** HTML text-align -> HfParagraph align (inherit template when absent) */
const alignOf = (el: Element, template: HfParagraph): HfParagraph['align'] =>
  blockAlignOf(el) ?? template.align

/** top-level block nodes of the edit surface (divs/browser line boxes) */
function editBlocks(root: HTMLElement): Element[] {
  const out: Element[] = []
  let textOnly: string | null = null
  for (const child of Array.from(root.childNodes)) {
    if (child.nodeType === Node.ELEMENT_NODE) {
      const el = child as HTMLElement
      if (el.tagName === 'DIV' || el.tagName === 'P') out.push(el)
      else out.push(el) // unexpected wrapper (e.g. span): still a block-ish unit
    } else if (child.nodeType === Node.TEXT_NODE && (child.textContent ?? '').trim()) {
      // contentEditable can leave bare text when the user deletes every wrapper
      textOnly = child.textContent ?? ''
    }
  }
  if (out.length === 0 && textOnly !== null) {
    const div = document.createElement('div')
    div.textContent = textOnly
    out.push(div)
  }
  return out
}

/**
 * Parse the edited surface DOM back into the part: each block becomes one
 * paragraph (align from the block, per-run styles from the inline HTML), cells
 * rows spliced back at their original positions. Unset run styles inherit the
 * template paragraph's first run (same as the plain-text path).
 */
export function hfEditDomToValue(value: HeaderFooter | null, root: HTMLElement): HeaderFooter {
  const base = value ?? { text: '' }
  const paras = hfParasOf(base)
  const textParas = paras.filter((p) => !p.cells)
  const templates: HfParagraph[] = textParas.length > 0 ? textParas : [{ align: 'center', runs: [] }]
  const blocks = editBlocks(root)

  // one edited paragraph per block; <br> inside a block becomes a line split
  const edited: HfParagraph[] = []
  blocks.forEach((block, bi) => {
    const template = templates[Math.min(bi, templates.length - 1)]!
    const segs = mergeSegments(collectSegments(block))
    const baseRun = template.runs[0] ?? {}
    // split on soft line breaks into separate paragraphs
    const lines: Segment[][] = [[]]
    for (const seg of segs) {
      const pieces = seg.text.split('\n')
      pieces.forEach((piece, k) => {
        if (k > 0) lines.push([])
        if (piece) lines[lines.length - 1]!.push({ ...seg, text: piece })
      })
    }
    for (const line of lines) {
      edited.push({
        ...template,
        align: alignOf(block, template),
        runs: line.map((s) => segmentToRun(s, baseRun)),
      })
    }
  })

  const nextParas: HfParagraph[] = []
  let ei = 0
  for (const p of paras) {
    if (p.cells) nextParas.push(p)
    else if (ei < edited.length) nextParas.push(edited[ei++]!)
  }
  nextParas.push(...edited.slice(ei))

  const nextText = edited.map((p) => p.runs.map((r) => r.text).join('')).join('')
  return { ...base, text: nextText, paras: nextParas }
}
