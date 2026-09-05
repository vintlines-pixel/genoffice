import { useEffect, useRef, useState } from 'react'
import {
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type HfImage,
  type HfParagraph,
  type NewImage,
  type Run,
} from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import {
  hfLeadIndentCss,
  hfSegLeftCss,
  hfTabSegments,
  hfUsesLegacyHash,
  paraBorderCss,
} from '../editor/hf-dom'
import { hfParasOf, PAGE_TOKEN } from '../editor/hf-text'
import { HF_FONT_SIZES, TOTAL_TOKEN, hfEditDomToValue, hfToEditHtml } from '../editor/hf-rich'
import { fontFamiliesFor } from '../font-list'
import { cssRunFontFamily, runLetterSpacingCss } from '../line-metrics'
import {
  IconAlignCenter,
  IconAlignLeft,
  IconAlignRight,
  IconClose,
  IconFontColorA,
  IconPageNumber,
  IconReplaceImage,
} from './icons'

export interface HfValue {
  text: string
  pageNumber?: boolean
  paras?: HfParagraph[]
}

/** strip image: a parsed part image (parsedIndex = its parse-order slot, the
 *  imageEdits target) or one newly added this session */
export type HfStripImage = HfImage & { parsedIndex?: number }

/** Word standard-colors palette for the header/footer text-color popover */
const HF_COLORS = [
  '000000',
  'FFFFFF',
  'C00000',
  'FF0000',
  'ED7D31',
  'FFC000',
  'FFFF00',
  '00B050',
  '0070C0',
  '00B0F0',
  '002060',
  '7030A0',
  '808080',
  'C0C0C0',
]

function runStyle(run: Run): React.CSSProperties {
  const style: React.CSSProperties = {}
  if (run.bold) style.fontWeight = 600
  else if (run.bold === false) style.fontWeight = 'normal'
  if (run.italic) style.fontStyle = 'italic'
  else if (run.italic === false) style.fontStyle = 'normal'
  if (run.underline) style.textDecoration = 'underline'
  if (run.strike) style.textDecoration = `${style.textDecoration ?? ''} line-through`.trim()
  if (run.color) style.color = `#${run.color}`
  if (run.sizeHalfPoints) style.fontSize = `${run.sizeHalfPoints / 2}pt`
  const letterSpacing = runLetterSpacingCss(run)
  if (letterSpacing) style.letterSpacing = letterSpacing
  if (run.font || run.fontAscii) style.fontFamily = cssRunFontFamily(run.fontAscii, run.font)
  if (run.caps === 'all') style.textTransform = 'uppercase'
  else if (run.caps === 'small') style.fontVariantCaps = 'small-caps'
  else if (run.caps === 'none') {
    style.textTransform = 'none'
    style.fontVariantCaps = 'normal'
  }
  return style
}

/** pick a png/jpeg/gif from disk and read its natural size (shared by insert and replace) */
async function pickImageFromDisk(): Promise<NewImage | null> {
  const picked = await window.desktop.pickImage()
  if (!picked) return null
  const mime = picked.mime
  if (!/^image\/(png|jpeg|gif)$/.test(mime)) return null
  const dataUrl = `data:${mime};base64,${picked.base64}`
  const size = await new Promise<{ width: number; height: number }>((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight })
    img.onerror = () => reject(new Error('Failed to read image'))
    img.src = dataUrl
  })
  return {
    base64: picked.base64,
    mime: picked.mime as NewImage['mime'],
    widthPx: size.width,
    heightPx: size.height,
  }
}

/** document content colors (w:shd / w:pBdr), theme-independent; mirrors makeGapHfEl */
function paraStyle(para: HfParagraph): React.CSSProperties {
  const style: React.CSSProperties = {}
  if (para.bidi) style.direction = 'rtl'
  if (para.align) {
    style.textAlign =
      para.align === 'left' || para.align === 'center' || para.align === 'right'
        ? para.align
        : 'justify'
  }
  // frame placement wins over the paragraph's own jc (mirrors makeGapHfEl)
  if (para.frameXAlign) style.textAlign = para.frameXAlign
  if (para.shadingFill) style.backgroundColor = `#${para.shadingFill}`
  if (para.borders) {
    const line = (side: 't' | 'b' | 'l' | 'r') => paraBorderCss(para.borderLines?.[side])
    if (para.borders.includes('t')) style.borderTop = line('t')
    if (para.borders.includes('b')) style.borderBottom = line('b')
    if (para.borders.includes('l')) style.borderLeft = line('l')
    if (para.borders.includes('r')) style.borderRight = line('r')
    style.padding = '1px 4px'
  }
  return style
}

/**
 * Header / footer zone on the page: renders the rich paragraphs,
 * double-click enters in-place editing (plain text per paragraph; each line
 * keeps its paragraph format and first-run styling), blur commits. PAGE /
 * NUMPAGES sentinels edit as visible {PAGE} / {NUMPAGES} tokens.
 */
export function HeaderFooterArea({
  kind,
  value,
  images,
  readOnly,
  onCommit,
  onInsertImage,
  onRemoveImage,
  onReplaceImage,
  onResizeImage,
  pageNo,
  pageTotal,
  style,
  editNonce,
}: {
  kind: 'header' | 'footer'
  value: HfValue
  /** logo and other images in the part, display-only (text edits do not affect their saved bytes) */
  images?: HfStripImage[]
  readOnly?: boolean
  onCommit: (next: HfValue) => void
  /** insert a picture from disk into this header/footer (renders the affordance when set) */
  onInsertImage?: (image: NewImage) => void
  /** remove an image that is already in the part (parsed-image order) */
  onRemoveImage?: (index: number) => void
  /** swap an image that is already in the part for one from disk (parsed-image order) */
  onReplaceImage?: (index: number, image: NewImage) => void
  /** commit a new size for an image already in the part (parsed-image order) */
  onResizeImage?: (index: number, widthPx: number, heightPx: number) => void
  /** Page number shown for '#' (may be a section-formatted string); the continuous-flow canvas has no real page number, defaults to 1 */
  pageNo?: number | string
  /** Total page count shown for TOTAL_PAGES_MARK (NUMPAGES field), defaults to 1 */
  pageTotal?: number
  /** geometry override: on differing-width sections the strip is pinned to its own section's box */
  style?: React.CSSProperties
  /** bump to open editing from outside (double-click on a gap's display-only copy); a change while already editing re-opens on the current value, discarding the uncommitted edit */
  editNonce?: number
}) {
  const { t, lang } = useI18n()
  const [editing, setEditing] = useState(false)
  const [inserting, setInserting] = useState(false)
  const [colorOpen, setColorOpen] = useState(false)
  const editRef = useRef<HTMLDivElement>(null)
  const cancelRef = useRef(false)
  const initialHtmlRef = useRef('')
  /** live size of the image being drag-resized (index + displayed px) */
  const [resizing, setResizing] = useState<{ index: number; w: number; h: number } | null>(null)
  const resizingRef = useRef<{
    index: number
    w0: number
    h0: number
    x0: number
    ratio: number
  } | null>(null)

  /** drag the corner handle: live preview, aspect-ratio locked, commit on release */
  const startResize = (index: number, img: HTMLImageElement, e: React.MouseEvent) => {
    e.preventDefault()
    e.stopPropagation()
    const rect = img.getBoundingClientRect()
    resizingRef.current = {
      index,
      w0: rect.width,
      h0: rect.height,
      x0: e.clientX,
      ratio: rect.height / rect.width,
    }
    setResizing({ index, w: rect.width, h: rect.height })
    const move = (ev: MouseEvent) => {
      const st = resizingRef.current
      if (!st) return
      const w = Math.max(24, st.w0 + (ev.clientX - st.x0))
      setResizing({ index: st.index, w, h: w * st.ratio })
    }
    const up = (ev: MouseEvent) => {
      window.removeEventListener('mousemove', move)
      window.removeEventListener('mouseup', up)
      const st = resizingRef.current
      resizingRef.current = null
      const w = Math.max(24, st ? st.w0 + (ev.clientX - st.x0) : 0)
      setResizing(null)
      // zoom-aware: on-screen px include the canvas zoom factor
      const zoom = parseFloat(getComputedStyle(img.closest('.doc-zoom') ?? img).zoom || '1') || 1
      onResizeImage?.(index, Math.round(w / zoom), Math.round((w * (st?.ratio ?? 1)) / zoom))
    }
    window.addEventListener('mousemove', move)
    window.addEventListener('mouseup', up)
  }
  // An image-only part (logo header) has no text paragraphs: hfParasOf would
  // still synthesize one empty line, which stacks a stray blank line over the
  // body's first paragraph (it even intercepts its clicks).
  const imageOnly = !value.text && !value.pageNumber && !value.paras?.length
  const paras = imageOnly ? [] : hfParasOf(value)

  const insertImage = async () => {
    if (inserting) return
    setInserting(true)
    try {
      const picked = await pickImageFromDisk()
      if (!picked) return
      onInsertImage?.(picked)
    } catch {
      /* dialog cancelled or image unreadable */
    } finally {
      setInserting(false)
    }
  }

  /** swap an in-file image for one from disk (same pipeline as insert, keeps the paragraph) */
  const replaceImage = async (index: number) => {
    if (inserting) return
    setInserting(true)
    try {
      const picked = await pickImageFromDisk()
      if (!picked) return
      onReplaceImage?.(index, picked)
    } catch {
      /* dialog cancelled or image unreadable */
    } finally {
      setInserting(false)
    }
  }

  // External entry (double-click on a gap's display-only copy): open the editor
  // on the variant App switched to. editNonce is also a re-inject dependency
  // below, so a bump while editing re-opens on the fresh value (explicit intent
  // to switch strips discards the uncommitted edit).
  useEffect(() => {
    if (editNonce) setEditing(true)
  }, [editNonce])

  // The editing surface is a standalone element: content is injected here and React
  // does not manage its children; after commit the whole element unmounts, so text
  // nodes produced while typing don't linger (keeps section/variant switches clean)
  useEffect(() => {
    if (!editing) return
    const el = editRef.current
    if (!el) return
    // rich editable HTML: per-run styled spans + per-paragraph alignment blocks
    el.innerHTML = hfToEditHtml(value)
    cancelRef.current = false
    initialHtmlRef.current = el.innerHTML
    // execCommand color/lists wrap the selection in CSS spans (not <font> tags)
    try {
      document.execCommand('styleWithCSS', false, 'true')
    } catch {
      /* non-fatal: foreColor then falls back to <font> */
    }
    el.focus()
    const sel = window.getSelection()
    if (sel) {
      sel.selectAllChildren(el)
      sel.collapseToEnd()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editing, editNonce])

  const commit = () => {
    const el = editRef.current
    setEditing(false)
    setColorOpen(false)
    if (!el) return
    if (cancelRef.current) {
      cancelRef.current = false
      return
    }
    // Compare rich HTML: a formatting-only change (e.g. bold) leaves innerText
    // identical, so the old text-only guard would silently drop it.
    if (el.innerHTML === initialHtmlRef.current) return
    onCommit(hfEditDomToValue(value, el))
  }

  // ── formatting toolbar actions (execCommand for the classic toggles; span
  //    wrap for font size/family, which execCommand maps to 1-7 / <font>) ──

  const focusEdit = () => editRef.current?.focus()

  const exec = (cmd: string, val?: string): void => {
    focusEdit()
    try {
      document.execCommand(cmd, false, val)
    } catch {
      /* ignored */
    }
  }

  /** wrap the current selection in a styled span; restores the caret after it */
  const wrapSelection = (css: string): void => {
    const el = editRef.current
    const sel = window.getSelection()
    if (!el || !sel || sel.rangeCount === 0 || sel.isCollapsed) return
    const range = sel.getRangeAt(0)
    if (!el.contains(range.commonAncestorContainer)) return
    const span = document.createElement('span')
    span.setAttribute('style', css)
    try {
      range.surroundContents(span)
      sel.removeAllRanges()
      const r = document.createRange()
      r.selectNodeContents(span)
      r.collapse(false)
      sel.addRange(r)
    } catch {
      /* selection crosses run boundaries: leave it unchanged */
    }
  }

  const alignCmd: Record<'left' | 'center' | 'right', string> = {
    left: 'justifyLeft',
    center: 'justifyCenter',
    right: 'justifyRight',
  }

  const fontOptions = fontFamiliesFor(lang)

  const display = (text: string) => {
    const t = text
      .replaceAll(TOTAL_PAGES_MARK, String(pageTotal ?? 1))
      .replaceAll(PAGE_MARK, String(pageNo ?? 1))
    return hfUsesLegacyHash(value) ? t.replace('#', String(pageNo ?? 1)) : t
  }

  return (
    <div
      className={`page-hf page-hf-${kind}${editing ? ' page-hf-editing' : ''}`}
      style={{
        ...style,
        // Word-style gap between an image-only logo and the body: a margin on
        // the strip itself, so the dashed rule stays under the logo and the
        // body starts below the gap (the push-down estimate reserves it)
        ...(imageOnly && !editing ? { marginBottom: 12 } : {}),
      }}
      data-tip={
        readOnly
          ? undefined
          : t(kind === 'header' ? 'appDblclickEditHeader' : 'appDblclickEditFooter') +
            (value.pageNumber
              ? hfUsesLegacyHash(value)
                ? t('appHfPageNumHint')
                : t('appHfPageNumHint').replace('#', PAGE_TOKEN)
              : '')
      }
      onDoubleClick={() => {
        if (!readOnly && !editing) setEditing(true)
      }}
    >
      {images && images.some((im) => !im.floating) && (
        <div
          className="page-hf-images"
          contentEditable={false}
          style={{
            ...(images.find((im) => !im.floating)?.align === 'right'
              ? { justifyContent: 'flex-end' }
              : images.find((im) => !im.floating)?.align === 'center'
                ? { justifyContent: 'center' }
                : {}),
          }}
        >
          {images
            .filter((img) => !img.floating)
            .map((img, i) => (
              <span key={i} className="page-hf-imgwrap" data-resizable="1">
                <img
                  src={img.dataUrl}
                  alt=""
                  draggable={false}
                  style={{
                    ...(resizing && resizing.index === img.parsedIndex
                      ? { width: resizing.w, height: resizing.h }
                      : {}),
                    ...(!(resizing && resizing.index === img.parsedIndex) && img.widthPx
                      ? { width: img.widthPx }
                      : {}),
                    ...(!(resizing && resizing.index === img.parsedIndex) && img.heightPx
                      ? { height: img.heightPx }
                      : {}),
                  }}
                />
                {img.parsedIndex !== undefined && !readOnly && onResizeImage && (
                  <span
                    className="page-hf-imghandle"
                    onMouseDown={(e) =>
                      startResize(
                        img.parsedIndex!,
                        e.currentTarget.previousElementSibling as HTMLImageElement,
                        e,
                      )
                    }
                  />
                )}
                {/* images already in the file: replace/delete without leaving the strip
                    (their paragraphs were display-only before imageEdits existed) */}
                {img.parsedIndex !== undefined && !readOnly && (
                  <span className="page-hf-imgtools" onMouseDown={(e) => e.preventDefault()}>
                    {onReplaceImage && (
                      <button
                        data-tip={t('appHfImageReplace')}
                        onClick={() => void replaceImage(img.parsedIndex!)}
                      >
                        <IconReplaceImage size={13} />
                      </button>
                    )}
                    {onRemoveImage && (
                      <button
                        data-tip={t('appHfImageRemove')}
                        onClick={() => onRemoveImage(img.parsedIndex!)}
                      >
                        <IconClose size={13} />
                      </button>
                    )}
                  </span>
                )}
              </span>
            ))}
        </div>
      )}
      {!readOnly && onInsertImage && !editing && (
        <button
          className="page-hf-insert-image"
          data-tip={t('ribbonPictureTip')}
          disabled={inserting}
          onClick={() => void insertImage()}
        >
          {t('ribbonPicture')}
        </button>
      )}
      {editing && (
        <div
          className="page-hf-toolbar"
          onMouseDown={(e) => e.preventDefault()}
          onClick={(e) => e.stopPropagation()}
        >
          <select
            className="page-hf-tb-font"
            data-tip={t('ribbonFontFamilyTip')}
            value=""
            onChange={(e) => {
              const f = e.target.value
              if (f) wrapSelection(`font-family:${JSON.stringify(f)}`)
              e.target.value = ''
            }}
          >
            <option value="" disabled>
              {t('ribbonFontFamilyTip')}
            </option>
            {fontOptions.map((f) => (
              <option key={f} value={f} style={{ fontFamily: f }}>
                {f}
              </option>
            ))}
          </select>
          <select
            className="page-hf-tb-size"
            data-tip={t('ribbonFontSizeTip')}
            value=""
            onChange={(e) => {
              const n = Number(e.target.value)
              if (Number.isFinite(n) && n > 0) wrapSelection(`font-size:${n}pt`)
              e.target.value = ''
            }}
          >
            <option value="" disabled>
              {t('ribbonFontSizeTip')}
            </option>
            {HF_FONT_SIZES.map((s) => (
              <option key={s} value={s}>
                {s}
              </option>
            ))}
          </select>
          <span className="page-hf-tb-sep" />
          <button data-tip={t('ribbonBoldTip')} onClick={() => exec('bold')}>
            <b>B</b>
          </button>
          <button data-tip={t('ribbonItalicTip')} onClick={() => exec('italic')}>
            <i>I</i>
          </button>
          <button data-tip={t('ribbonUnderlineTip')} onClick={() => exec('underline')}>
            <u>U</u>
          </button>
          <span className="page-hf-tb-sep" />
          <span className="page-hf-tb-color-wrap">
            <button
              className="page-hf-tb-color"
              data-tip={t('ribbonFontColor')}
              onClick={() => setColorOpen((v) => !v)}
            >
              <IconFontColorA size={14} />
            </button>
            {colorOpen && (
              <span className="page-hf-tb-colors" onMouseDown={(e) => e.preventDefault()}>
                {HF_COLORS.map((c) => (
                  <button
                    key={c}
                    className="page-hf-tb-swatch"
                    style={{ background: `#${c}` }}
                    data-tip={`#${c}`}
                    onClick={() => {
                      exec('foreColor', `#${c}`)
                      setColorOpen(false)
                    }}
                  />
                ))}
              </span>
            )}
          </span>
          <span className="page-hf-tb-sep" />
          <button data-tip={t('ribbonAlignLeftTip')} onClick={() => exec(alignCmd.left)}>
            <IconAlignLeft size={14} />
          </button>
          <button data-tip={t('ribbonAlignCenterTip')} onClick={() => exec(alignCmd.center)}>
            <IconAlignCenter size={14} />
          </button>
          <button data-tip={t('ribbonAlignRightTip')} onClick={() => exec(alignCmd.right)}>
            <IconAlignRight size={14} />
          </button>
          <span className="page-hf-tb-sep" />
          <button data-tip={t('ribbonPageNumber')} onClick={() => exec('insertText', PAGE_TOKEN)}>
            <IconPageNumber size={14} />
          </button>
          <button data-tip={t('hfTotalPagesTip')} onClick={() => exec('insertText', TOTAL_TOKEN)}>
            ∑
          </button>
          <span className="page-hf-tb-sep" />
          <button className="page-hf-tb-close" onClick={() => editRef.current?.blur()}>
            <IconClose size={14} /> {t('ribbonClose')}
          </button>
        </div>
      )}
      {editing ? (
        <div
          ref={editRef}
          className="page-hf-edit-surface"
          contentEditable
          suppressContentEditableWarning
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === 'Escape') {
              e.preventDefault()
              e.stopPropagation()
              cancelRef.current = true
              ;(e.target as HTMLElement).blur()
            }
          }}
        />
      ) : (
        <HfContent paras={paras} display={display} />
      )}
    </div>
  )
}

function HfContent({
  paras,
  display,
}: {
  paras: HfParagraph[]
  display: (text: string) => string
}) {
  return (
    <>
      {paras.map((para, i) =>
        para.cells ? (
          // layout-table row: read-only flex columns (excluded from text editing)
          <div key={i} className="page-hf-para page-hf-row">
            {para.cells.map((cell, j) => (
              <div
                key={j}
                className="page-hf-cell"
                style={{
                  ...(cell.widthPct ? { width: `${cell.widthPct}%` } : {}),
                  // document content color (w:shd), theme-independent
                  ...(cell.fill ? { backgroundColor: `#${cell.fill}` } : {}),
                  ...(cell.align
                    ? {
                        textAlign:
                          cell.align === 'left' || cell.align === 'center' || cell.align === 'right'
                            ? cell.align
                            : ('justify' as const),
                      }
                    : {}),
                }}
              >
                {/* one block line per cell paragraph (mirrors makeGapHfEl) */}
                {(cell.paras.length > 0 ? cell.paras : [[]]).map((runs, k) => (
                  <div key={k} className="page-hf-cell-para">
                    {runs.length === 0 ? ' ' : null}
                    {runs.map((run, l) => (
                      <span key={l} style={runStyle(run)}>
                        {run.image && (
                          <img
                            className="page-hf-cell-img"
                            src={run.image.dataUrl}
                            alt=""
                            draggable={false}
                            style={{
                              ...(run.image.widthPx ? { width: run.image.widthPx } : {}),
                              ...(run.image.heightPx ? { height: run.image.heightPx } : {}),
                            }}
                          />
                        )}
                        {display(run.text)}
                      </span>
                    ))}
                  </div>
                ))}
              </div>
            ))}
          </div>
        ) : (
          (() => {
            const tabbed = hfTabSegments(para, display)
            if (!tabbed) {
              return (
                <div
                  key={i}
                  className={`page-hf-para${para.frameXAlign ? ' page-hf-frame' : ''}`}
                  style={paraStyle(para)}
                >
                  {para.runs.length === 0 ? ' ' : null}
                  {para.runs.map((run, j) => (
                    <span key={j} style={runStyle(run)}>
                      {display(run.text)}
                    </span>
                  ))}
                </div>
              )
            }
            const leadIndent = hfLeadIndentCss(tabbed)
            return (
              <div
                key={i}
                className={`page-hf-para page-hf-tabbed${para.frameXAlign ? ' page-hf-frame' : ''}`}
                style={{
                  ...paraStyle(para),
                  ...(tabbed.minHeightPt ? { minHeight: `${tabbed.minHeightPt}pt` } : {}),
                  // tab layout happens in left-aligned space; w:jc becomes an explicit shift
                  textAlign: 'left',
                  ...(leadIndent ? { textIndent: leadIndent } : {}),
                }}
              >
                {tabbed.lead.map((run, j) => (
                  <span key={j} style={runStyle(run)}>
                    {display(run.text)}
                  </span>
                ))}
                {tabbed.segments.map((seg, k) => (
                  <span
                    key={`t${k}`}
                    className={`page-hf-tabseg page-hf-tabseg-${seg.anchor}`}
                    style={{ left: hfSegLeftCss(seg, tabbed) }}
                  >
                    {seg.runs.map((run, j) => (
                      <span key={j} style={runStyle(run)}>
                        {display(run.text)}
                      </span>
                    ))}
                  </span>
                ))}
              </div>
            )
          })()
        ),
      )}
    </>
  )
}
