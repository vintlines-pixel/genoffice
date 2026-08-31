/**
 * Standalone PDF scroll perf probe, split in two modes because this sandbox
 * cannot spawn child processes from node (spawn EPERM):
 *   1) `node e2e/pdf-scroll-perf.probe.mjs gen <pdfPath>`  — write the test PDF
 *   2) launch electron via Start-Process with --remote-debugging-port
 *   3) `node e2e/pdf-scroll-perf.probe.mjs drive <cdpPort>` — connect over CDP,
 *      instrument the renderer, drive real wheel scrolling, report jank
 */
import { chromium } from 'playwright'
import { writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'

const mode = process.argv[2]

const PAGES = 60
const LINES_PER_PAGE = 48
const LINE =
  'The quick brown fox jumps over the lazy dog 0123456789 perf probe line of text.'

function textHeavyPdf() {
  const offsets = []
  let body = '%PDF-1.4\n'
  const push = (obj) => {
    offsets.push(body.length)
    body += `${offsets.length} 0 obj\n${obj}\nendobj\n`
  }
  push('<</Type/Catalog/Pages 2 0 R>>') // obj 1
  const kids = []
  for (let p = 0; p < PAGES; p++) kids.push(`${4 + p * 2} 0 R`)
  push(`<</Type/Pages/Kids[${kids.join('')}] /Count ${PAGES}>>`) // obj 2
  push('<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>') // obj 3
  for (let p = 0; p < PAGES; p++) {
    const lines = []
    for (let l = 0; l < LINES_PER_PAGE; l++)
      lines.push(`(page ${p + 1} line ${l + 1}: ${LINE}) Tj T*`)
    const content = `BT /F1 11 Tf 13.5 TL 36 756 Td\n${lines.join('\n')}\nET\n`
    push(
      `<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]/Resources<</Font<</F1 3 0 R>>>>/Contents ${
        5 + p * 2
      } 0 R>>`,
    )
    push(`<</Length ${content.length}>>\nstream\n${content}endstream`)
  }
  const xrefStart = body.length
  body += `xref\n0 ${offsets.length + 1}\n0000000000 65535 f \n`
  for (const off of offsets) body += `${String(off).padStart(10, '0')} 00000 n \n`
  body += `trailer\n<</Size ${offsets.length + 1}/Root 1 0 R>>\nstartxref\n${xrefStart}\n%%EOF\n`
  return Buffer.from(body, 'latin1')
}

async function domStats(page) {
  return page.evaluate(() => ({
    pages: document.querySelectorAll('.pdf-page').length,
    canvases: document.querySelectorAll('.pdf-page-content canvas').length,
    textSpans: document.querySelectorAll('.textLayer span').length,
    scrollTop: document.querySelector('.pdf-scroll')?.scrollTop ?? 0,
  }))
}

const run = async () => {
  if (mode === 'gen') {
    const out = resolve(process.argv[3])
    await writeFile(out, textHeavyPdf())
    console.log('wrote', out)
    return
  }
  if (mode !== 'drive') throw new Error('usage: gen <pdf> | drive <cdpPort>')
  const port = Number(process.argv[3])

  const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
  const ctx = browser.contexts()[0]
  let editor = ctx.pages().find((p) => p.url().includes('pdf/out'))
  for (let i = 0; !editor && i < 150; i++) {
    await new Promise((r) => setTimeout(r, 200))
    editor = ctx.pages().find((p) => p.url().includes('pdf/out'))
  }
  if (!editor) {
    console.log('pages:', ctx.pages().map((p) => p.url()).join('\n'))
    throw new Error('pdf editor page never appeared')
  }
  console.log('editor page:', editor.url())
  await editor.locator('.pdf-page').first().waitFor({ state: 'visible', timeout: 30000 })
  await editor.waitForTimeout(2500)
  console.log('initial DOM:', JSON.stringify(await domStats(editor)))

  const start = () =>
    editor.evaluate(() => {
      const w = window
      w.__perf = { longTasks: [], frames: [] }
      new PerformanceObserver((list) => {
        for (const e of list.getEntries())
          w.__perf.longTasks.push({ dur: e.duration, name: e.name })
      }).observe({ entryTypes: ['longtask'] })
      let last = performance.now()
      const tick = (t) => {
        w.__perf.frames.push(t - last)
        last = t
        requestAnimationFrame(tick)
      }
      requestAnimationFrame(tick)
    })
  const perf = () => editor.evaluate(() => window.__perf)
  const reset = () => editor.evaluate(() => (window.__perf = { longTasks: [], frames: [] }))

  const box = await editor.locator('.pdf-scroll').boundingBox()
  await editor.mouse.move(box.x + box.width / 2, box.y + box.height / 2)

  // ── phase 1: continuous wheel scroll down ──
  await start()
  for (let i = 0; i < 80; i++) {
    await editor.mouse.wheel(0, 240)
    await editor.waitForTimeout(35)
  }
  await editor.waitForTimeout(1200)
  summarize('scroll down (80 wheel ticks)', await perf(), await domStats(editor))

  // ── phase 2: scroll back up (revisit released pages) ──
  await reset()
  for (let i = 0; i < 80; i++) {
    await editor.mouse.wheel(0, -240)
    await editor.waitForTimeout(35)
  }
  await editor.waitForTimeout(1200)
  summarize('scroll back up (80 wheel ticks)', await perf(), await domStats(editor))

  // ── phase 3: zoom in (forces re-render of visible pages) ──
  await reset()
  await editor.keyboard.press('Control+=')
  await editor.waitForTimeout(2500)
  summarize('zoom in once', await perf(), await domStats(editor))

  await browser.close().catch(() => {})
}

run().catch((e) => {
  console.error(e)
  process.exitCode = 1
})
