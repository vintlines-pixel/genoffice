/** #6专项评估测量：Enter 在页尾时，到底什么在动（灰带位置/下一页顶边/纸面高度/滚动） */
import { _electron as electron } from '@playwright/test'
import { mkdtemp, copyFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve, dirname } from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import { autoDontSaveOnClose } from './lib/auto-dont-save.mjs'

const HERE = dirname(fileURLToPath(import.meta.url))
const ROOT = resolve(HERE, '../..')
const SHELL_DIR = join(ROOT, 'apps/shell')
const FIXTURE = join(ROOT, 'e2e/artifacts/header-img-test.docx')
const require = createRequire(join(SHELL_DIR, 'package.json'))
const { ELECTRON_RUN_AS_NODE: _n, ...hostEnv } = process.env
const scratch = await mkdtemp(join(tmpdir(), 'go-ev-'))
const docPath = join(scratch, 't.docx')
await copyFile(FIXTURE, docPath)
const app = await electron.launch({
  executablePath: require('electron'),
  args: [SHELL_DIR, docPath],
  env: { ...hostEnv, GENOFFICE_USER_DATA: await mkdtemp(join(tmpdir(), 'go-ev-ud-')), GENOFFICE_LANG: 'zh', GENOFFICE_E2E_VIDEO: '0' },
})
try {
  await autoDontSaveOnClose(app)
  let page
  for (let i = 0; i < 250 && !page; i++) {
    for (const w of app.windows()) if (w.url().includes('docs/out')) page = w
    if (!page) await new Promise((r) => setTimeout(r, 150))
  }
  await page.locator('.doc-page').first().waitFor({ state: 'visible', timeout: 30_000 })
  await page.waitForTimeout(3000)

  // caret to the last paragraph above the first gap (page 1 tail)
  const idx = await page.evaluate(() => {
    const gapRect = document.querySelector('.page-gap').getBoundingClientRect()
    const kids = [...document.querySelectorAll('.ProseMirror > *')]
    for (let i = kids.length - 1; i >= 0; i--) {
      const r = kids[i].getBoundingClientRect()
      if (r.bottom <= gapRect.top + 2 && r.height > 0 && (kids[i].textContent ?? '').trim()) return i
    }
    return -1
  })
  await page.locator('.ProseMirror > *').nth(idx).click()
  await page.keyboard.press('End')
  await page.waitForTimeout(400)

  // rAF sampling: what visually moves when Enter lands
  await page.evaluate(() => {
    window.__geo = []
    const t0 = performance.now()
    const sample = () => {
      const gap = document.querySelector('.page-gap')
      const gapR = gap?.getBoundingClientRect()
      const bandTop = gapR ? gapR.top + parseFloat(getComputedStyle(gap).getPropertyValue('--gap-mb') || '0') : -1
      // first block below the gap = page 2's first visible block
      let page2Top = -1
      const pm = document.querySelector('.ProseMirror')
      for (const el of pm.children) {
        const r = el.getBoundingClientRect()
        if (gapR && r.top >= gapR.top + 2 && r.height > 0) {
          page2Top = Math.round(r.top)
          break
        }
      }
      window.__geo.push({
        t: Math.round(performance.now() - t0),
        gapH: gapR ? Math.round(gapR.height) : -1,
        bandTop: Math.round(bandTop),
        page2Top,
        scroll: Math.round(document.querySelector('.editor-scroll').scrollTop),
      })
      if (window.__geoOn) requestAnimationFrame(sample)
    }
    window.__geoOn = true
    requestAnimationFrame(sample)
  })
  for (let k = 0; k < 3; k++) {
    await page.keyboard.press('Enter')
    await page.waitForTimeout(700)
  }
  await page.waitForTimeout(1500)
  const geo = await page.evaluate(() => {
    window.__geoOn = false
    return window.__geo
  })
  // print only frames where ANY tracked value changed
  let prev = null
  for (const g of geo) {
    if (!prev || g.gapH !== prev.gapH || g.bandTop !== prev.bandTop || g.page2Top !== prev.page2Top || g.scroll !== prev.scroll) {
      console.log(`t=${g.t}ms gapH=${g.gapH} bandTop=${g.bandTop} page2Top=${g.page2Top} scroll=${g.scroll}`)
    }
    prev = g
  }
  console.log('frames sampled:', geo.length)
} finally {
  await app.close().catch(() => {})
}
