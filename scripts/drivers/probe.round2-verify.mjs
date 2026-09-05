/**
 * Verify fixes for the user's 11-issue round: #1 #2 #4 #7 #8 #9.
 */
import { _electron as electron } from '@playwright/test'
import { mkdtemp, copyFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve, dirname } from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import { autoDontSaveOnClose } from './lib/auto-dont-save.mjs'

const HERE = dirname(fileURLToPath(import.meta.url))
const ROOT = resolve(HERE, '../..')
const SHELL_DIR = join(ROOT, 'apps/shell')
const FIXTURE = join(ROOT, 'e2e/artifacts/header-img-test.docx')
const SHOTS = join(ROOT, 'e2e/artifacts/round2-verify')
const require = createRequire(join(SHELL_DIR, 'package.json'))
const { ELECTRON_RUN_AS_NODE: _n, ...hostEnv } = process.env

const results = []
const check = (name, ok, detail = '') => {
  results.push(`${ok ? 'PASS' : 'FAIL'} ${name}${detail ? ' — ' + detail : ''}`)
}

const run = async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'go-r2-'))
  const docPath = join(scratch, 'quote.docx')
  await copyFile(FIXTURE, docPath)
  const app = await electron.launch({
    executablePath: require('electron'),
    args: [SHELL_DIR, docPath],
    env: { ...hostEnv, GENOFFICE_USER_DATA: await mkdtemp(join(tmpdir(), 'go-r2-ud-')), GENOFFICE_LANG: 'zh', GENOFFICE_E2E_VIDEO: '0' },
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

    // ── #9: save does not scroll ──
    const p3 = page.locator('.ProseMirror p').nth(20)
    await p3.scrollIntoViewIfNeeded()
    await p3.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' x9')
    const before = await page.evaluate(() => Math.round(document.querySelector('.editor-scroll').scrollTop))
    await page.keyboard.press('Control+s')
    await page.waitForTimeout(1800)
    const after = await page.evaluate(() => Math.round(document.querySelector('.editor-scroll').scrollTop))
    check('#9 save keeps scroll position', Math.abs(before - after) <= 1, `scrollTop ${before} -> ${after}`)
    await page.keyboard.press('Control+z')
    await page.waitForTimeout(300)

    // ── #1: dblclick a gap footer copy → overlay opens in place ──
    await page.evaluate(() => {
      const sc = document.querySelector('.editor-scroll')
      sc.scrollTop = sc.scrollHeight / 6
    })
    await page.waitForTimeout(500)
    const gapCopy = page.locator('.page-gap-hf.page-hf-footer').first()
    await gapCopy.scrollIntoViewIfNeeded()
    const gb = await gapCopy.boundingBox()
    const scrollBefore = await page.evaluate(() => Math.round(document.querySelector('.editor-scroll').scrollTop))
    await page.mouse.dblclick(gb.x + gb.width / 2, gb.y + gb.height / 2)
    await page.waitForTimeout(600)
    const overlayState = await page.evaluate(() => {
      const ov = document.querySelector('.page-hf-overlay')
      const surface = ov?.querySelector('.page-hf-edit-surface')
      return {
        overlay: !!ov,
        editing: !!surface,
        focused: document.activeElement === surface,
      }
    })
    const scrollAfterOpen = await page.evaluate(() => Math.round(document.querySelector('.editor-scroll').scrollTop))
    check('#1 overlay opens in place', overlayState.overlay && overlayState.editing, JSON.stringify(overlayState))
    check('#1 viewport does not jump', Math.abs(scrollBefore - scrollAfterOpen) <= 1, `${scrollBefore} -> ${scrollAfterOpen}`)
    await page.screenshot({ path: join(SHOTS, 'v1-overlay.png') })
    await page.keyboard.press('Escape')
    await page.waitForTimeout(400)
    check('#1 Escape closes overlay', await page.evaluate(() => !document.querySelector('.page-hf-overlay')))

    // ── #2: resize handle on the strip image ──
    const stripImg = page.locator('.page-wrap > .page-hf-header .page-hf-images img').first()
    await stripImg.scrollIntoViewIfNeeded()
    const ib = await stripImg.boundingBox()
    await page.mouse.move(ib.x + ib.width / 2, ib.y + ib.height / 2)
    await page.waitForTimeout(300)
    const handle = page.locator('.page-wrap > .page-hf-header .page-hf-imghandle')
    const hb = await handle.boundingBox().catch(() => null)
    if (hb) {
      await page.mouse.move(hb.x + hb.width / 2, hb.y + hb.height / 2)
      await page.mouse.down()
      await page.mouse.move(hb.x + hb.width / 2 - 60, hb.y + hb.height / 2, { steps: 6 })
      await page.mouse.up()
      await page.waitForTimeout(600)
      const resized = await page.evaluate(() => {
        const img = document.querySelector('.page-wrap > .page-hf-header .page-hf-images img')
        return Math.round(img.getBoundingClientRect().width)
      })
      check('#2 image resize applies', resized < Math.round(ib.width) - 30, `width ${Math.round(ib.width)} -> ${resized}`)
    } else {
      check('#2 image resize applies', false, 'handle not visible')
    }
    await page.screenshot({ path: join(SHOTS, 'v2-resized.png') })

    // ── #4: gap between header bottom and first paragraph ──
    const gapPx = await page.evaluate(() => {
      const strip = document.querySelector('.page-wrap > .page-hf-header')
      const p0 = document.querySelector('.ProseMirror p[data-idx="0"]')
      return Math.round(p0.getBoundingClientRect().top - strip.getBoundingClientRect().bottom)
    })
    check('#4 header-to-body gap exists', gapPx >= 8, `gap ${gapPx}px`)

    // ── #7: remove image, then ribbon header text edit must not resurrect it ──
    await stripImg.scrollIntoViewIfNeeded()
    const ib2 = await stripImg.boundingBox()
    await page.mouse.move(ib2.x + ib2.width / 2, ib2.y + ib2.height / 2)
    await page.waitForTimeout(300)
    await page.locator('.page-wrap > .page-hf-header .page-hf-imgtools button[data-tip="删除图片"]').click()
    await page.waitForTimeout(500)
    const afterRemove = await page.evaluate(
      () => document.querySelectorAll('.page-wrap > .page-hf-header .page-hf-images img').length,
    )
    check('#7 image removed from strip', afterRemove === 0, `imgs=${afterRemove}`)
    await page.getByRole('button', { name: '插入', exact: true }).first().click()
    await page.waitForTimeout(300)
    await page.getByRole('button', { name: '页眉', exact: true }).first().click()
    await page.waitForTimeout(300)
    const hfInput = page.locator('.hf-menu input')
    await hfInput.fill('CONF')
    await page.locator('.hf-menu .btn-primary').click()
    await page.waitForTimeout(600)
    const resurrection = await page.evaluate(() => {
      const strip = document.querySelector('.page-wrap > .page-hf-header')
      return {
        imgs: strip?.querySelectorAll('.page-hf-images img').length ?? -1,
        text: strip?.textContent?.replace('图片', '').trim().slice(0, 20),
      }
    })
    check('#7 deleted image stays gone after ribbon text edit', resurrection.imgs === 0, JSON.stringify(resurrection))
    await page.screenshot({ path: join(SHOTS, 'v7-no-resurrection.png') })

    // ── #8: export-pdf opens preview WITHOUT auto-exporting ──
    await app.evaluate(({ dialog }) => {
      let asked = 0
      dialog.showSaveDialog = async (...a) => {
        asked += 1
        return (await dialog.showMessageBox) ? { canceled: true, filePaths: [] } : { canceled: true }
      }
      dialog.showSaveDialogSync = () => {
        asked += 1
        return undefined
      }
      globalThis.__saveDialogCount = () => asked
    })
    await app.evaluate(({ webContents }) => {
      const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('docs/out'))
      wc?.send('menu:command', 'export-pdf')
    })
    await page.waitForTimeout(2500)
    const r8 = await page.evaluate(() => ({
      preview: !!document.querySelector('.pagination-preview'),
    }))
    const dialogCount = await app.evaluate(() => globalThis.__saveDialogCount?.() ?? -1)
    check('#8 preview opens first', r8.preview, JSON.stringify(r8))
    check('#8 no auto save-dialog', dialogCount === 0, `save dialog asked ${dialogCount} times`)
    await page.screenshot({ path: join(SHOTS, 'v8-preview-first.png') })

    console.log(results.join('\n'))
    await writeFile(join(SHOTS, 'results.txt'), results.join('\n'))
  } finally {
    await app.close().catch(() => {})
  }
}

run().catch((err) => {
  console.error(String(err))
  if (results.length) console.log(results.join('\n'))
  process.exitCode = 1
})
