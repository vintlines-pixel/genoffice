/** sheets export-pdf on the quotation file: picture header + column pagination check */
import { _electron as electron } from '@playwright/test'
import { execSync } from 'node:child_process'
import { mkdtemp, copyFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve, dirname } from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import { autoDontSaveOnClose } from './lib/auto-dont-save.mjs'

const HERE = dirname(fileURLToPath(import.meta.url))
const ROOT = resolve(HERE, '../..')
const SHELL_DIR = join(ROOT, 'apps/shell')
const FIXTURE = join(ROOT, 'e2e/artifacts/quotation.xlsx')
const OUT = join(ROOT, 'e2e/artifacts/sheets-pdf-check/exported3.pdf')
const require = createRequire(join(SHELL_DIR, 'package.json'))
const { ELECTRON_RUN_AS_NODE: _n, ...hostEnv } = process.env

const scratch = await mkdtemp(join(tmpdir(), 'go-spdf3-'))
const docPath = join(scratch, 'quote.xlsx')
const pdfPath = join(scratch, 'quote.pdf')
await copyFile(FIXTURE, docPath)
const app = await electron.launch({
  executablePath: require('electron'),
  args: [SHELL_DIR, docPath],
  env: { ...hostEnv, GENOFFICE_USER_DATA: await mkdtemp(join(tmpdir(), 'go-spdf3-ud-')), GENOFFICE_LANG: 'zh', GENOFFICE_E2E_VIDEO: '0' },
})
try {
  await autoDontSaveOnClose(app)
  app.process().stdout?.on('data', (d) => {
    const t = d.toString()
    if (t.includes('[hdrimg]')) console.log('[main]', t.trim())
  })
  let page
  for (let i = 0; i < 250 && !page; i++) {
    for (const w of app.windows()) if (w.url().includes('sheets/out')) page = w
    if (!page) await new Promise((r) => setTimeout(r, 150))
  }
  page.on('console', (m) => {
    const t = m.text()
    if (m.type() === 'error' || /Export PDF|error|failed/i.test(t)) console.log('[console]', m.type(), t.slice(0, 250))
  })
  await page.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, { timeout: 30_000 })
  await page.waitForTimeout(3500)
  await app.evaluate(({ dialog }, p) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: p })
    dialog.showSaveDialogSync = () => p
  }, pdfPath)
  const exportBtn = page.locator('button[data-tip="导出为 PDF"], button[aria-label="导出为 PDF"]').first()
  await exportBtn.click()
  const confirm = page.locator('.primary-action', { hasText: '导出' }).first()
  await confirm.waitFor({ state: 'visible', timeout: 8000 }).catch(() => {})
  const vis = await confirm.isVisible().catch(() => false)
  console.log('confirm visible:', vis)
  if (!vis) {
    const status = await page.evaluate(() => document.querySelector('[class*=status], [class*=message]')?.textContent?.slice(0, 150)).catch(() => '')
    console.log('status after click:', status)
  }
  await confirm.click()
  await page.waitForTimeout(2500)
  const status2 = await page.evaluate(() => document.querySelector('[class*=status], [class*=message]')?.textContent?.slice(0, 150)).catch(() => '')
  console.log('status after export:', status2)
  for (let i = 0; i < 60; i++) {
    await page.waitForTimeout(1000)
    try {
      const { statSync } = await import('node:fs')
      let size = 0
      try {
        size = statSync(pdfPath).size
      } catch {}
      if (size > 1000) {
        await copyFile(pdfPath, OUT)
        console.log('PDF exported:', size, 'bytes')
        process.exit(0)
      }
    } catch {}
  }
  console.log('PDF export timed out')
  process.exitCode = 1
} finally {
  await app.close().catch(() => {})
}
