/** Excel print: menu Print (Ctrl+P) → payload → printWorkbook IPC → system
 *  print dialog. The real dialog is app-modal and blocks automation, so the
 *  probe polls the __printInvoked side-effect marker set inside printWorkbook,
 *  then force-quits the app to dismiss the dialog. */
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
const FIXTURE = join(ROOT, 'e2e/artifacts/quotation.xlsx')
const require = createRequire(join(SHELL_DIR, 'package.json'))
const { ELECTRON_RUN_AS_NODE: _n, ...hostEnv } = process.env
const scratch = await mkdtemp(join(tmpdir(), 'go-sprint-'))
const docPath = join(scratch, 'quote.xlsx')
await copyFile(FIXTURE, docPath)
const DEBUG_PORT = 9331
const app = await electron.launch({
  executablePath: require('electron'),
  args: [SHELL_DIR, docPath],
  env: {
    ...hostEnv,
    GENOFFICE_USER_DATA: await mkdtemp(join(tmpdir(), 'go-sprint-ud-')),
    GENOFFICE_LANG: 'zh',
    GENOFFICE_E2E_VIDEO: '0',
    XLSX_DEBUG_PORT: String(DEBUG_PORT),
  },
})
let sawEvidence = false
try {
  await autoDontSaveOnClose(app)
  let page
  for (let i = 0; i < 250 && !page; i++) {
    for (const w of app.windows()) if (w.url().includes('sheets/out')) page = w
    if (!page) await new Promise((r) => setTimeout(r, 150))
  }
  await page.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, { timeout: 30_000 })
  await page.waitForTimeout(3500)

  // trigger via the dev HTTP hook (same sendSheetsMenuAction path as Ctrl+P)
  const { default: http } = await import('node:http')
  const hookStatus = await new Promise((resolve) => {
    const req = http.get(`http://127.0.0.1:${DEBUG_PORT + 1}/menu?action=print`, (res) => {
      let body = ''
      res.on('data', (c) => (body += c))
      res.on('end', () => resolve(`${res.statusCode} ${body}`))
    })
    req.on('error', (e) => resolve('ERR ' + e.message))
  })
  console.log('HOOK:', hookStatus)

  // poll for the side-effect marker, then force-quit (dialog blocks app.close)
  for (let i = 0; i < 40 && !sawEvidence; i++) {
    await page.waitForTimeout(250)
    sawEvidence = await app.evaluate(() => Boolean(globalThis.__printInvoked))
  }
} finally {
  try {
    app.process().kill()
  } catch {}
  await app.close().catch(() => {})
}
if (sawEvidence) {
  console.log('SHEETS PRINT OK — print pipeline invoked (dialog would open for the user)')
} else {
  console.log('SHEETS PRINT FAILED — print pipeline never invoked')
  process.exitCode = 1
}
