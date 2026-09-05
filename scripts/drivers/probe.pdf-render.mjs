/** render a PDF's first pages via system Edge for visual comparison */
import { chromium } from '@playwright/test'
import { join, resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const HERE = dirname(fileURLToPath(import.meta.url))
const ROOT = resolve(HERE, '../..')
const pdfArg = process.argv[2] ?? join(ROOT, 'e2e/artifacts/sheets-pdf-check/exported6.pdf')
const outDir = process.argv[3] ?? join(ROOT, 'e2e/artifacts/sheets-pdf-check')
const pdfUrl = 'file:///' + pdfArg.split('\\').join('/')
const browser = await chromium.launch({ channel: 'msedge' })
const page = await browser.newPage({ viewport: { width: 900, height: 1300 } })
await page.goto(pdfUrl)
await page.waitForTimeout(3000)
await page.screenshot({ path: join(outDir, 'pdf-view-hdr2.png') })
await browser.close()
console.log('rendered')
