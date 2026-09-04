/// PDF export & preview: renders the print HTML (laid out by the renderer) in
/// a hidden scripting-disabled window and writes webContents.printToPDF's
/// output where the save dialog points. Preview renders the same payload to a
/// throwaway PDF and shows it in a persistent viewer window (Chromium's
/// built-in PDF viewer), so the user sees the real pagination before saving.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { BrowserWindow, dialog } from 'electron'

import { showSaveDialogWithMemory } from '@genoffice/electron-utils'

import type { IpcMainInvokeEvent } from 'electron'
import type {
  WorkbookExportPdfRequest,
  WorkbookExportPdfResult,
  WorkbookPreviewPdfResult,
} from '../shared/desktop-api'

/// printToPDF options shared by export and preview.
function printOptions(request: WorkbookExportPdfRequest): Electron.PrintToPDFOptions {
  const displayHeaderFooter =
    request.headerTemplate !== undefined || request.footerTemplate !== undefined
  return {
    landscape: request.landscape,
    pageSize: request.pageSize,
    margins: request.margins,
    scale: request.scale,
    printBackground: true,
    ...(request.pageRanges !== undefined ? { pageRanges: request.pageRanges } : {}),
    ...(displayHeaderFooter
      ? {
          displayHeaderFooter: true,
          // Chromium falls back to its own date/title header when a
          // template is missing, so always pass both.
          headerTemplate: request.headerTemplate ?? '<span></span>',
          footerTemplate: request.footerTemplate ?? '<span></span>',
        }
      : {}),
  }
}

/// Loads the print HTML into a hidden scripting-disabled window.
async function withPrintWindow<T>(
  request: WorkbookExportPdfRequest,
  run: (window: BrowserWindow) => Promise<T>,
): Promise<T> {
  const workDir = await mkdtemp(join(tmpdir(), 'ai-excel-pdf-'))
  const window = new BrowserWindow({
    show: false,
    webPreferences: { sandbox: true, javascript: false },
  })
  try {
    await writeFile(join(workDir, 'print.html'), request.html, 'utf8')
    await window.loadFile(join(workDir, 'print.html'))
    return await run(window)
  } finally {
    window.destroy()
    await rm(workDir, { recursive: true, force: true })
  }
}

export async function exportPdf(
  event: IpcMainInvokeEvent,
  request: WorkbookExportPdfRequest,
): Promise<WorkbookExportPdfResult> {
  const parent = BrowserWindow.fromWebContents(event.sender)
  const dialogOptions = {
    defaultPath: request.fileName,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  }
  const selection = await showSaveDialogWithMemory(dialog, parent, dialogOptions)
  const outputPath = selection.canceled ? null : selection.filePath
  if (!outputPath) return { canceled: true }

  return withPrintWindow(request, async (window) => {
    const pdf = await window.webContents.printToPDF(printOptions(request))
    await writeFile(outputPath, pdf)
    return { canceled: false, path: outputPath }
  })
}

/// The persistent preview window; one at a time, reused across previews.
let previewWindow: BrowserWindow | null = null
let previewPdfPath: string | null = null

export async function previewPdf(
  request: WorkbookExportPdfRequest,
): Promise<WorkbookPreviewPdfResult> {
  try {
    // A stable path: each preview overwrites it, and the viewer reloads it.
    previewPdfPath ??= join(tmpdir(), 'ai-excel-pdf-preview.pdf')
    const pdf = await withPrintWindow(request, (window) =>
      window.webContents.printToPDF(printOptions(request)),
    )
    await writeFile(previewPdfPath, pdf)
    if (previewWindow === null || previewWindow.isDestroyed()) {
      previewWindow = new BrowserWindow({
        show: false,
        title: request.fileName,
        width: 1_024,
        height: 800,
      })
      previewWindow.on('closed', () => {
        previewWindow = null
      })
      await previewWindow.loadFile(previewPdfPath, { query: { v: String(Date.now()) } })
      previewWindow.show()
    } else {
      // The query busts the reuse case: same-path loadFile would otherwise
      // be treated as a same-document navigation and keep the old render.
      await previewWindow.loadFile(previewPdfPath, { query: { v: String(Date.now()) } })
      if (previewWindow.isMinimized()) previewWindow.restore()
      previewWindow.focus()
    }
    return { ok: true }
  } catch (error: unknown) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}
