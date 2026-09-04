/**
 * Legacy Word 97-2003 (.doc) import: the editor pipeline only reads .docx,
 * so a .doc is converted next to itself (`report.doc` → `report.docx`) and
 * the converted copy is opened. Conversion tries, in order:
 *
 *   1. LibreOffice (`soffice --headless --convert-to docx`) when installed
 *   2. Word COM automation on Windows when Word is installed
 *   3. Plain-text import (`docToText` → blank-template paragraphs)
 *
 * The original .doc is never modified; an existing `<name>.docx` is opened
 * as-is (it is either a previous conversion or the user's own file).
 */
import { spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import { dirname } from 'node:path'

import { docToText } from '@genoffice/file-parse'

/** legacy Word 97-2003 document (not .docx — anchored at the end) */
export function isLegacyDocPath(filePath: string): boolean {
  return /\.doc$/i.test(filePath)
}

/** `report.doc` → `report.docx` (same folder; the conversion target) */
export function docxTargetFor(docPath: string): string {
  return docPath.replace(/\.doc$/i, '.docx')
}

/** Spawn one child, collect its exit; resolves { code, timedOut }. */
export function runProcess(
  program: string,
  args: readonly string[],
  timeoutMs: number,
): Promise<{ code: number | null; timedOut: boolean }> {
  return new Promise((resolve) => {
    let timedOut = false
    let child: ReturnType<typeof spawn> | null = null
    try {
      child = spawn(program, args, { windowsHide: true, stdio: 'ignore' })
    } catch {
      resolve({ code: null, timedOut: false })
      return
    }
    const timer = setTimeout(() => {
      timedOut = true
      child?.kill()
    }, timeoutMs)
    child.on('error', () => {
      clearTimeout(timer)
      resolve({ code: null, timedOut: false })
    })
    child.on('exit', (code) => {
      clearTimeout(timer)
      resolve({ code, timedOut })
    })
  })
}

/** LibreOffice's soffice executable, or null when not installed. */
export function findSoffice(
  env: NodeJS.ProcessEnv,
  platform: NodeJS.Platform,
  exists: (path: string) => boolean = existsSync,
): string | null {
  const override = env.SOFFICE_PATH
  if (override && exists(override)) return override
  const candidates =
    platform === 'win32'
      ? [
            'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
            'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
          ].filter((p) => exists(p))
      : platform === 'darwin'
        ? ['/Applications/LibreOffice.app/Contents/MacOS/soffice'].filter((p) => exists(p))
        : []
  return candidates[0] ?? (platform !== 'win32' && platform !== 'darwin' ? 'soffice' : null)
}

const CONVERT_TIMEOUT_MS = 120_000

/** LibreOffice conversion; returns whether the target appeared. */
export async function convertDocWithSoffice(
  soffice: string,
  docPath: string,
  targetPath: string,
  run: typeof runProcess = runProcess,
): Promise<boolean> {
  const { code, timedOut } = await run(
    soffice,
    ['--headless', '--norestore', '--convert-to', 'docx', '--outdir', dirname(docPath), docPath],
    CONVERT_TIMEOUT_MS,
  )
  return !timedOut && code === 0 && existsSync(targetPath)
}

/** Word COM conversion (Windows + Word installed); returns whether the target appeared. */
export async function convertDocWithWord(
  docPath: string,
  targetPath: string,
  run: typeof runProcess = runProcess,
): Promise<boolean> {
  // wdFormatXMLDocument = 16 (.docx); ReadOnly:=true, AddToRecentFiles:=false,
  // DisplayAlerts:=0 keep the automation silent.
  const ps =
    `$ErrorActionPreference='Stop';` +
    `$w=New-Object -ComObject Word.Application;$w.Visible=$false;$w.DisplayAlerts=0;` +
    `try{` +
    `$d=$w.Documents.Open('${docPath.replace(/'/g, "''")}', $false, $true);` +
    `$d.SaveAs2('${targetPath.replace(/'/g, "''")}', 16);` +
    `$d.Close($false)}finally{$w.Quit()}`
  const { code, timedOut } = await run(
    'powershell.exe',
    ['-NoProfile', '-NonInteractive', '-Command', ps],
    CONVERT_TIMEOUT_MS,
  )
  return !timedOut && code === 0 && existsSync(targetPath)
}

export interface DocConversionSteps {
  readonly sofficePath: string | null
  readonly convertSoffice: (docPath: string, targetPath: string) => Promise<boolean>
  readonly convertWord: (docPath: string, targetPath: string) => Promise<boolean>
  readonly docToText: (docPath: string) => Promise<string>
  readonly writeTextDocx: (targetPath: string, paragraphs: readonly string[]) => Promise<void>
  readonly exists: (path: string) => boolean
}

/** Production steps; tests inject fakes. */
export function defaultConversionSteps(docToTextImpl: typeof docToText): DocConversionSteps {
  const sofficePath = findSoffice(process.env, process.platform)
  return {
    sofficePath,
    convertSoffice: (docPath, targetPath) =>
      sofficePath
        ? convertDocWithSoffice(sofficePath, docPath, targetPath)
        : Promise.resolve(false),
    convertWord: (docPath, targetPath) =>
      process.platform === 'win32'
        ? convertDocWithWord(docPath, targetPath)
        : Promise.resolve(false),
    docToText: async (docPath) => {
      const { readFile } = await import('node:fs/promises')
      return docToTextImpl(new Uint8Array(await readFile(docPath)))
    },
    writeTextDocx: async (targetPath, paragraphs) => {
      const { buildBlankDocx } = await import('@genoffice/docx-engine')
      const { writeFile } = await import('node:fs/promises')
      await writeFile(targetPath, await buildBlankDocx({ paragraphs }))
    },
    exists: existsSync,
  }
}

export interface LegacyDocOutcome {
  /** how the docx next to the .doc was produced ('existing' = opened as-is) */
  readonly via: 'existing' | 'soffice' | 'word' | 'text'
}

/**
 * Produce `<name>.docx` next to the .doc and return how. Throws when every
 * converter fails (the caller surfaces the error dialog).
 */
export async function convertLegacyDoc(
  docPath: string,
  steps: DocConversionSteps = defaultConversionSteps(docToText),
): Promise<LegacyDocOutcome> {
  const targetPath = docxTargetFor(docPath)
  if (steps.exists(targetPath)) return { via: 'existing' }
  if (steps.sofficePath !== null && (await steps.convertSoffice(docPath, targetPath))) {
    return { via: 'soffice' }
  }
  if (await steps.convertWord(docPath, targetPath)) return { via: 'word' }
  const text = await steps.docToText(docPath)
  // Paragraph breaks carry no content; an empty result still needs one body paragraph.
  const paragraphs = text.split('\n').map((line) => line.trimEnd())
  await steps.writeTextDocx(targetPath, paragraphs.length > 0 ? paragraphs : [''])
  return { via: 'text' }
}
