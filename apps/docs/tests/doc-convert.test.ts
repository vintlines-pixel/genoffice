import { describe, expect, it } from 'vitest'

import {
  convertDocWithSoffice,
  convertDocWithWord,
  convertLegacyDoc,
  docxTargetFor,
  findSoffice,
  isLegacyDocPath,
  type DocConversionSteps,
} from '../src/main/doc-convert'

describe('path helpers', () => {
  it('recognizes .doc but not .docx', () => {
    expect(isLegacyDocPath('report.doc')).toBe(true)
    expect(isLegacyDocPath('report.docx')).toBe(false)
    expect(isLegacyDocPath('a.DOC')).toBe(true)
  })

  it('replaces the trailing .doc with .docx in the same folder', () => {
    expect(docxTargetFor('C:\\docs\\report.doc')).toBe('C:\\docs\\report.docx')
    expect(docxTargetFor('report.DOC')).toBe('report.docx')
  })
})

describe('findSoffice', () => {
  const exists = () => true

  it('honors the SOFFICE_PATH override', () => {
    const env = { SOFFICE_PATH: 'C:\\tools\\soffice.exe' }
    expect(findSoffice(env, 'win32', exists)).toBe('C:\\tools\\soffice.exe')
  })

  it('resolves the standard Windows install paths', () => {
    const env: NodeJS.ProcessEnv = {}
    expect(findSoffice(env, 'win32', (p) => p.includes('Program Files\\LibreOffice'))).toBe(
      'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
    )
  })

  it('resolves the macOS app bundle', () => {
    const env: NodeJS.ProcessEnv = {}
    expect(findSoffice(env, 'darwin', exists)).toBe(
      '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    )
  })

  it('falls back to PATH `soffice` on linux; null when nothing exists', () => {
    const env: NodeJS.ProcessEnv = {}
    expect(findSoffice(env, 'linux', () => false)).toBe('soffice')
    expect(findSoffice(env, 'win32', () => false)).toBeNull()
    expect(findSoffice(env, 'darwin', () => false)).toBeNull()
  })
})

describe('convertDocWithSoffice', () => {
  it('spawns soffice --convert-to docx into the source dir', async () => {
    const seen: { program: string; args: string[] } = { program: '', args: [] }
    const ok = await convertDocWithSoffice(
      'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
      'C:\\docs\\report.doc',
      'C:\\docs\\report.docx',
      async (program, args) => {
        seen.program = program
        seen.args = [...args]
        return { code: 0, timedOut: false }
      },
    )
    expect(seen.program).toContain('soffice.exe')
    expect(seen.args).toEqual(
      expect.arrayContaining(['--headless', '--convert-to', 'docx', '--outdir', 'C:\\docs', 'C:\\docs\\report.doc']),
    )
    // target must exist for it to count
    expect(ok).toBe(false)
  })

  it('reports failure on non-zero exit or timeout', async () => {
    const nonzero = await convertDocWithSoffice(
      'soffice',
      'a.doc',
      'a.docx',
      async (_p, _a) => ({ code: 1, timedOut: false }),
    )
    expect(nonzero).toBe(false)
  })
})

describe('convertDocWithWord', () => {
  it('drives Word COM via a SaveAs2(16) script', async () => {
    const seen: { program: string; args: string[] } = { program: '', args: [] }
    await convertDocWithWord('C:\\docs\\report.doc', 'C:\\docs\\report.docx', async (program, args) => {
      seen.program = program
      seen.args = [...args]
      return { code: 0, timedOut: false }
    })
    expect(seen.program).toBe('powershell.exe')
    expect(seen.args.join(' ')).toContain("SaveAs2('C:\\docs\\report.docx', 16)")
  })
})

describe('convertLegacyDoc', () => {
  const steps = (overrides: Partial<DocConversionSteps>): DocConversionSteps => ({
    sofficePath: null,
    convertSoffice: async () => false,
    convertWord: async () => false,
    docToText: async () => 'line1\nline2\n\nline4',
    writeTextDocx: async () => undefined,
    exists: () => false,
    ...overrides,
  })

  it('opens an existing docx target as-is', async () => {
    const outcome = await convertLegacyDoc('report.doc', steps({ exists: () => true }))
    expect(outcome.via).toBe('existing')
  })

  it('prefers LibreOffice, then Word, then text import', async () => {
    const trace: string[] = []
    const soffice = await convertLegacyDoc(
      'report.doc',
      steps({
        sofficePath: 'soffice',
        convertSoffice: async () => {
          trace.push('soffice')
          return true
        },
      }),
    )
    expect(soffice.via).toBe('soffice')
    expect(trace).toEqual(['soffice'])

    const word = await convertLegacyDoc(
      'report.doc',
      steps({
        sofficePath: 'soffice',
        convertSoffice: async () => false,
        convertWord: async () => {
          trace.push('word')
          return true
        },
      }),
    )
    expect(word.via).toBe('word')
    expect(trace).toEqual(['soffice', 'word'])
  })

  it('text-imports as paragraphs (split on breaks, right-trimmed)', async () => {
    const wrote: { target: string; paragraphs: readonly string[] }[] = []
    const outcome = await convertLegacyDoc(
      'report.doc',
      steps({
        writeTextDocx: async (target, paragraphs) => {
          wrote.push({ target, paragraphs })
        },
      }),
    )
    expect(outcome.via).toBe('text')
    expect(wrote[0]?.target).toBe('report.docx')
    expect(wrote[0]?.paragraphs).toEqual(['line1', 'line2', '', 'line4'])
  })

  it('keeps a single empty body paragraph when no text survives', async () => {
    const wrote: { paragraphs: readonly string[] }[] = []
    await convertLegacyDoc(
      'report.doc',
      steps({
        docToText: async () => '',
        writeTextDocx: async (_t, paragraphs) => {
          wrote.push({ paragraphs })
        },
      }),
    )
    expect(wrote[0]?.paragraphs).toEqual([''])
  })
})
