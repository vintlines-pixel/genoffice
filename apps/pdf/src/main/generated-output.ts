import { existsSync } from 'node:fs'
import { join } from 'node:path'

/** Pick a safe, unused PDF path inside the configured GenOffice save directory. */
export function uniqueGeneratedPdfPath(
  dir: string,
  suggestedName: string,
  pathExists: (path: string) => boolean = existsSync,
): string {
  // Control characters and Windows-invalid filename characters are rejected
  // from generated file names. Split on both path separators first — basename()
  // would treat a leading "x:" as a Windows drive and silently drop it.
  // eslint-disable-next-line no-control-regex
  const lastSegment = String(suggestedName || 'merged.pdf').split(/[/\\]/).pop() ?? ''
  let fileName = lastSegment.replace(/[:*?"<>|\u0000-\u001f]/g, '_').trim()
  if (!fileName || fileName === '.' || fileName === '..') fileName = 'merged.pdf'
  if (!/\.pdf$/i.test(fileName)) fileName += '.pdf'

  const stem = fileName.slice(0, -4)
  let candidate = join(dir, fileName)
  for (let i = 2; pathExists(candidate); i++) candidate = join(dir, `${stem}-${i}.pdf`)
  return candidate
}
