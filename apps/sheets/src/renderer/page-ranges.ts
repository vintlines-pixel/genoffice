/// Export-PDF page-range parsing: the dialog's free-text input ('1-3, 5')
/// → Chromium's pageRanges string ('1-3,5'), or null when malformed.

/// Validates and normalizes one page-range input; null when malformed.
/// Segments are single pages ('5') or spans ('2-7', start ≤ end, 1-based).
export function parsePageRanges(input: string): string | null {
  const text = input.trim()
  if (text === '') return null
  const parts: string[] = []
  for (const segment of text.split(',')) {
    const match = /^(\d{1,7})(?:\s*-\s*(\d{1,7}))?$/.exec(segment.trim())
    if (match === null) return null
    const start = Number(match[1])
    if (start === 0) return null
    const endText = match[2]
    if (endText === undefined) {
      parts.push(String(start))
      continue
    }
    const end = Number(endText)
    if (end === 0 || end < start) return null
    parts.push(end === start ? String(start) : `${start}-${end}`)
  }
  return parts.join(',')
}
