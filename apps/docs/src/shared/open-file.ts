/** .docx opens directly; a legacy .doc is converted next to itself first. */
export function findDocxPath(argv: readonly string[]): string | null {
  return (
    argv.find((arg) => {
      const value = arg.trim()
      return value.length > 0 && !value.startsWith('-') && /\.docx?$/i.test(value)
    }) ?? null
  )
}
