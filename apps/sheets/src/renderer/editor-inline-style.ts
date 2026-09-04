/**
 * Excel parity for styling INSIDE the in-cell editor: selecting characters
 * and hitting Bold / font color must restyle only those characters (rich
 * text runs, `cell.p`), not the whole cell.
 *
 * Routes ribbon style commands to Univer's `sheet.command.set-range-*`
 * wrappers — the same commands its own Ctrl+B/I/U shortcuts use. Their
 * handler checks the EDITOR_ACTIVATED context: while the in-cell editor is
 * open they apply the style to the editor's text selection as rich runs;
 * otherwise they style the cell. Commands without a rich-run equivalent
 * (double underline, fill, borders, number format…) return null so the
 * caller keeps the whole-cell path.
 */
import type { UniverRuntime } from './univer-state'

/// True while Univer's in-cell editor is open (the workbook exposes
/// isCellEditing on its facade object).
export function isCellEditorOpen(runtime: UniverRuntime): boolean {
  const workbook = runtime.univerAPI.getActiveWorkbook() as
    | { isCellEditing?(): boolean }
    | null
    | undefined
  return workbook?.isCellEditing?.() === true
}

export interface InlineStyleCommand {
  readonly id: string
  readonly params: Record<string, unknown>
}

/// Maps a `name:argument` ribbon style command (see parseStyleCommand) to
/// Univer's range-style wrapper command; null = no rich-run equivalent.
export function inlineStyleEditorCommand(
  name: string,
  argument: string,
): InlineStyleCommand | null {
  switch (name) {
    case 'bold':
      return { id: 'sheet.command.set-range-bold', params: {} }
    case 'italic':
      return { id: 'sheet.command.set-range-italic', params: {} }
    case 'underline':
      // Rich runs carry a single underline; the double variant stays a
      // whole-cell format.
      return argument === 'double' ? null : { id: 'sheet.command.set-range-underline', params: {} }
    case 'strike':
      return { id: 'sheet.command.set-range-stroke', params: {} }
    case 'font-size': {
      const size = Number(argument)
      return Number.isFinite(size) && size > 0
        ? { id: 'sheet.command.set-range-fontsize', params: { value: size } }
        : null
    }
    case 'font-family':
      return argument
        ? { id: 'sheet.command.set-range-font-family', params: { value: argument } }
        : null
    case 'font-color':
      // 'auto' is Excel's automatic (black) color; otherwise a #rrggbb hex.
      return {
        id: 'sheet.command.set-range-text-color',
        params: { value: argument === 'auto' || argument === '' ? '#000000' : argument },
      }
    default:
      return null
  }
}
