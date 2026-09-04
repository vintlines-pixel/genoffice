import { useState } from 'react'

import type { CustomMargins } from './edit-journal'
import { useI18n, type StringKey } from './i18n/locale'
import { NORMAL_MARGINS, type PrintMargins } from './print-settings'

/// Excel's Page Layout → Margins → Custom Margins as a dialog: the six inch
/// values (left/right/top/bottom plus header/footer offsets from the paper's
/// top/bottom edge). OK journals them as page-setup state; the save writes
/// the worksheet's <pageMargins> element.

type EdgeKey =
  | 'left'
  | 'right'
  | 'top'
  | 'bottom'
  | 'header'
  | 'footer'

const EDGE_KEYS: readonly EdgeKey[] = ['left', 'right', 'top', 'bottom', 'header', 'footer']

const EDGE_LABELS: Record<EdgeKey, StringKey> = {
  left: 'dlgMarginsLeft',
  right: 'dlgMarginsRight',
  top: 'dlgMarginsTop',
  bottom: 'dlgMarginsBottom',
  header: 'dlgMarginsHeader',
  footer: 'dlgMarginsFooter',
}

/// Parses one edge value; anything outside 0–10 (Excel's dialog range,
/// mirroring the save schema) is rejected. Empty input is rejected too —
/// Number('') would otherwise silently become 0.
function parseEdge(value: string): number | null {
  const trimmed = value.trim()
  if (trimmed === '') return null
  if (!/^\d{0,2}(?:\.\d{1,3})?$/.test(trimmed)) return null
  const parsed = Number(trimmed)
  if (!Number.isFinite(parsed) || parsed < 0 || parsed > 10) return null
  return parsed
}

export function MarginsDialog({
  initial,
  onApply,
  onClose,
}: {
  /// Concrete inch values to prefill (journal + file resolved); null falls
  /// back to the normal preset.
  readonly initial: PrintMargins | null
  /// Returns an error message, or null on success.
  readonly onApply: (margins: CustomMargins) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const source = initial ?? NORMAL_MARGINS
  const [values, setValues] = useState<Record<EdgeKey, string>>({
    left: String(source.left),
    right: String(source.right),
    top: String(source.top),
    bottom: String(source.bottom),
    header: String(source.header),
    footer: String(source.footer),
  })
  const [error, setError] = useState<string | null>(null)

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('dlgMarginsTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgMarginsTitle')}</header>
        <div className="dialog-grid" style={{ gridTemplateColumns: '1fr 1fr 1fr' }}>
          {EDGE_KEYS.map((key) => (
            <label key={key}>
              {t(EDGE_LABELS[key])}
              <input
                value={values[key]}
                inputMode="decimal"
                onChange={(event) => {
                  setValues((prev) => ({ ...prev, [key]: event.target.value }))
                  setError(null)
                }}
              />
            </label>
          ))}
          <p className="dialog-note dialog-span">{t('dlgMarginsNote')}</p>
        </div>
        {error && (
          <p className="dialog-note" role="alert">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="primary-action"
            onClick={() => {
              const edges = Object.fromEntries(
                EDGE_KEYS.map((key) => [key, parseEdge(values[key])]),
              ) as Record<EdgeKey, number | null>
              const invalid = EDGE_KEYS.some((key) => edges[key] === null)
              if (invalid) {
                setError(t('dlgMarginsInvalid'))
                return
              }
              const failure = onApply({
                left: edges.left!,
                right: edges.right!,
                top: edges.top!,
                bottom: edges.bottom!,
                header: edges.header!,
                footer: edges.footer!,
              })
              setError(failure)
              if (failure === null) onClose()
            }}
          >
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
