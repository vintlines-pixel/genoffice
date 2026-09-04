import { useState } from 'react'

import type { WorkbookExportPdfRequest } from '../shared/desktop-api'
import { useI18n } from './i18n/locale'
import { parsePageRanges } from './page-ranges'

/// Excel's Export/P print dialog, reduced to what the export pipeline
/// supports: a page range over the paginated output (Chromium pageRanges)
/// and a preview button that shows the real pagination in the built-in PDF
/// viewer window before anything is saved. Preview and export share the
/// exact same payload, so what you preview is what you get.

export function ExportPdfDialog({
  draft,
  onPreview,
  onExport,
  onClose,
}: {
  /// The active sheet's print payload (already laid out with the effective
  /// page setup).
  readonly draft: WorkbookExportPdfRequest
  readonly onPreview: (payload: WorkbookExportPdfRequest) => void
  readonly onExport: (payload: WorkbookExportPdfRequest) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [custom, setCustom] = useState(false)
  const [rangeText, setRangeText] = useState('')
  const [error, setError] = useState<string | null>(null)

  const adjusted = (): WorkbookExportPdfRequest | null => {
    if (!custom) return draft
    const ranges = parsePageRanges(rangeText)
    if (ranges === null) {
      setError(t('dlgExportPdfInvalidRange'))
      return null
    }
    return { ...draft, pageRanges: ranges }
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('dlgExportPdfTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgExportPdfTitle')}</header>
        <div className="dialog-grid" style={{ gridTemplateColumns: '1fr' }}>
          <span className="dialog-span" style={{ display: 'flex', gap: '12px' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
              <input
                type="radio"
                checked={!custom}
                onChange={() => {
                  setCustom(false)
                  setError(null)
                }}
              />
              {t('dlgExportPdfRangeAll')}
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
              <input
                type="radio"
                checked={custom}
                onChange={() => {
                  setCustom(true)
                  setError(null)
                }}
              />
              {t('dlgExportPdfRangeCustom')}
            </label>
            <input
              style={{ flex: 1 }}
              value={custom ? rangeText : ''}
              disabled={!custom}
              placeholder={t('dlgExportPdfRangePlaceholder')}
              onChange={(event) => {
                setRangeText(event.target.value)
                setError(null)
              }}
            />
          </span>
          <p className="dialog-note dialog-span">{t('dlgExportPdfNote')}</p>
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
            className="secondary"
            onClick={() => {
              const payload = adjusted()
              if (payload !== null) onPreview(payload)
            }}
          >
            {t('dlgExportPdfPreview')}
          </button>
          <button
            className="primary-action"
            onClick={() => {
              const payload = adjusted()
              if (payload !== null) onExport(payload)
            }}
          >
            {t('dlgExportPdfExport')}
          </button>
        </div>
      </div>
    </div>
  )
}
