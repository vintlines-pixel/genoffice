/**
 * Auto-answer the app's native close-guard dialogs (Save / Don't Save / Cancel)
 * with "Don't Save" so driver runs close cleanly without a blocking dialog.
 *
 * Drivers that leave the document dirty would otherwise strand a modal on
 * quit. Install right after launch: the stub replaces dialog.showMessageBox(±Sync)
 * in the main process and picks the Don't Save / Close Anyway button whenever
 * one is offered; other dialogs fall through to their first button.
 */
export async function autoDontSaveOnClose(app) {
  await app.evaluate(({ dialog }) => {
    const pickResponse = (opts) => {
      const buttons = opts?.buttons ?? []
      const dontSave = buttons.findIndex((b) =>
        /don.?t save|不保存|不儲存|저장 안 함/i.test(String(b)),
      )
      if (dontSave >= 0) return dontSave
      const closeAnyway = buttons.findIndex((b) =>
        /close anyway|discard|放弃更改|放棄變更|变更の破棄/i.test(String(b)),
      )
      return closeAnyway >= 0 ? closeAnyway : 0
    }
    const reply = (a, b) => {
      const opts = b?.buttons ? b : a
      return { response: pickResponse(opts), checkboxChecked: false }
    }
    dialog.showMessageBox = async (a, b) => reply(a, b)
    dialog.showMessageBoxSync = (a, b) => reply(a, b).response
  })
}
