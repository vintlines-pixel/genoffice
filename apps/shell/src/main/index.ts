import { execSync, spawn } from 'node:child_process'
import {
  copyFileSync,
  cpSync,
  existsSync,
  readFileSync,
  readdirSync,
  renameSync,
  writeFileSync,
} from 'node:fs'
import { basename, dirname, extname, join } from 'node:path'
import {
  BrowserWindow,
  Menu,
  app,
  dialog,
  ipcMain,
  nativeImage,
  nativeTheme,
  session,
  shell,
  webContents,
} from 'electron'
import type { MenuItemConstructorOptions, NativeImage, WebContents } from 'electron'
import menuDocxIcon1x from './assets/menu-docx.png?asset'
import menuDocxIcon2x from './assets/menu-docx@2x.png?asset'
import menuXlsxIcon1x from './assets/menu-xlsx.png?asset'
import menuXlsxIcon2x from './assets/menu-xlsx@2x.png?asset'
import menuPptxIcon1x from './assets/menu-pptx.png?asset'
import menuPptxIcon2x from './assets/menu-pptx@2x.png?asset'
import menuPdfIcon1x from './assets/menu-pdf.png?asset'
import menuPdfIcon2x from './assets/menu-pdf@2x.png?asset'
import menuMdIcon1x from './assets/menu-md.png?asset'
import menuMdIcon2x from './assets/menu-md@2x.png?asset'
import menuHomeIcon1x from './assets/menu-home.png?asset'
import menuHomeIcon2x from './assets/menu-home@2x.png?asset'
import { isLang, normalizeLang, setUiLang, type Lang } from '@genoffice/i18n'
import { tMain, type ShellI18nKey } from './shell-strings'
import {
  DEFAULT_SAVE_DIR_KEY,
  DROP_OPEN_CHANNEL,
  GITHUB_REPO_URL,
  appMenuLabels,
  contextMenuLabels,
  editMenuTemplate,
  installContextMenu,
  installNavigationGuard,
  isUsableSaveDir,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
  windowMenuTemplate,
} from '@genoffice/electron-utils'
import { readAppSettings, writeAppSetting, writeAppSettings } from './app-settings'
import {
  ANALYTICS_ENABLED_KEY,
  analyticsEnabledFrom,
  createAnalytics,
  ensureAnalyticsClientState,
  extractPackagedAnalyticsKeys,
  markAnalyticsFirstLaunchSent,
} from './analytics'
import type { Analytics, AnalyticsKeys } from './analytics'
import {
  LAST_RUN_VERSION_KEY,
  STAR_PROMPT_KEY,
  asStarPromptState,
  isUpgradeLaunch,
  shouldShowStarPrompt,
  shouldShowUpgradeStarPrompt,
  withDocOpen,
  withFirstRun,
  withResolved,
  withShown,
} from './star-prompt'
import {
  clearCloudProjectsStore,
  cloudProjectExternalUrl,
  readCloudProjectsStore,
  syncCloudProjects,
} from './cloud-projects'
import { handleDroppedFiles } from './dropped-files'
import { ProjectStore } from '@genoffice/project-store'
import {
  ensureGenofficeLogin,
  genofficeLogout,
  gskConvertPdfToDocx,
  gskLoginInfo,
  hasGskAuth,
  loadGenofficeAuth,
  resolveGskEntry,
  setGskProxyUrl,
  startGenofficeLogin,
} from '@genoffice/ai-search'

import {
  buildDocsMenu,
  configureDocsRuntime,
  docsFileRenamed,
  docsQueryDirty,
  requestDocsClose,
  readRecentFiles,
  readStarredFiles,
  recordRecentFile,
  removeRecentFiles,
  replaceRecentFile,
  registerAiIpc,
  registerProjectIpc,
  toggleStarredFile,
  registerDocsIpc,
  setDocsExtraFileMenuItems,
  setDocsMenuGate,
  setDocsShellHooks,
  createAiDocument,
  projectFileRenamed,
  setDocsShellWindow,
  setDocsFileSavedHook,
  setDocsFileOpenedHook,
  setSessionPathResolver,
  defaultSaveDir,
  uniquePathIn,
} from '../../../docs/src/main/docs-main'
import { blankXlsxBuffer } from '../../../sheets/src/gateway/csv-import'
import { blankPdfBuffer } from '../../../pdf/src/main/blank-pdf'
import {
  configureSheetsRuntime,
  hasActiveQueuedWorkbook,
  installSheetsMenu,
  markSheetsShuttingDown,
  requestSheetsClose,
  resolveSheetsSessionPath,
  markSheetsUntitledPath,
  sendSheetsMenuAction,
  sheetsFileRenamed,
  setSheetsCloseTabHook,
  setSheetsExtraFileMenuItems,
  setSheetsShellWindow,
  setSheetsWorkbookOpenedHook,
  startSheetsCaptureServer,
  stopSheetsSidecar,
} from '../../../sheets/src/main/sheets-main'
import {
  configureSlidesRuntime,
  installSlidesMenu,
  replaceSlidesRecentFile,
  requestSlidesClose,
  setSlidesCloseTabHook,
  setSlidesExtraFileMenuItems,
  setSlidesOpenedHook,
  setSlidesShellWindow,
  setSlidesShowBleed,
  slidesFileRenamed,
} from '../../../slides/src/main/slides-main'
import {
  configurePdfRuntime,
  flushPdfSave,
  markPdfUntitledPath,
  pdfIsDirty,
  requestPdfClose,
  requestPdfSaveAs,
  sendPdfPrintRequest,
  setPdfRenamedHook,
  setPdfSaveAsInFlight,
} from '../../../pdf/src/main/pdf-main'
import { PDF_CHANNELS } from '../../../pdf/src/shared/ipc'
import { convertPdfFileToDocxLocalWithPrompt, PdfLoadError } from './pdf2docx-local'
import { convertPdfFileToPptxLocalWithPrompt } from './pdf2pptx-local'
import { convertPdfFileToXlsxLocalWithPrompt } from './pdf2xlsx-local'
import { closePdfPasswordDialog, promptPdfPassword } from './pdf-password-dialog'
import {
  configureMarkdownRuntime,
  markdownFileRenamed,
  requestMarkdownClose,
  requestMarkdownSave,
  sendMarkdownExportRequest,
  sendMarkdownPrintRequest,
  setMarkdownDocxExportedHook,
  setMarkdownFileSavedHook,
} from '../../../markdown/src/main/markdown-main'
import type {
  AccountLoginEvent,
  RecentEntry,
  RecentPage,
  RenameResult,
  StarPromptShow,
  UiTheme,
} from '../shared/home-api'
import { HOME_CHANNELS } from '../shared/home-api'
import type { TabKind } from '../shared/tabs-api'
import { TABS_CHANNELS } from '../shared/tabs-api'
import { showErrorDialog } from './error-dialog'
import { normalizeRecentQuery, pageRecentPaths, statExistingPaths } from './recent-files'
import { TabManager } from './tab-manager'
import { getProxyBootstrap, startMainProcessProxy } from './main-process-proxy'
import { applyUpdateChannel, initAutoUpdater } from './updater'
import { isUpdateChannel, type UpdateChannel } from '../shared/update-api'

/**
 * GenOffice unified shell: ONE Electron app, ONE BrowserWindow, hosting the
 * docs and sheets modules as WebContentsView tabs behind a WPS-style tab
 * strip. The shell owns the lifecycle — single-instance lock, file-
 * association routing by extension, and per-active-tab menu switching.
 * Renderers load from each module's build output (apps/docs/out,
 * apps/sheets/out), so build those before running the shell.
 */

// ANY unpacked run (`npm run shell`, `npm run dev`, `npx electron .`) must not
// share the installed app's userData or single-instance lock — otherwise a dev
// run silently quits and forwards its argv to the running installed GenOffice.
// GENOFFICE_USER_DATA: test drivers point this at a scratch dir so an
// automated instance can run alongside the dev instance (separate lock).
if (!app.isPackaged)
  app.setPath(
    'userData',
    process.env.GENOFFICE_USER_DATA ?? join(app.getPath('appData'), 'GenOffice Dev'),
  )

// The product rename from "AI Office" to GenOffice changed the userData path; migrate old user data once
if (app.isPackaged) {
  const oldDir = join(app.getPath('appData'), 'AI Office')
  const newDir = app.getPath('userData')
  const newEmpty = !existsSync(newDir) || readdirSync(newDir).length === 0
  if (newEmpty && existsSync(oldDir)) cpSync(oldDir, newDir, { recursive: true })
}

// module build outputs: packaged builds carry them as extraResources
// (resources/modules/*, resources/native/*); dev/unpacked resolves them
// relative to apps/shell in the monorepo layout.
const SIDECAR_EXE = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
const APPS_ROOT = join(app.getAppPath(), '..')
const DOCS_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'docs')
  : join(APPS_ROOT, 'docs', 'out')
const SHEETS_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'sheets')
  : join(APPS_ROOT, 'sheets', 'out')
const SLIDES_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'slides')
  : join(APPS_ROOT, 'slides', 'out')
const PDF_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'pdf')
  : join(APPS_ROOT, 'pdf', 'out')
const MARKDOWN_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'markdown')
  : join(APPS_ROOT, 'markdown', 'out')
const SIDECAR_BIN = app.isPackaged
  ? join(process.resourcesPath, 'native', SIDECAR_EXE)
  : join(APPS_ROOT, 'sheets', 'native', 'xlsx-engine', 'target', 'release', SIDECAR_EXE)

configureDocsRuntime({
  preloadPath: join(DOCS_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.DOCS_RENDERER_URL,
  rendererFile: join(DOCS_OUT, 'renderer', 'index.html'),
})
configureSheetsRuntime({
  preloadPath: join(SHEETS_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.SHEETS_RENDERER_URL,
  rendererFile: join(SHEETS_OUT, 'renderer', 'index.html'),
  sidecarPath: SIDECAR_BIN,
  openGeneratedPath: (path) => openGeneratedDocument(path),
})
configureSlidesRuntime({
  preloadPath: join(SLIDES_OUT, 'preload', 'index.js'),
  rendererDevUrl: process.env.SLIDES_RENDERER_URL,
  rendererFilePath: join(SLIDES_OUT, 'renderer', 'index.html'),
  openGeneratedPath: (path) => openGeneratedDocument(path),
})
configurePdfRuntime({
  preloadPath: join(PDF_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.PDF_RENDERER_URL,
  rendererFile: join(PDF_OUT, 'renderer', 'index.html'),
  openGeneratedPath: (path) => openGeneratedDocument(path),
  createDocument: createAiDocument,
})
configureMarkdownRuntime({
  preloadPath: join(MARKDOWN_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.MARKDOWN_RENDERER_URL,
  rendererFile: join(MARKDOWN_OUT, 'renderer', 'index.html'),
  openGeneratedPath: (path) => openGeneratedDocument(path),
})

// ---- UI language ----
// Persisted in userData/app-settings.json so the editor modules can read the
// same file when they pick up i18n later. GENOFFICE_LANG overrides for tests.

const APP_SETTINGS_PATH = () => join(app.getPath('userData'), 'app-settings.json')

let uiLang: Lang | null = null

function currentLang(): Lang {
  if (uiLang) return uiLang
  if (process.env.GENOFFICE_LANG) {
    uiLang = normalizeLang(process.env.GENOFFICE_LANG)
    setUiLang(uiLang)
    return uiLang
  }
  const saved = readAppSettings(APP_SETTINGS_PATH()).language
  if (isLang(saved)) uiLang = saved
  uiLang ??= normalizeLang(app.getLocale())
  setUiLang(uiLang)
  return uiLang
}

function persistLang(lang: Lang): void {
  uiLang = lang
  setUiLang(lang)
  writeAppSetting(APP_SETTINGS_PATH(), 'language', lang)
}

let cachedUpdateChannel: UpdateChannel | null = null

function currentUpdateChannel(): UpdateChannel {
  if (cachedUpdateChannel) return cachedUpdateChannel
  const saved = readAppSettings(APP_SETTINGS_PATH()).updateChannel
  cachedUpdateChannel = isUpdateChannel(saved) ? saved : 'stable'
  return cachedUpdateChannel
}

let cachedTheme: UiTheme | null = null

function currentTheme(): UiTheme {
  if (cachedTheme) return cachedTheme
  const saved = readAppSettings(APP_SETTINGS_PATH()).theme
  cachedTheme = saved === 'light' || saved === 'dark' ? saved : 'system'
  return cachedTheme
}

// ---- anonymous usage analytics (see src/main/analytics.ts) ----
// Stays a no-op until initAnalytics() runs at startup; keyless builds
// (source/forks) keep the no-op forever, so every track() call is safe.

let analytics: Analytics = { active: false, track: () => {} }

let cachedAnalyticsEnabled: boolean | null = null

function analyticsEnabled(): boolean {
  cachedAnalyticsEnabled ??= analyticsEnabledFrom(readAppSettings(APP_SETTINGS_PATH()))
  return cachedAnalyticsEnabled
}

function resolveAnalyticsKeys(): AnalyticsKeys | null {
  // Only packaged extraMetadata is authoritative. Source/dev runs never read
  // runtime credentials and therefore remain a strict no-op.
  if (!app.isPackaged) return null
  try {
    return extractPackagedAnalyticsKeys(
      JSON.parse(readFileSync(join(app.getAppPath(), 'package.json'), 'utf8')),
      app.isPackaged,
    )
  } catch {
    return null
  }
}

function persistAnalyticsPreference(enabled: boolean): boolean {
  const previous = cachedAnalyticsEnabled
  // Change the in-memory gate before touching disk. The synchronous atomic
  // write prevents another event from being handled in between.
  cachedAnalyticsEnabled = enabled
  try {
    writeAppSettings(APP_SETTINGS_PATH(), { [ANALYTICS_ENABLED_KEY]: enabled })
    return true
  } catch (error) {
    cachedAnalyticsEnabled = previous
    throw error
  }
}

function initAnalytics(): void {
  try {
    let clientState: ReturnType<typeof ensureAnalyticsClientState> | null = null
    const getClientState = () => (clientState ??= ensureAnalyticsClientState(APP_SETTINGS_PATH()))
    analytics = createAnalytics({
      keys: resolveAnalyticsKeys(),
      getClientId: () => getClientState().clientId,
      isEnabled: analyticsEnabled,
      shouldTrackFirstLaunch: () => getClientState().firstLaunchPending,
      onFirstLaunchSent: () => markAnalyticsFirstLaunchSent(APP_SETTINGS_PATH()),
      // Country-only approximation from OS regional settings. This avoids an
      // IP lookup while populating GA4's built-in Country dimension.
      getCountryCode: () => app.getLocaleCountryCode(),
      // evaluated per event: ui_lang follows live language switches
      baseParams: () => ({
        app_version: app.getVersion(),
        platform: process.platform,
        os_version: process.getSystemVersion(),
        ui_lang: currentLang(),
      }),
    })
  } catch {
    // analytics must never block startup
  }
}

// ---- first-run onboarding ----
// The GenTeam community page opened from the onboarding's second slide.
// Stable short link served by the genoffice.ai site; it 302s to the tokened
// invite link, which stays out of this repo and rotates server-side.
const GENTEAM_URL = 'https://genoffice.ai/join'

// Genspark credit-usage page opened from the account menu's credits row.
// Kept main-side so the renderer never supplies the URL.
const CREDIT_USAGE_URL = 'https://www.genspark.ai/credit-usage'

// ---- "star us on GitHub" prompt (see star-prompt.ts for the rules) ----

const readStarPrompt = () =>
  asStarPromptState(readAppSettings(APP_SETTINGS_PATH())[STAR_PROMPT_KEY])
const writeStarPrompt = (state: ReturnType<typeof readStarPrompt>) =>
  writeAppSetting(APP_SETTINGS_PATH(), STAR_PROMPT_KEY, state)

/** set at startup when this is the first launch after an upgrade; consumed by
 * the first starPromptShouldShow query of the session */
let upgradeStarPromptPending = false

/** a granted show, cached for the session: repeated queries (React StrictMode
 * double-effects, AppFrame remounts) must return the same answer instead of
 * burning another lifetime show or flipping to a snoozed "false" */
let starPromptSessionGrant: StarPromptShow | null = null

/** every successful document open counts toward the prompt's value threshold */
function recordStarPromptDocOpen(): void {
  try {
    const state = readStarPrompt()
    const next = withDocOpen(state)
    if (next !== state) writeStarPrompt(next)
  } catch {
    // settings write failures must never break opening a document
  }
}

// Stargazer count for the settings About pane; fetched main-side (the
// renderer CSP has no api.github.com) and cached per session — the exact
// number is decoration, staleness is fine.
let cachedGithubStars: number | null = null

async function fetchGithubStars(): Promise<number | null> {
  if (cachedGithubStars !== null) return cachedGithubStars
  try {
    const response = await fetch('https://api.github.com/repos/genspark-ai/genoffice', {
      headers: { Accept: 'application/vnd.github+json' },
      signal: AbortSignal.timeout(5000),
    })
    if (!response.ok) return null
    const body: unknown = await response.json()
    const count = (body as { stargazers_count?: unknown }).stargazers_count
    if (typeof count !== 'number' || !Number.isFinite(count)) return null
    cachedGithubStars = count
    return count
  } catch {
    return null
  }
}



const tm = (key: Parameters<typeof tMain>[1], params?: Parameters<typeof tMain>[2]) =>
  tMain(currentLang(), key, params)

// ---- the shell window + its tab manager (recreated if the user closes it on macOS) ----

let shellWindow: BrowserWindow | null = null
let tabManager: TabManager | null = null

/**
 * When the user creates a file from a specific project view, remember which
 * project the next save should belong to. key: 'doc' | 'sheet' | 'slide', value: projectId.
 * Consumed by each app's saveHook once the file first hits disk (P1 item 3).
 */
const pendingNewFileProject = new Map<string, string>()

/**
 * P1: after a file first hits disk, if a pending project was set earlier via
 * "create from project view", move the new file into that project automatically.
 * Called from createShellWindow's opened/saved hooks.
 */
function applyPendingProject(filePath: string): void {
  const ext = extname(filePath).slice(1).toLowerCase()
  let key: string | undefined
  if (ext === 'docx') key = 'doc'
  else if (ext === 'xlsx' || ext === 'xlsm' || ext === 'xls' || ext === 'csv') key = 'sheet'
  else if (ext === 'pptx') key = 'slide'
  else if (ext === 'md' || ext === 'markdown') key = 'markdown'
  else if (ext === 'pdf') key = 'pdf'
  if (!key) return
  const projectId = pendingNewFileProject.get(key)
  if (!projectId) return
  pendingNewFileProject.delete(key)
  try {
    const store = new ProjectStore(app.getPath('userData'))
    store.ensureDefaultProject()
    store.resolveProjectForFile(filePath) // assign to default first (idempotent)
    store.moveFileToProject(filePath, projectId)
  } catch (err) {
    console.warn('[shell] applyPendingProject failed:', err)
  }
}

function applyMenuFor(kind: TabKind): void {
  switch (kind) {
    case 'docs':
      buildDocsMenu()
      break
    case 'sheets':
      installSheetsMenu()
      break
    case 'slides':
      installSlidesMenu()
      break
    case 'pdf':
      buildPdfMenu()
      break
    case 'markdown':
      buildMarkdownMenu()
      break
    default:
      buildHomeMenu()
  }
}

function createShellWindow(): void {
  const win = new BrowserWindow({
    width: 1360,
    height: 900,
    minWidth: 980,
    minHeight: 600,
    title: 'GenOffice',
    // vibrancy: editor modules punch translucent regions (e.g. the slides
    // thumbnail pane) through to the desktop
    ...(process.platform === 'darwin'
      ? { titleBarStyle: 'hiddenInset' as const, vibrancy: 'sidebar' as const }
      : {}),
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  shellWindow = win
  // dragging the window by the tab strip's blank (draggable) area produces no
  // DOM event anywhere — will-move is the only signal to dismiss popovers
  win.on('will-move', () => broadcastChromePressed())
  // A detached editor window claims the process-global menu/active-editor targets
  // while focused; take them back when the shell window regains focus
  win.on('focus', () => tabManager?.refreshActiveTargets())

  const manager = new TabManager(
    win,
    () => win.webContents.send(TABS_CHANNELS.changed, manager.list()),
    applyMenuFor,
    // no extension: these tabs have no file on disk yet; the title becomes the
    // real filename (the localized untitled default + .docx etc.) once the first save lands
    (kind) =>
      kind === 'docs'
        ? tm('untitledDoc')
        : kind === 'slides'
          ? tm('untitledDeck')
          : kind === 'markdown'
            ? tm('untitledMarkdown')
            : tm('untitledSheet'),
  )
  tabManager = manager

  // pushRecent-triggered docs menu rebuilds must not clobber the active tab's menu
  setDocsMenuGate(() => manager.list().some((t) => t.active && t.kind === 'docs'))

  setDocsShellWindow(win)
  setSheetsShellWindow(win)
  setSlidesShellWindow(win)
  setSlidesShowBleed((wc, on) => manager.setContentBleed(wc, on))
  setDocsShellHooks({
    openTab: (openPath, options) => manager.openDocsTab(openPath, options),
    openAiDocTab: (content) =>
      manager.openDocsTab(undefined, { newBlank: true, aiContent: content }),
    listTabs: () =>
      manager
        .list()
        .filter((t) => t.kind === 'docs')
        .map((t) => ({ id: t.id, title: t.title, focused: t.active })),
    focusTab: (id) => manager.activateTab(id),
    closeActiveTab: () => manager.closeActiveTab(),
    openGeneratedPath: (path) => openGeneratedDocument(path),
  })
  setSheetsCloseTabHook(() => manager.closeActiveTab())
  // ⌘W targets the focused window: in a detached slides editor window it closes
  // that window (running its own close guard), not the shell's active tab
  setSlidesCloseTabHook(() => {
    const focused = BrowserWindow.getFocusedWindow()
    if (focused && focused !== win) focused.close()
    else manager.closeActiveTab()
  })
  // When ⌘O opens a file inside a tab, sync the tab title/path (used for de-dup by path) and record it as recent.
  // The first save / save-as fires this too, so applyPendingProject also runs here.
  setSheetsWorkbookOpenedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  setSlidesOpenedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  // docs' save-as / silent first save lands on a new path → sync the tab title too
  setDocsFileSavedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  // ⌘O / open-path inside a docs tab: sync the tab title immediately, same
  // contract as the sheets/slides opened hooks (a plain save to the original
  // path never renames the tab, so the open must — r115)
  setDocsFileOpenedHook((wcId, path) => {
    manager.setTabFileFor(wcId, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  // markdown untitled first save / Save As lands on a new path
  setMarkdownFileSavedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  // pdf content-derived auto-rename: the file moved on disk, follow it everywhere
  setPdfRenamedHook((wc, oldPath, newPath) => {
    manager.setTabFileFor(wc.id, newPath)
    replaceRecentFile(oldPath, newPath)
    projectFileRenamed(oldPath, newPath)
  })
  // markdown "convert & open in Docs" → route the fresh .docx to a docs tab
  setMarkdownDocxExportedHook((path) => {
    openDocumentPath(path)
  })

  // Closing the whole window walks every dirty sheets/pdf/slides/docs tab through
  // the same save/don't-save/cancel prompt; any cancel aborts the close.
  // docs dirtiness lives renderer-side, so any live docs tab forces the async path
  // and gets queried there (clean tabs pass through without activation).
  let closeConfirmed = false
  win.on('close', (event) => {
    if (closeConfirmed) return
    const dirtySheets = manager.dirtySheetsTabs()
    const dirtyPdf = manager.dirtyPdfTabs()
    const dirtyMarkdown = manager.dirtyMarkdownTabs()
    const dirtySlides = manager.dirtySlidesTabs()
    const docsTabs = manager.docsTabs()
    if (
      dirtySheets.length === 0 &&
      dirtyPdf.length === 0 &&
      dirtyMarkdown.length === 0 &&
      dirtySlides.length === 0 &&
      docsTabs.length === 0
    )
      return
    event.preventDefault()
    void (async () => {
      for (const tab of dirtySheets) {
        manager.activateTab(tab.id)
        if (!(await requestSheetsClose(tab.webContents, win))) return
      }
      for (const tab of dirtyPdf) {
        manager.activateTab(tab.id)
        if (!(await requestPdfClose(tab.webContents, win))) return
      }
      for (const tab of dirtyMarkdown) {
        manager.activateTab(tab.id)
        if (!(await requestMarkdownClose(tab.webContents, win))) return
      }
      for (const tab of dirtySlides) {
        manager.activateTab(tab.id)
        if (!(await requestSlidesClose(tab.webContents, win))) return
      }
      for (const tab of docsTabs) {
        if (!(await docsQueryDirty(tab.webContents))) continue
        manager.activateTab(tab.id)
        if (!(await requestDocsClose(tab.webContents, win))) return
      }
      closeConfirmed = true
      if (!win.isDestroyed()) win.close()
    })()
  })

  win.on('closed', () => {
    if (shellWindow === win) shellWindow = null
    if (tabManager === manager) tabManager = null
  })

  if (process.env.ELECTRON_RENDERER_URL) {
    void win.loadURL(process.env.ELECTRON_RENDERER_URL)
  } else {
    void win.loadFile(join(__dirname, '../renderer/index.html'))
  }
}

// ---- routing: one dispatch function for every open path ----

const DOCX_RE = /\.docx$/i
const XLSX_RE = /\.(xlsx|xlsm|xls|csv)$/i
const PPTX_RE = /\.pptx$/i
const PDF_RE = /\.pdf$/i
const MD_RE = /\.(md|markdown)$/i

/** document formats we recognize but don't open — surfaced as a dialog, not silently dropped */
const UNSUPPORTED_DOC_RE = /\.(doc|rtf|odt|ppt|pps|odp|ods|xlsb|pages|key|numbers)$/i

/**
 * Single source of truth for the open-dialog filter. Includes the
 * legacy .doc/.ppt binaries so they are selectable and surface the explicit
 * "not supported" dialog via openDocumentPath instead of being grayed out.
 */
const OPEN_DIALOG_EXTENSIONS = [
  'docx',
  'doc',
  'xlsx',
  'xlsm',
  'xls',
  'csv',
  'pptx',
  'ppt',
  'pdf',
  'md',
  'markdown',
]

function supportedFileIn(argv: string[]): string | null {
  return (
    argv.find(
      (arg) =>
        (DOCX_RE.test(arg) ||
          XLSX_RE.test(arg) ||
          PPTX_RE.test(arg) ||
          PDF_RE.test(arg) ||
          MD_RE.test(arg)) &&
        existsSync(arg),
    ) ?? null
  )
}

function unsupportedFileIn(argv: string[]): string | null {
  return argv.find((arg) => UNSUPPORTED_DOC_RE.test(arg) && existsSync(arg)) ?? null
}

function notifyUnsupportedFile(filePath: string): void {
  const ext = extname(filePath).slice(1).toLowerCase() || basename(filePath)
  showAppWarning(tm('errUnsupportedExt', { ext }))
}

/** shell-hosted warning box; focused when a shell window exists, standalone otherwise */
function showAppWarning(message: string): void {
  const options = { type: 'warning' as const, message }
  if (shellWindow) {
    shellWindow.show()
    shellWindow.focus()
    void dialog.showMessageBox(shellWindow, options)
  } else {
    void dialog.showMessageBox(options)
  }
}

/**
 * Files dropped from the OS into any renderer arrive via installDropOpenBridge
 * and route through the normal File > Open pipeline; detached editor windows
 * can host the drop target, so the shell must reveal itself after opening.
 */
function registerDroppedFilesIpc(): void {
  ipcMain.on(DROP_OPEN_CHANNEL, (_event, raw: unknown) =>
    handleDroppedFiles(raw, {
      openDocumentPath,
      revealShellWindow,
      showWarning: showAppWarning,
      unsupportedMessage: (exts) => tm('errUnsupportedExt', { ext: exts.join(', ') }),
    }),
  )
}

/** the single router: extension decides which module owns the file; false = nothing opened */
function openDocumentPath(filePath: string): boolean {
  const opened = routeDocumentPath(filePath)
  if (opened) {
    recordStarPromptDocOpen()
    // extension only — never the file name or path
    analytics.track('file_open', { ext: extname(filePath).slice(1).toLowerCase() })
  }
  return opened
}

/**
 * Open a just-written export. Unlike File > Open, an already-open PDF tab is
 * reloaded from disk so a re-export to the same path shows the new bytes
 * instead of the previous in-memory document (which may also hold unsaved
 * annotations). In-memory edits on that tab are discarded — Save would
 * overwrite the file we just exported.
 */
function openGeneratedDocument(filePath: string): boolean {
  if (tabManager && PDF_RE.test(filePath)) {
    const existing = tabManager.findPdfTabByPath(filePath)
    if (existing) {
      tabManager.reloadTab(existing)
      tabManager.activateTab(existing)
      return true
    }
  }
  return openDocumentPath(filePath)
}

function routeDocumentPath(filePath: string): boolean {
  if (!existsSync(filePath) || !tabManager) return false
  if (DOCX_RE.test(filePath)) {
    recordRecentFile(filePath)
    const existing = tabManager.findDocsTabByPath(filePath)
    if (existing) tabManager.activateTab(existing)
    else tabManager.openDocsTab(filePath)
    return true
  }
  if (XLSX_RE.test(filePath)) {
    recordRecentFile(filePath)
    const existing = tabManager.findSheetsTabByPath(filePath)
    if (existing) {
      tabManager.activateTab(existing)
    } else {
      tabManager.openSheetsTab(filePath)
      startQueuedWorkbookNudge()
    }
    return true
  }
  if (PPTX_RE.test(filePath)) {
    recordRecentFile(filePath)
    const existing = tabManager.findSlidesTabByPath(filePath)
    if (existing) {
      tabManager.activateTab(existing)
    } else {
      // For a new tab the path goes through the pending queue; the renderer consumes it after mounting
      tabManager.openSlidesTab(filePath)
    }
    return true
  }
  if (PDF_RE.test(filePath)) {
    recordRecentFile(filePath)
    const existing = tabManager.findPdfTabByPath(filePath)
    if (existing) tabManager.activateTab(existing)
    else tabManager.openPdfTab(filePath)
    return true
  }
  if (MD_RE.test(filePath)) {
    recordRecentFile(filePath)
    const existing = tabManager.findMarkdownTabByPath(filePath)
    if (existing) tabManager.activateTab(existing)
    else tabManager.openMarkdownTab(filePath)
    return true
  }
  notifyUnsupportedFile(filePath)
  return false
}

/**
 * "New spreadsheet" creates the backing .xlsx in the default folder up front and
 * opens it as a regular file tab — the blank in-memory demo mode has no save
 * pipeline, so the file must exist before edits. Falls back to the old blank
 * tab if the write fails.
 */
async function newSheetTab(): Promise<void> {
  try {
    const filePath = uniquePathIn(defaultSaveDir(), `${tm('untitledSheet')}.xlsx`)
    writeFileSync(filePath, await blankXlsxBuffer())
    // eligible for content-derived auto-rename after the first AI generation
    markSheetsUntitledPath(filePath)
    // route directly (not via openDocumentPath) so creating a sheet emits
    // only file_new — the file_open event is reserved for opening existing files
    if (routeDocumentPath(filePath)) recordStarPromptDocOpen()
    analytics.track('file_new', { kind: 'xlsx' })
  } catch (err) {
    console.warn('[shell] blank workbook create failed, opening in-memory blank tab:', err)
    try {
      tabManager?.openSheetsTab(undefined, { newBlank: true })
    } catch (fallbackErr) {
      surfaceNewTabError(fallbackErr)
    }
  }
}

/**
 * A throw anywhere in the create-tab path (view creation, sidecar resolution,
 * renderer load) used to be swallowed by `void`-ed promises and ipc-invoke
 * rejections, so the click looked like a pure no-op — the exact "AI Sheets /
 * AI Slides do nothing" alpha report. Surface the failure instead.
 */
function surfaceNewTabError(err: unknown): void {
  console.error('[shell] new tab failed:', err)
  showErrorDialog(shellWindow, tm('errNewTabFailed'), err)
}

function newDocTab(): void {
  try {
    tabManager?.openDocsTab(undefined, { newBlank: true })
    // creating a document is as much a value moment as opening one
    recordStarPromptDocOpen()
    analytics.track('file_new', { kind: 'docx' })
  } catch (err) {
    surfaceNewTabError(err)
  }
}

function newSlideTab(): void {
  try {
    tabManager?.openSlidesTab()
    recordStarPromptDocOpen()
    analytics.track('file_new', { kind: 'pptx' })
  } catch (err) {
    surfaceNewTabError(err)
  }
}

function newMarkdownTab(): void {
  try {
    tabManager?.openMarkdownTab()
    recordStarPromptDocOpen()
    analytics.track('file_new', { kind: 'md' })
  } catch (err) {
    surfaceNewTabError(err)
  }
}

/**
 * "New PDF" creates a blank single-page .pdf in the default folder up front and
 * opens it as a regular file tab — the PDF module has no in-memory blank mode
 * (openPdfTab requires a path), same pattern as the blank workbook above.
 */
async function newPdfTab(): Promise<void> {
  try {
    const filePath = uniquePathIn(defaultSaveDir(), `${tm('untitledPdf')}.pdf`)
    writeFileSync(filePath, await blankPdfBuffer())
    // Opt the file into content-derived auto-naming on its first save
    markPdfUntitledPath(filePath)
    // PDF has no opened/saved shell hook — assign the pending project right here
    applyPendingProject(filePath)
    // route directly (not via openDocumentPath) so creating a pdf emits only
    // file_new and counts one doc-open — same as the blank workbook above
    if (routeDocumentPath(filePath)) recordStarPromptDocOpen()
    analytics.track('file_new', { kind: 'pdf' })
  } catch (err) {
    surfaceNewTabError(err)
  }
}

/**
 * The sheets renderer subscribes to menu actions only after Univer finishes
 * mounting (seconds on cold start), so a single 'open' can fire into the
 * void. Re-send until the queued workbook is consumed; consumption clears the
 * queue entry main-side (sheets-main), which stops the loop. The nudge only
 * reaches the active tab, so it gates on that tab's own queue entry —
 * background tabs from a multi-select Open pull their path themselves via the
 * renderer's has-queued-workbook poll.
 */
let workbookNudgeTimer: ReturnType<typeof setInterval> | null = null

function startQueuedWorkbookNudge(): void {
  if (workbookNudgeTimer) clearInterval(workbookNudgeTimer)
  const startedAt = Date.now()
  sendSheetsMenuAction('open')
  workbookNudgeTimer = setInterval(() => {
    if (
      !hasActiveQueuedWorkbook() ||
      Date.now() - startedAt > 30_000 ||
      !tabManager?.findSheetsTab()
    ) {
      if (workbookNudgeTimer) clearInterval(workbookNudgeTimer)
      workbookNudgeTimer = null
      return
    }
    sendSheetsMenuAction('open')
  }, 700)
}

// ---- home IPC ----

function statEntries(paths: string[]): RecentEntry[] {
  return statExistingPaths(paths, new Set(readStarredFiles()))
}

function registerHomeIpc(): void {
  // signed-in means GenOffice's own device-code login; the shared gsk CLI key
  // is only a silent fallback, deliberately not shown here to nudge users onto our key
  ipcMain.handle(HOME_CHANNELS.accountStatus, async () => {
    if (!loadGenofficeAuth()) return { loggedIn: false }
    await getProxyBootstrap()
    const info = await gskLoginInfo()
    return info
      ? { loggedIn: true, email: info.email, creditBalance: info.creditBalance }
      : { loggedIn: true }
  })

  // login progress is streamed to the requesting renderer; the auth URL is
  // kept main-side so the "open manually" rescue never opens a renderer-supplied URL
  let pendingLoginUrl = ''
  ipcMain.handle(HOME_CHANNELS.accountLogin, async (event) => {
    analytics.track('login_click')
    const sender = event.sender
    pendingLoginUrl = ''
    await getProxyBootstrap()
    const send = (payload: AccountLoginEvent) => {
      if (!sender.isDestroyed()) sender.send(HOME_CHANNELS.accountLoginEvent, payload)
    }
    // open the browser on the first url event only; later events refresh the rescue URL
    let opened = false
    const launched = startGenofficeLogin((progress) => {
      if (progress.url) {
        pendingLoginUrl = progress.url
        if (!opened) {
          opened = true
          void shell.openExternal(progress.url)
        }
      }
      if (progress.phase === 'success') analytics.track('login_success')
      send(progress)
    })
    if (launched) send({ phase: 'launched' })
    return launched
  })

  ipcMain.handle(HOME_CHANNELS.accountLoginOpenUrl, () => {
    if (pendingLoginUrl) void shell.openExternal(pendingLoginUrl)
  })

  ipcMain.handle(HOME_CHANNELS.accountLogout, async () => {
    await genofficeLogout()
    // the cloud projects cache belongs to the account that just signed out
    clearCloudProjectsStore(cloudProjectsStorePath())
  })

  ipcMain.handle(HOME_CHANNELS.getAppVersion, (): string => app.getVersion())

  ipcMain.handle(HOME_CHANNELS.recents, (_event, query: unknown): RecentPage =>
    pageRecentPaths(readRecentFiles(), query, new Set(readStarredFiles())),
  )

  // Starred files sort by mtime, which requires stat-ing them all first; they are hand-picked and few, so this is fine
  ipcMain.handle(HOME_CHANNELS.starred, (_event, query: unknown): RecentPage => {
    const { offset, limit, ext } = normalizeRecentQuery(query)
    const all = statEntries(readStarredFiles()).sort((a, b) => b.mtimeMs - a.mtimeMs)
    const filtered = ext ? all.filter((entry) => entry.ext === ext) : all
    return {
      entries: limit === 0 ? [] : filtered.slice(offset, offset + limit),
      total: filtered.length,
      totalAll: all.length,
    }
  })

  ipcMain.handle(HOME_CHANNELS.statPaths, (_event, paths: unknown): RecentEntry[] =>
    statEntries(stringPaths(paths)),
  )

  ipcMain.handle(HOME_CHANNELS.toggleStar, (_event, path: unknown) => {
    if (typeof path === 'string') toggleStarredFile(path)
  })

  ipcMain.handle(HOME_CHANNELS.openPath, (_event, path: unknown) => {
    if (typeof path === 'string') openDocumentPath(path)
  })

  ipcMain.handle(HOME_CHANNELS.browse, async (event) => {
    const win = BrowserWindow.fromWebContents(event.sender) ?? shellWindow
    if (!win) return
    const result = await showOpenDialogWithMemory(dialog, win, {
      title: tm('dlgOpenTitle'),
      filters: [
        { name: tm('filterSupported'), extensions: OPEN_DIALOG_EXTENSIONS },
        { name: tm('filterWord'), extensions: ['docx', 'doc'] },
        { name: tm('filterExcel'), extensions: ['xlsx', 'xlsm', 'xls', 'csv'] },
        { name: tm('filterPpt'), extensions: ['pptx', 'ppt'] },
        { name: tm('filterPdf'), extensions: ['pdf'] },
        { name: tm('filterMarkdown'), extensions: ['md', 'markdown'] },
      ],
      properties: ['openFile', 'multiSelections'],
    })
    if (!result.canceled) for (const path of result.filePaths) openDocumentPath(path)
  })

  ipcMain.handle(HOME_CHANNELS.newDoc, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('doc', opts.projectId)
    }
    newDocTab()
  })

  ipcMain.handle(HOME_CHANNELS.newSheet, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('sheet', opts.projectId)
    }
    void newSheetTab()
  })

  ipcMain.handle(HOME_CHANNELS.newSlide, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('slide', opts.projectId)
    }
    newSlideTab()
  })

  ipcMain.handle(HOME_CHANNELS.newMarkdown, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('markdown', opts.projectId)
    }
    newMarkdownTab()
  })

  ipcMain.handle(HOME_CHANNELS.newPdf, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('pdf', opts.projectId)
    }
    void newPdfTab()
  })

  ipcMain.handle(HOME_CHANNELS.removeRecent, (_event, paths: unknown) => {
    removeRecentFiles(stringPaths(paths))
  })

  ipcMain.handle(HOME_CHANNELS.revealPath, (_event, path: unknown) => {
    if (typeof path === 'string' && existsSync(path)) shell.showItemInFolder(path)
  })

  ipcMain.handle(
    HOME_CHANNELS.renameFile,
    (_event, path: unknown, newName: unknown): RenameResult => {
      if (typeof path !== 'string' || typeof newName !== 'string')
        return { ok: false, error: tm('errBadArgs') }
      const name = newName.trim()
      if (!name || /[\\/:]/.test(name)) return { ok: false, error: tm('errBadName') }
      if (!existsSync(path)) return { ok: false, error: tm('errMissing') }
      const target = join(dirname(path), name)
      if (target === path) return { ok: true, path }
      if (existsSync(target)) return { ok: false, error: tm('errExists') }
      try {
        renameSync(path, target)
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : tm('errRenameFailed') }
      }
      replaceRecentFile(path, target)
      // project-store's fileMap/chatIdByPath re-key too, so AI chat history follows the file
      projectFileRenamed(path, target)
      // the slides module's own recent list switches to the new path as well (used by the start screen)
      if (/\.pptx$/i.test(target)) void replaceSlidesRecentFile(path, target)
      // open tabs sync their title/path; each editor then syncs its internal save path and title bar
      const affected = tabManager?.renameTabFile(path, target) ?? []
      for (const t of affected) {
        if (t.kind === 'slides') slidesFileRenamed(t.webContents, path, target)
        else if (t.kind === 'docs') docsFileRenamed(t.webContents, path, target)
        else if (t.kind === 'sheets') sheetsFileRenamed(t.webContents, path, target)
        else if (t.kind === 'markdown') markdownFileRenamed(t.webContents, path, target)
      }
      return { ok: true, path: target }
    },
  )

  ipcMain.handle(HOME_CHANNELS.duplicateFile, (_event, path: unknown) => {
    if (typeof path !== 'string' || !existsSync(path)) return
    const ext = extname(path)
    const base = basename(path, ext)
    const dir = dirname(path)
    for (let i = 1; ; i++) {
      const target = join(dir, `${base} ${tm('copySuffix')}${i === 1 ? '' : ` ${i}`}${ext}`)
      if (existsSync(target)) continue
      copyFileSync(path, target)
      recordRecentFile(target)
      return
    }
  })

  ipcMain.handle(HOME_CHANNELS.deleteFiles, async (_event, paths: unknown) => {
    const list = stringPaths(paths)
    for (const p of list) {
      try {
        await shell.trashItem(p)
      } catch {
        // file already gone or trash unavailable; still drop it from the list
      }
    }
    removeRecentFiles(list)
  })

  ipcMain.handle(HOME_CHANNELS.openTrash, () => {
    if (process.platform === 'darwin') {
      void shell.openPath(join(app.getPath('home'), '.Trash'))
    } else if (process.platform === 'win32') {
      spawn('explorer.exe', ['shell:RecycleBin'], { detached: true }).unref()
    } else {
      void shell.openPath(join(app.getPath('home'), '.local', 'share', 'Trash', 'files'))
    }
  })

  ipcMain.handle(HOME_CHANNELS.getLanguage, (): Lang => currentLang())

  ipcMain.handle(HOME_CHANNELS.setLanguage, (_event, lang: unknown) => {
    if (!isLang(lang) || lang === currentLang()) return
    persistLang(lang)
    // the switcher lives on the home page, so the home menu is the active one
    buildHomeMenu()
    installDockMenu()
    installBackToHomeItems()
    for (const wc of webContents.getAllWebContents()) wc.send('app:language-changed', lang)
  })

  ipcMain.handle(HOME_CHANNELS.getUpdateChannel, (): UpdateChannel => currentUpdateChannel())

  ipcMain.handle(HOME_CHANNELS.setUpdateChannel, (_event, channel: unknown) => {
    if (!isUpdateChannel(channel) || channel === currentUpdateChannel()) return
    cachedUpdateChannel = channel
    writeAppSetting(APP_SETTINGS_PATH(), 'updateChannel', channel)
    applyUpdateChannel(channel)
  })

  ipcMain.handle(
    HOME_CHANNELS.onboardingSeen,
    (): boolean => readAppSettings(APP_SETTINGS_PATH()).onboardingSeen === true,
  )

  ipcMain.handle(HOME_CHANNELS.setOnboardingSeen, (): boolean => {
    try {
      writeAppSetting(APP_SETTINGS_PATH(), 'onboardingSeen', true)
      return true
    } catch {
      return false
    }
  })

  ipcMain.handle(HOME_CHANNELS.getTheme, (): UiTheme => currentTheme())
  // editor tabs ask via the app-wide channel (symmetric with app:get-language)
  ipcMain.handle('app:get-theme', (): UiTheme => currentTheme())

  ipcMain.handle(HOME_CHANNELS.setTheme, (_event, theme: unknown) => {
    if (theme !== 'light' && theme !== 'dark' && theme !== 'system') return
    if (theme === currentTheme()) return
    cachedTheme = theme
    writeAppSetting(APP_SETTINGS_PATH(), 'theme', theme)
    nativeTheme.themeSource = theme
    for (const wc of webContents.getAllWebContents()) wc.send('app:theme-changed', theme)
  })

  ipcMain.handle(HOME_CHANNELS.getAnalyticsEnabled, (): boolean => analyticsEnabled())

  ipcMain.handle(HOME_CHANNELS.setAnalyticsEnabled, (_event, enabled: unknown): boolean => {
    if (typeof enabled !== 'boolean') return false
    return persistAnalyticsPreference(enabled)
  })

  // effective folder where new/untitled files land; the editor mains resolve
  // the same setting themselves (configuredDefaultSaveDir via docs' defaultSaveDir)
  ipcMain.handle(HOME_CHANNELS.getDefaultSaveDir, (): string => defaultSaveDir())

  ipcMain.handle(HOME_CHANNELS.pickDefaultSaveDir, async (): Promise<string | null> => {
    const result = await showOpenDialogWithMemory(dialog, shellWindow, {
      title: tm('dlgPickSaveDir'),
      defaultPath: defaultSaveDir(),
      properties: ['openDirectory', 'createDirectory'],
    })
    const picked = result.filePaths[0]
    if (result.canceled || !picked) return null
    if (!isUsableSaveDir(picked)) {
      showErrorDialog(shellWindow, tm('errSaveDirUnusable'), picked)
      return null
    }
    writeAppSetting(APP_SETTINGS_PATH(), DEFAULT_SAVE_DIR_KEY, picked)
    return picked
  })

  ipcMain.handle(HOME_CHANNELS.openGenTeam, () => {
    shell.openExternal(GENTEAM_URL).catch(() => {
      // no browser handler available; nothing actionable for the user here
    })
  })

  ipcMain.handle(HOME_CHANNELS.openCreditUsage, () => {
    shell.openExternal(CREDIT_USAGE_URL).catch(() => {
      // no browser handler available; nothing actionable for the user here
    })
  })

  ipcMain.handle(HOME_CHANNELS.openGitHubRepo, () => {
    shell.openExternal(GITHUB_REPO_URL).catch(() => {
      // no browser handler available; nothing actionable for the user here
    })
  })

  ipcMain.handle(HOME_CHANNELS.githubStars, () => fetchGithubStars())

  // returning true also counts as "shown": the renderer displays it
  // unconditionally, so no separate mark-shown round-trip is needed
  ipcMain.handle(HOME_CHANNELS.starPromptShouldShow, (): StarPromptShow => {
    if (starPromptSessionGrant) return starPromptSessionGrant
    const now = Date.now()
    const state = readStarPrompt()
    const docOpens = state.docOpens ?? 0
    // dev preview of the card without waiting out the value thresholds
    // (same pattern as GENOFFICE_FAKE_UPDATE); nothing is recorded
    if (!app.isPackaged && process.env.GENOFFICE_FORCE_STAR_PROMPT) return { show: true, docOpens }
    const grant = (): StarPromptShow => {
      writeStarPrompt(withShown(state, now))
      starPromptSessionGrant = { show: true, docOpens }
      return starPromptSessionGrant
    }
    // first launch after an upgrade: skip the value gates once for a
    // never-prompted user (they are a proven repeat user already)
    if (upgradeStarPromptPending) {
      upgradeStarPromptPending = false
      if (shouldShowUpgradeStarPrompt(state)) return grant()
    }
    if (!shouldShowStarPrompt(state, now)) return { show: false, docOpens }
    return grant()
  })

  ipcMain.handle(HOME_CHANNELS.starPromptAction, (_event, action: unknown) => {
    if (action !== 'starred' && action !== 'later') return
    // the card was reacted to — drop the session grant so a later query (new
    // shell window on macOS) re-evaluates the real rules (snooze / resolved)
    starPromptSessionGrant = null
    // 'later' needs no write: the display was already counted by the query
    if (action === 'starred') writeStarPrompt(withResolved(readStarPrompt()))
  })

  const cloudProjectsStorePath = () => join(app.getPath('userData'), 'cloud-projects.json')

  ipcMain.handle(HOME_CHANNELS.cloudProjectsCached, () =>
    readCloudProjectsStore(cloudProjectsStorePath()),
  )

  ipcMain.handle(HOME_CHANNELS.cloudProjects, () => syncCloudProjects(cloudProjectsStorePath()))

  ipcMain.handle(HOME_CHANNELS.openCloudProject, (_event, projectUrl: unknown) => {
    const url = cloudProjectExternalUrl(projectUrl)
    if (url) void shell.openExternal(url)
  })
}

function stringPaths(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((p): p is string => typeof p === 'string') : []
}

// electron-vite emits ?asset files under hashed names, which breaks nativeImage's
// automatic `@2x` sibling lookup — attach the retina representation by hand
function loadMenuIcon(path1x: string, path2x: string): NativeImage {
  const icon = nativeImage.createFromPath(path1x)
  icon.addRepresentation({ scaleFactor: 2, buffer: readFileSync(path2x) })
  return icon
}

// loaded once, not on every menu open
interface MenuIconSet {
  docx: NativeImage
  xlsx: NativeImage
  pptx: NativeImage
  pdf: NativeImage
  md: NativeImage
  home: NativeImage
}
let menuIconCache: MenuIconSet | null = null
function menuIcons(): MenuIconSet {
  menuIconCache ??= {
    docx: loadMenuIcon(menuDocxIcon1x, menuDocxIcon2x),
    xlsx: loadMenuIcon(menuXlsxIcon1x, menuXlsxIcon2x),
    pptx: loadMenuIcon(menuPptxIcon1x, menuPptxIcon2x),
    pdf: loadMenuIcon(menuPdfIcon1x, menuPdfIcon2x),
    md: loadMenuIcon(menuMdIcon1x, menuMdIcon2x),
    home: loadMenuIcon(menuHomeIcon1x, menuHomeIcon2x),
  }
  return menuIconCache
}

const TAB_MENU_ICON: Record<TabKind, keyof MenuIconSet> = {
  home: 'home',
  docs: 'docx',
  sheets: 'xlsx',
  slides: 'pptx',
  pdf: 'pdf',
  markdown: 'md',
}

// tab views see neither DOM events nor a focus change when the user clicks the
// shell chrome — relay the press so open popovers in documents can dismiss.
// The pressed document must be excluded: it already dismissed (or is opening)
// its own popovers via its local pointerdown listeners, and the async IPC
// round-trip would otherwise close a popover that very press just opened
// (home row menus died this way: pointerdown → broadcast → menu unmounts
// before the click event ever reached the menu item).
function broadcastChromePressed(exclude?: WebContents): void {
  for (const wc of webContents.getAllWebContents()) {
    if (wc !== exclude) wc.send('app:chrome-pressed')
  }
}

function registerTabsIpc(): void {
  ipcMain.on(TABS_CHANNELS.chromePressed, (event) => broadcastChromePressed(event.sender))
  ipcMain.handle(TABS_CHANNELS.list, () => tabManager?.list() ?? [])
  ipcMain.handle(TABS_CHANNELS.activate, (_event, id: string) => tabManager?.activateTab(id))
  ipcMain.handle(TABS_CHANNELS.close, (_event, id: string) => tabManager?.closeTab(id))
  ipcMain.handle(TABS_CHANNELS.reorder, (_event, id: string, toIndex: number) => {
    if (typeof id === 'string' && Number.isInteger(toIndex)) tabManager?.reorderTab(id, toIndex)
  })
  // "all tabs" overflow menu — native popup because the editors' WebContentsView
  // would cover any DOM dropdown the shell renderer draws below the tab strip
  ipcMain.handle(TABS_CHANNELS.showMenu, (_event, x: unknown, y: unknown) => {
    if (!tabManager || !shellWindow) return
    const menu = Menu.buildFromTemplate(
      tabManager.list().map((tab) => ({
        label: tab.title,
        type: 'checkbox' as const,
        checked: tab.active,
        icon: menuIcons()[TAB_MENU_ICON[tab.kind]],
        click: () => tabManager?.activateTab(tab.id),
      })),
    )
    menu.popup({
      window: shellWindow,
      ...(typeof x === 'number' && typeof y === 'number'
        ? { x: Math.round(x), y: Math.round(y) }
        : {}),
    })
  })
  // "+" new-file menu — native for the same reason as the tab list above
  ipcMain.handle(TABS_CHANNELS.showNewMenu, (_event, x: unknown, y: unknown) => {
    if (!tabManager || !shellWindow) return
    const menu = Menu.buildFromTemplate([
      // enabled:false so pre-Sonoma macOS / Windows (no 'header' support) degrade
      // to an inert label instead of a clickable no-op item
      { label: tm('menuSectionNew'), type: 'header', enabled: false },
      {
        label: tm('menuNewDoc'),
        icon: menuIcons().docx,
        click: () => newDocTab(),
      },
      {
        label: tm('menuNewSheet'),
        icon: menuIcons().xlsx,
        click: () => void newSheetTab(),
      },
      {
        label: tm('menuNewSlide'),
        icon: menuIcons().pptx,
        click: () => newSlideTab(),
      },
      {
        label: tm('menuNewMarkdown'),
        icon: menuIcons().md,
        click: () => newMarkdownTab(),
      },
      {
        label: tm('menuNewPdf'),
        icon: menuIcons().pdf,
        click: () => void newPdfTab(),
      },
      { type: 'separator' },
      { label: tm('menuOpen'), click: () => void openFileViaDialog() },
    ])
    menu.popup({
      window: shellWindow,
      ...(typeof x === 'number' && typeof y === 'number'
        ? { x: Math.round(x), y: Math.round(y) }
        : {}),
    })
  })
}

// ---- home menu ----

async function openFileViaDialog(): Promise<void> {
  const win = shellWindow ?? BrowserWindow.getFocusedWindow()
  if (!win) return
  const result = await showOpenDialogWithMemory(dialog, win, {
    filters: [{ name: tm('filterSupported'), extensions: OPEN_DIALOG_EXTENSIONS }],
    properties: ['openFile', 'multiSelections'],
  })
  if (!result.canceled) for (const path of result.filePaths) openDocumentPath(path)
}

function buildHomeMenu(): void {
  const isMac = process.platform === 'darwin'
  const template: MenuItemConstructorOptions[] = [
    ...(isMac ? [{ role: 'appMenu' as const }] : []),
    {
      label: tm('menuFile'),
      submenu: [
        { label: tm('menuSectionNew'), type: 'header', enabled: false },
        {
          label: tm('menuNewDoc'),
          accelerator: 'CmdOrCtrl+N',
          click: () => newDocTab(),
        },
        {
          label: tm('menuNewSheet'),
          click: () => void newSheetTab(),
        },
        { label: tm('menuNewSlide'), click: () => newSlideTab() },
        { label: tm('menuNewMarkdown'), click: () => newMarkdownTab() },
        { label: tm('menuNewPdf'), click: () => void newPdfTab() },
        { type: 'separator' },
        {
          label: tm('menuOpen'),
          accelerator: 'CmdOrCtrl+O',
          click: () => void openFileViaDialog(),
        },
        { type: 'separator' },
        { role: 'close', label: tm('menuClose') },
      ],
    },
    editMenuTemplate(process.platform, appMenuLabels(currentLang())),
    windowMenuTemplate(process.platform, appMenuLabels(currentLang())),
    {
      role: 'help',
      label: tm('menuHelp'),
      submenu: [{ label: tm('thirdPartyNotices'), click: () => void openThirdPartyNotices() }],
    },
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

// ---- pdf menu (pdf-main has no menu of its own; the shell owns pdf tabs, so it builds one) ----

function buildPdfMenu(): void {
  const isMac = process.platform === 'darwin'
  const template: MenuItemConstructorOptions[] = [
    ...(isMac ? [{ role: 'appMenu' as const }] : []),
    {
      label: tm('menuFile'),
      submenu: [
        {
          label: tm('menuOpen'),
          accelerator: 'CmdOrCtrl+O',
          click: () => void openFileViaDialog(),
        },
        { type: 'separator' },
        {
          label: tm('backToHome'),
          accelerator: 'Shift+CmdOrCtrl+H',
          click: () => tabManager?.openHomeTab(),
        },
        { type: 'separator' },
        {
          label: tm('menuSave'),
          accelerator: 'CmdOrCtrl+S',
          click: () => {
            const tab = tabManager?.activePdfTab()
            if (tab) void flushPdfSave(tab.webContents)
          },
        },
        {
          label: tm('menuSaveAs'),
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => void savePdfAs(),
        },
        { type: 'separator' },
        // local pdf2docx is the default Word export; cloud stays as a
        // secondary option because scanned PDFs still need its OCR — hidden
        // entirely when gsk is unavailable (not signed in / CLI missing)
        {
          label: tm('menuExportDocx'),
          click: () => void exportPdfAsDocxLocal(),
        },
        ...(hasGskAuth()
          ? [
              {
                label: tm('menuExportDocxCloud'),
                click: () => void exportPdfAsDocx(),
              },
            ]
          : []),
        // local pdf2pptx (P25): one slide per page, no cloud counterpart
        {
          label: tm('menuExportPptx'),
          click: () => void exportPdfAsPptxLocal(),
        },
        // local pdf2xlsx (P26): one worksheet per page, no cloud counterpart
        {
          label: tm('menuExportXlsx'),
          click: () => void exportPdfAsXlsxLocal(),
        },
        { type: 'separator' },
        {
          label: tm('menuPrint'),
          accelerator: 'CmdOrCtrl+P',
          click: () => {
            const tab = tabManager?.activePdfTab()
            if (tab) sendPdfPrintRequest(tab.webContents)
          },
        },
        { type: 'separator' },
        {
          label: tm('menuClose'),
          accelerator: 'CmdOrCtrl+W',
          click: () => tabManager?.closeActiveTab(),
        },
      ],
    },
    editMenuTemplate(process.platform, appMenuLabels(currentLang())),
    windowMenuTemplate(process.platform, appMenuLabels(currentLang())),
    {
      role: 'help',
      label: tm('menuHelp'),
      submenu: [{ label: tm('thirdPartyNotices'), click: () => void openThirdPartyNotices() }],
    },
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

// ---- markdown menu (markdown-main has no menu of its own; the shell owns markdown tabs) ----

function buildMarkdownMenu(): void {
  const isMac = process.platform === 'darwin'
  const template: MenuItemConstructorOptions[] = [
    ...(isMac ? [{ role: 'appMenu' as const }] : []),
    {
      label: tm('menuFile'),
      submenu: [
        {
          label: tm('menuOpen'),
          accelerator: 'CmdOrCtrl+O',
          click: () => void openFileViaDialog(),
        },
        { type: 'separator' },
        {
          label: tm('backToHome'),
          accelerator: 'Shift+CmdOrCtrl+H',
          click: () => tabManager?.openHomeTab(),
        },
        { type: 'separator' },
        {
          label: tm('menuSave'),
          accelerator: 'CmdOrCtrl+S',
          click: () => {
            const tab = tabManager?.activeMarkdownTab()
            if (tab) void requestMarkdownSave(tab.webContents, 'save')
          },
        },
        {
          label: tm('menuSaveAs'),
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => {
            const tab = tabManager?.activeMarkdownTab()
            if (tab) void requestMarkdownSave(tab.webContents, 'saveAs')
          },
        },
        { type: 'separator' },
        {
          label: tm('menuExportDocx'),
          click: () => {
            const tab = tabManager?.activeMarkdownTab()
            if (tab) sendMarkdownExportRequest(tab.webContents, 'docx')
          },
        },
        {
          label: tm('menuExportPdf'),
          click: () => {
            const tab = tabManager?.activeMarkdownTab()
            if (tab) sendMarkdownExportRequest(tab.webContents, 'pdf')
          },
        },
        {
          label: tm('menuOpenInDocs'),
          click: () => {
            const tab = tabManager?.activeMarkdownTab()
            if (tab) sendMarkdownExportRequest(tab.webContents, 'docs')
          },
        },
        { type: 'separator' },
        {
          label: tm('menuPrint'),
          accelerator: 'CmdOrCtrl+P',
          click: () => {
            const tab = tabManager?.activeMarkdownTab()
            if (tab) sendMarkdownPrintRequest(tab.webContents)
          },
        },
        { type: 'separator' },
        {
          label: tm('menuClose'),
          accelerator: 'CmdOrCtrl+W',
          click: () => tabManager?.closeActiveTab(),
        },
      ],
    },
    editMenuTemplate(process.platform, appMenuLabels(currentLang())),
    windowMenuTemplate(process.platform, appMenuLabels(currentLang())),
    {
      role: 'help',
      label: tm('menuHelp'),
      submenu: [{ label: tm('thirdPartyNotices'), click: () => void openThirdPartyNotices() }],
    },
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

/**
 * Save As for pdf tabs: write pending edits to the picked path only, then open the copy.
 * Non-destructive: the original file is never written, and a cancelled dialog changes
 * nothing on disk (dialog first, no flush into the source).
 */
/** In-flight guard (same pattern as exportPdfAsDocx): a re-trigger while the dialog
    or write is active must not start a second flow that overwrites the first one's
    waiter/target grant or clears its autosave pause early */
let savingPdfAs = false

async function savePdfAs(): Promise<void> {
  const tab = tabManager?.activePdfTab()
  if (!tab?.filePath || !shellWindow || savingPdfAs) return
  savingPdfAs = true
  // Pause renderer autosave for the whole flow: the dialog blurs the window, and a
  // blur-triggered autosave would write the pending edits into the original file
  setPdfSaveAsInFlight(tab.webContents, true)
  try {
    const picked = await showSaveDialogWithMemory(dialog, shellWindow, {
      defaultPath: tab.filePath,
      filters: [{ name: tm('filterPdf'), extensions: ['pdf'] }],
    })
    if (picked.canceled || !picked.filePath || picked.filePath === tab.filePath) return
    if (pdfIsDirty(tab.webContents.id)) {
      // Renderer applies its pending edits onto the source bytes; the pdf main
      // process writes the result to the picked path only
      if (!(await requestPdfSaveAs(tab.webContents, picked.filePath))) return
    } else {
      // No pending edits → a byte-identical copy
      copyFileSync(tab.filePath, picked.filePath)
    }
    openDocumentPath(picked.filePath)
  } finally {
    savingPdfAs = false
    setPdfSaveAsInFlight(tab.webContents, false)
  }
}

/**
 * In-flight guard: covers the whole flow (dialogs included, conversion takes
 * ~10s+) so re-triggering from the menu can never start a second paid conversion
 */
let exportingPdfDocx = false

/**
 * Export as Word for pdf tabs: flush pending edits, confirm the 5-credit cost,
 * pick the destination, then upload + cloud-convert via gsk file_convert. Not
 * logged in → offer browser login and let the user re-trigger the export
 * afterwards. The destination is picked before converting so cancelling the
 * save dialog never wastes a paid conversion.
 */
async function exportPdfAsDocx(): Promise<void> {
  const tab = tabManager?.activePdfTab()
  if (!tab?.filePath || !shellWindow) return
  if (exportingPdfDocx) {
    // Re-triggered while a previous export (dialogs or cloud conversion) is
    // still in flight: tell the user instead of silently ignoring the click.
    void dialog.showMessageBox(shellWindow, {
      type: 'info',
      message: tm('pdfDocxBusyMsg'),
    })
    return
  }
  exportingPdfDocx = true
  try {
    if (!(await flushPdfSave(tab.webContents))) return
    if (!hasGskAuth()) {
      // hasGskAuth() is also false when the gsk CLI itself cannot be resolved
      // (broken install); Sign In could not launch in that case, so surface
      // the real problem instead of a login dialog that cannot succeed.
      if (!resolveGskEntry()) {
        void dialog.showMessageBox(shellWindow, {
          type: 'error',
          message: tm('pdfDocxNoCliMsg'),
        })
        return
      }
      const { response } = await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLoginMsg'),
        detail: tm('pdfDocxLoginDetail'),
        buttons: [tm('pdfDocxBtnLogin'), tm('btnCancel')],
        defaultId: 0,
        cancelId: 1,
        noLink: true,
      })
      if (response === 0) ensureGenofficeLogin((url) => void shell.openExternal(url))
      return
    }
    const balance = (await gskLoginInfo())?.creditBalance
    const balanceLine =
      balance === undefined
        ? ''
        : ` ${tm('pdfDocxConfirmBalance', { balance: Math.floor(balance).toLocaleString('en-US') })}`
    const confirm = await dialog.showMessageBox(shellWindow, {
      type: 'question',
      message: tm('pdfDocxConfirmMsg'),
      detail: `${tm('pdfDocxConfirmDetail')}${balanceLine}`,
      buttons: [tm('pdfDocxBtnConvert'), tm('btnCancel')],
      defaultId: 0,
      cancelId: 1,
      noLink: true,
    })
    if (confirm.response !== 0) return
    const picked = await showSaveDialogWithMemory(dialog, shellWindow, {
      defaultPath: tab.filePath.replace(/\.pdf$/i, '.docx'),
      filters: [{ name: tm('filterWord'), extensions: ['docx'] }],
    })
    if (picked.canceled || !picked.filePath) return
    // If the destination is already open in a docs tab, close it first (its
    // normal unsaved-changes guard applies) so the converted file opens fresh
    // instead of leaving a stale tab whose next save would clobber the result.
    // Cancelling the close aborts the export before any credits are spent.
    const staleTabId = tabManager?.findDocsTabByPath(picked.filePath)
    if (staleTabId) {
      await tabManager?.closeTab(staleTabId)
      // closeTab activates the docs tab for its unsaved-changes prompt (and a
      // fallback tab after a successful close), so bring the pdf tab back
      // either way — especially when the user cancels and the export aborts.
      tabManager?.activateTab(tab.id)
      if (tabManager?.findDocsTabByPath(picked.filePath)) return
    }
    shellWindow.setProgressBar(2)
    const bytes = await gskConvertPdfToDocx(tab.filePath)
    writeFileSync(picked.filePath, bytes)
    openDocumentPath(picked.filePath)
  } catch (err) {
    if (shellWindow && !shellWindow.isDestroyed()) {
      void dialog.showMessageBox(shellWindow, {
        type: 'error',
        message: tm('pdfDocxFailedMsg'),
        detail: err instanceof Error ? err.message : String(err),
      })
    }
  } finally {
    exportingPdfDocx = false
    if (shellWindow && !shellWindow.isDestroyed()) shellWindow.setProgressBar(-1)
  }
}

/**
 * Export as Word for pdf tabs, fully local (pdf2docx P4): flush pending
 * edits, pick the destination, convert in-process via PDFium wasm, write the
 * file and open it in a Docs tab. No login, no credits. Shares the in-flight
 * guard with the cloud export so the two can never run concurrently.
 */
async function exportPdfAsDocxLocal(): Promise<void> {
  const tab = tabManager?.activePdfTab()
  if (!tab?.filePath || !shellWindow) return
  if (exportingPdfDocx) {
    void dialog.showMessageBox(shellWindow, {
      type: 'info',
      message: tm('pdfDocxBusyMsg'),
    })
    return
  }
  exportingPdfDocx = true
  try {
    if (!(await flushPdfSave(tab.webContents))) return
    const picked = await showSaveDialogWithMemory(dialog, shellWindow, {
      defaultPath: tab.filePath.replace(/\.pdf$/i, '.docx'),
      filters: [{ name: tm('filterWord'), extensions: ['docx'] }],
    })
    if (picked.canceled || !picked.filePath) return
    // same stale-tab handling as the cloud export (see exportPdfAsDocx)
    const staleTabId = tabManager?.findDocsTabByPath(picked.filePath)
    if (staleTabId) {
      await tabManager?.closeTab(staleTabId)
      tabManager?.activateTab(tab.id)
      if (tabManager?.findDocsTabByPath(picked.filePath)) return
    }
    shellWindow.setProgressBar(2)
    // encrypted PDFs prompt for the password (P23), looping on wrong entries;
    // null result = user cancelled the prompt → abort silently
    const pdfPath = tab.filePath
    const result = await convertPdfFileToDocxLocalWithPrompt(
      pdfPath,
      (retry) =>
        promptPdfPassword(shellWindow, {
          fileName: basename(pdfPath),
          retry,
          busy: false,
          lang: currentLang(),
          strings: {
            title: tm('pdfPwdTitle'),
            prompt: tm('pdfPwdPrompt'),
            retryPrompt: tm('pdfPwdRetryPrompt'),
            ok: tm('pdfPwdOk'),
            cancel: tm('btnCancel'),
            verifying: tm('pdfPwdVerifying'),
            label: tm('pdfPwdLabel'),
            placeholder: tm('pdfPwdPlaceholder'),
            show: tm('pdfPwdShow'),
            hide: tm('pdfPwdHide'),
          },
        }),
      (page, total) => {
        if (shellWindow && !shellWindow.isDestroyed() && total > 0) {
          shellWindow.setProgressBar(page / total)
        }
      },
    )
    if (result === null) return
    writeFileSync(picked.filePath, result.docx)

    // degrade transparency (plan §7.6 dual-track split): whole scan → point
    // to the cloud/OCR flow; individual image-fallback pages → name them;
    // OCR-recovered scans ('ocr') are SUCCESSES — announce the recovery (the
    // user should proofread machine-read text), never the image-export notice
    const ocrPages = result.pageResults.filter((r) => r.status === 'ocr').map((r) => r.page)
    const imagePages = result.pageResults
      .filter((r) => r.status !== 'ok' && r.status !== 'ocr')
      .map((r) => r.page)
    if (result.scannedDocument) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalScannedMsg'),
        detail: tm('pdfDocxLocalScannedDetail'),
      })
    } else if (imagePages.length > 0 && ocrPages.length > 0) {
      // mixed documents surface BOTH facts in one dialog: which pages shipped
      // as images and which carry machine-read text the user should proofread
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalDegradedMsg'),
        detail:
          tm('pdfDocxLocalDegradedDetail', { pages: imagePages.join(', ') }) +
          '\n\n' +
          tm('pdfDocxLocalOcrDetail', { pages: ocrPages.join(', ') }),
      })
    } else if (imagePages.length > 0) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalDegradedMsg'),
        detail: tm('pdfDocxLocalDegradedDetail', { pages: imagePages.join(', ') }),
      })
    } else if (ocrPages.length > 0) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalOcrMsg'),
        detail: tm('pdfDocxLocalOcrDetail', { pages: ocrPages.join(', ') }),
      })
    }
    openDocumentPath(picked.filePath)
  } catch (err) {
    if (shellWindow && !shellWindow.isDestroyed()) {
      // structured load failures (P22): password-protected / damaged PDFs get
      // a human-readable explanation instead of the raw PDFium error string
      const detail =
        err instanceof PdfLoadError
          ? err.code === 'password-required'
            ? tm('pdfDocxLocalEncryptedDetail')
            : err.code === 'unsupported'
              ? // certificate-based or otherwise unsupported security (FPDF
                // error 5): a hard PDFium boundary — no password can open it
                // locally, so the message must NOT suggest one (P24 C)
                tm('pdfDocxLocalUnsupportedEncDetail')
              : tm('pdfDocxLocalCorruptDetail')
          : err instanceof Error
            ? err.message
            : String(err)
      void dialog.showMessageBox(shellWindow, {
        type: 'error',
        message: tm('pdfDocxFailedMsg'),
        detail,
      })
    }
  } finally {
    // the prompt window may still be open when the loop exits through cancel
    // or a non-password error thrown mid-retry
    closePdfPasswordDialog()
    exportingPdfDocx = false
    if (shellWindow && !shellWindow.isDestroyed()) shellWindow.setProgressBar(-1)
  }
}

/**
 * Export as PowerPoint for pdf tabs, fully local (pdf2pptx P25): flush
 * pending edits, pick the destination, convert in-process via PDFium wasm,
 * write the file and open it in a Slides tab. No login, no credits. Shares
 * the in-flight guard with the Word exports so pdfium never runs two
 * conversions at once.
 */
async function exportPdfAsPptxLocal(): Promise<void> {
  const tab = tabManager?.activePdfTab()
  if (!tab?.filePath || !shellWindow) return
  if (exportingPdfDocx) {
    void dialog.showMessageBox(shellWindow, {
      type: 'info',
      message: tm('pdfPptxBusyMsg'),
    })
    return
  }
  exportingPdfDocx = true
  try {
    if (!(await flushPdfSave(tab.webContents))) return
    const picked = await showSaveDialogWithMemory(dialog, shellWindow, {
      defaultPath: tab.filePath.replace(/\.pdf$/i, '.pptx'),
      filters: [{ name: tm('filterPpt'), extensions: ['pptx'] }],
    })
    if (picked.canceled || !picked.filePath) return
    // same stale-tab handling as the Word exports (see exportPdfAsDocx),
    // against the slides tab that may already show the destination file
    const staleTabId = tabManager?.findSlidesTabByPath(picked.filePath)
    if (staleTabId) {
      await tabManager?.closeTab(staleTabId)
      tabManager?.activateTab(tab.id)
      if (tabManager?.findSlidesTabByPath(picked.filePath)) return
    }
    shellWindow.setProgressBar(2)
    // encrypted PDFs prompt for the password (P23), looping on wrong entries;
    // null result = user cancelled the prompt → abort silently
    const pdfPath = tab.filePath
    const result = await convertPdfFileToPptxLocalWithPrompt(
      pdfPath,
      (retry) =>
        promptPdfPassword(shellWindow, {
          fileName: basename(pdfPath),
          retry,
          busy: false,
          lang: currentLang(),
          strings: {
            title: tm('pdfPwdTitle'),
            prompt: tm('pdfPwdPrompt'),
            retryPrompt: tm('pdfPwdRetryPrompt'),
            ok: tm('pdfPwdOk'),
            cancel: tm('btnCancel'),
            verifying: tm('pdfPwdVerifying'),
            label: tm('pdfPwdLabel'),
            placeholder: tm('pdfPwdPlaceholder'),
            show: tm('pdfPwdShow'),
            hide: tm('pdfPwdHide'),
          },
        }),
      (page, total) => {
        if (shellWindow && !shellWindow.isDestroyed() && total > 0) {
          shellWindow.setProgressBar(page / total)
        }
      },
    )
    if (result === null) return
    writeFileSync(picked.filePath, result.pptx)

    // degrade transparency (same split as the Word export): whole scan vs
    // individual image-fallback pages
    const imagePages = result.pageResults.filter((r) => r.status !== 'ok').map((r) => r.page)
    if (result.scannedDocument) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalScannedMsg'),
        detail: tm('pdfPptxLocalScannedDetail'),
      })
    } else if (imagePages.length > 0) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalDegradedMsg'),
        detail: tm('pdfDocxLocalDegradedDetail', { pages: imagePages.join(', ') }),
      })
    }
    openDocumentPath(picked.filePath)
  } catch (err) {
    if (shellWindow && !shellWindow.isDestroyed()) {
      // structured load failures (P22): same explanations as the Word export
      const detail =
        err instanceof PdfLoadError
          ? err.code === 'password-required'
            ? tm('pdfDocxLocalEncryptedDetail')
            : err.code === 'unsupported'
              ? tm('pdfDocxLocalUnsupportedEncDetail')
              : tm('pdfDocxLocalCorruptDetail')
          : err instanceof Error
            ? err.message
            : String(err)
      void dialog.showMessageBox(shellWindow, {
        type: 'error',
        message: tm('pdfPptxFailedMsg'),
        detail,
      })
    }
  } finally {
    closePdfPasswordDialog()
    exportingPdfDocx = false
    if (shellWindow && !shellWindow.isDestroyed()) shellWindow.setProgressBar(-1)
  }
}

/**
 * Export as Excel for pdf tabs, fully local (pdf2xlsx P26): flush pending
 * edits, pick the destination, convert in-process via PDFium wasm, write the
 * file and open it in a Sheets tab. No login, no credits. Shares the
 * in-flight guard with the Word/PowerPoint exports so pdfium never runs two
 * conversions at once.
 */
async function exportPdfAsXlsxLocal(): Promise<void> {
  const tab = tabManager?.activePdfTab()
  if (!tab?.filePath || !shellWindow) return
  if (exportingPdfDocx) {
    void dialog.showMessageBox(shellWindow, {
      type: 'info',
      message: tm('pdfXlsxBusyMsg'),
    })
    return
  }
  exportingPdfDocx = true
  try {
    if (!(await flushPdfSave(tab.webContents))) return
    const picked = await showSaveDialogWithMemory(dialog, shellWindow, {
      defaultPath: tab.filePath.replace(/\.pdf$/i, '.xlsx'),
      filters: [{ name: tm('filterExcel'), extensions: ['xlsx'] }],
    })
    if (picked.canceled || !picked.filePath) return
    // same stale-tab handling as the Word exports (see exportPdfAsDocx),
    // against the sheets tab that may already show the destination file
    const staleTabId = tabManager?.findSheetsTabByPath(picked.filePath)
    if (staleTabId) {
      await tabManager?.closeTab(staleTabId)
      tabManager?.activateTab(tab.id)
      if (tabManager?.findSheetsTabByPath(picked.filePath)) return
    }
    shellWindow.setProgressBar(2)
    // encrypted PDFs prompt for the password (P23), looping on wrong entries;
    // null result = user cancelled the prompt → abort silently
    const pdfPath = tab.filePath
    const result = await convertPdfFileToXlsxLocalWithPrompt(
      pdfPath,
      (retry) =>
        promptPdfPassword(shellWindow, {
          fileName: basename(pdfPath),
          retry,
          busy: false,
          lang: currentLang(),
          strings: {
            title: tm('pdfPwdTitle'),
            prompt: tm('pdfPwdPrompt'),
            retryPrompt: tm('pdfPwdRetryPrompt'),
            ok: tm('pdfPwdOk'),
            cancel: tm('btnCancel'),
            verifying: tm('pdfPwdVerifying'),
            label: tm('pdfPwdLabel'),
            placeholder: tm('pdfPwdPlaceholder'),
            show: tm('pdfPwdShow'),
            hide: tm('pdfPwdHide'),
          },
        }),
      (page, total) => {
        if (shellWindow && !shellWindow.isDestroyed() && total > 0) {
          shellWindow.setProgressBar(page / total)
        }
      },
    )
    if (result === null) return
    writeFileSync(picked.filePath, result.xlsx)

    // degrade transparency: pages that could not become cells got a notice
    // row on their worksheet instead of an image (a spreadsheet has none)
    const noticePages = result.pageResults.filter((r) => r.status !== 'ok').map((r) => r.page)
    if (result.scannedDocument) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLocalScannedMsg'),
        detail: tm('pdfXlsxLocalScannedDetail'),
      })
    } else if (noticePages.length > 0) {
      await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfXlsxLocalSkippedMsg'),
        detail: tm('pdfXlsxLocalSkippedDetail', { pages: noticePages.join(', ') }),
      })
    }
    openDocumentPath(picked.filePath)
  } catch (err) {
    if (shellWindow && !shellWindow.isDestroyed()) {
      // structured load failures (P22): same explanations as the Word export
      const detail =
        err instanceof PdfLoadError
          ? err.code === 'password-required'
            ? tm('pdfDocxLocalEncryptedDetail')
            : err.code === 'unsupported'
              ? tm('pdfDocxLocalUnsupportedEncDetail')
              : tm('pdfDocxLocalCorruptDetail')
          : err instanceof Error
            ? err.message
            : String(err)
      void dialog.showMessageBox(shellWindow, {
        type: 'error',
        message: tm('pdfXlsxFailedMsg'),
        detail,
      })
    }
  } finally {
    closePdfPasswordDialog()
    exportingPdfDocx = false
    if (shellWindow && !shellWindow.isDestroyed()) shellWindow.setProgressBar(-1)
  }
}

// The pdf renderer's converter dropdown funnels into the same local conversion
// flows as the File menu items (dialogs, password prompt, in-flight guard included)
ipcMain.handle(PDF_CHANNELS.convertOffice, async (e, format: unknown) => {
  // only the active pdf tab may trigger a conversion (its file is the source)
  if (tabManager?.activePdfTab()?.webContents.id !== e.sender.id) return
  if (format === 'docx') await exportPdfAsDocxLocal()
  else if (format === 'xlsx') await exportPdfAsXlsxLocal()
  else if (format === 'pptx') await exportPdfAsPptxLocal()
})

function openThirdPartyNotices(): Promise<string> {
  const path = app.isPackaged
    ? join(process.resourcesPath, 'THIRD-PARTY-NOTICES.txt')
    : join(app.getAppPath(), 'build', 'THIRD-PARTY-NOTICES.txt')
  return shell.openPath(path)
}

/** every module's File menu gets a way back to the launcher */
function installBackToHomeItems(): void {
  const backToHomeItem: MenuItemConstructorOptions = {
    label: tm('backToHome'),
    accelerator: 'Shift+CmdOrCtrl+H',
    click: () => tabManager?.openHomeTab(),
  }
  setDocsExtraFileMenuItems([backToHomeItem])
  setSheetsExtraFileMenuItems([backToHomeItem])
  setSlidesExtraFileMenuItems([backToHomeItem])
}

function installDockMenu(): void {
  if (process.platform !== 'darwin') return
  app.dock?.setMenu(
    Menu.buildFromTemplate([
      { label: tm('menuHome'), click: () => tabManager?.openHomeTab() },
      {
        label: tm('menuNewDoc'),
        click: () => newDocTab(),
      },
      {
        label: tm('menuNewSheet'),
        click: () => void newSheetTab(),
      },
      { label: tm('menuNewSlide'), click: () => newSlideTab() },
      { label: tm('menuNewMarkdown'), click: () => newMarkdownTab() },
      { label: tm('menuNewPdf'), click: () => void newPdfTab() },
    ]),
  )
}

// ---- lifecycle (the shell is the only owner) ----

let pendingLaunchPath = supportedFileIn(process.argv) ?? unsupportedFileIn(process.argv)

// show() does not un-minimize, and on macOS ⌘W destroys the shell window while the
// app keeps running — either way a file opened from Finder would land out of sight.
function revealShellWindow(): void {
  if (!shellWindow) createShellWindow()
  if (shellWindow?.isMinimized()) shellWindow.restore()
  shellWindow?.show()
  shellWindow?.focus()
}

// On macOS a file opened from Finder is not in argv; it arrives via the open-file event (before ready).
// If another instance already holds the lock, this process exits, and the path must ride along in
// the lock request's additionalData to the surviving instance — so the lock request is deferred
// until ready, after the path is known.
app.on('open-file', (event, filePath) => {
  event.preventDefault()
  if (!app.isReady()) {
    pendingLaunchPath = filePath
    return
  }
  revealShellWindow()
  if (!openDocumentPath(filePath)) tabManager?.openHomeTab()
})

app.on('second-instance', (_event, argv, _cwd, additionalData) => {
  const file =
    supportedFileIn(argv) ??
    unsupportedFileIn(argv) ??
    (additionalData as { launchPath?: string } | null)?.launchPath
  revealShellWindow()
  if (!file || !openDocumentPath(file)) tabManager?.openHomeTab()
})

installNavigationGuard(app)
installContextMenu(app, () => contextMenuLabels(currentLang()))
registerAiIpc()
registerProjectIpc()
registerDocsIpc()
registerHomeIpc()
registerTabsIpc()
registerDroppedFilesIpc()

// sheets' project:resolveChat goes through the handler registered by docs-main; the sessionId reverse lookup hooks in here
setSessionPathResolver(resolveSheetsSessionPath)

/** Dev-only pid marker for the takeover below; scoped to userData like the lock itself. */
const devPidFile = () => join(app.getPath('userData'), 'dev-instance.pid')

app.whenReady().then(async () => {
  const lockData = () => (pendingLaunchPath ? { launchPath: pendingLaunchPath } : {})
  let hasLock = app.requestSingleInstanceLock(lockData())
  if (!hasLock && !app.isPackaged) {
    // Dev watch restart: electron-vite SIGTERMs the previous instance and spawns this
    // one immediately. Chromium turns that SIGTERM into a graceful quit (Node's
    // process.on('SIGTERM') never fires in the main process), and the quit can wedge
    // in the close-confirmation flow — the zombie then keeps the single-instance lock,
    // this instance quits, and electron-vite's on-close handler exits with it, killing
    // the renderer dev server (blank shell window until a manual dev restart).
    // The previous instance is doomed either way: kill it and take over the lock.
    try {
      const oldPid = Number(readFileSync(devPidFile(), 'utf-8').trim())
      if (Number.isFinite(oldPid) && oldPid > 0 && oldPid !== process.pid) {
        // pid-recycling guard: only kill if that pid is still an Electron process
        const cmd = execSync(`ps -o command= -p ${oldPid}`).toString()
        if (cmd.includes('Electron')) process.kill(oldPid, 'SIGKILL')
      }
    } catch {
      // no previous instance recorded / already gone (ps exits non-zero)
    }
    for (let i = 0; i < 20 && !hasLock; i++) {
      await new Promise((r) => setTimeout(r, 150))
      hasLock = app.requestSingleInstanceLock(lockData())
    }
  }
  if (!hasLock) {
    app.quit()
    return
  }
  if (!app.isPackaged) {
    try {
      writeFileSync(devPidFile(), String(process.pid))
    } catch {
      // best-effort: without the marker the next restart just retries the lock
    }
  }

  startMainProcessProxy()
  app.setAccessibilitySupportEnabled(true)
  // Settle the shared uiLang from saved settings BEFORE any tab renderer can
  // ask 'app:get-language': the editor handlers return the i18n module's
  // mutable lang, whose 'zh' default otherwise wins the race for whichever
  // tab loads first (e.g. sheets booting in Chinese while docs shows English).
  currentLang()
  // native menus/dialogs/scrollbars follow the persisted theme from first paint
  nativeTheme.themeSource = currentTheme()
  // stamp the star-prompt install-age clock on the first launch carrying the feature,
  // and detect upgrade launches (version changed since the previous run)
  try {
    const settings = readAppSettings(APP_SETTINGS_PATH())
    const starState = readStarPrompt()
    const stamped = withFirstRun(starState, Date.now())
    if (stamped !== starState) writeStarPrompt(stamped)

    const prevVersion =
      typeof settings[LAST_RUN_VERSION_KEY] === 'string'
        ? (settings[LAST_RUN_VERSION_KEY] as string)
        : null
    const currentVersion = app.getVersion()
    upgradeStarPromptPending = isUpgradeLaunch(
      prevVersion,
      currentVersion,
      settings.onboardingSeen === true,
    )
    if (prevVersion !== currentVersion)
      writeAppSetting(APP_SETTINGS_PATH(), LAST_RUN_VERSION_KEY, currentVersion)
  } catch {
    // settings write failures must never block startup
  }
  initAnalytics()
  analytics.track('app_launch')
  startSheetsCaptureServer()
  createShellWindow()
  // deferred to ready: labels need currentLang(), which reads app.getLocale()
  installBackToHomeItems()
  installDockMenu()
  initAutoUpdater(() => shellWindow, currentUpdateChannel())

  if (!pendingLaunchPath || !openDocumentPath(pendingLaunchPath)) tabManager?.openHomeTab()
  pendingLaunchPath = null

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createShellWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

app.on('before-quit', () => {
  // No close prompt may fall through to "Save" during shutdown
  markSheetsShuttingDown()
  stopSheetsSidecar()
})
