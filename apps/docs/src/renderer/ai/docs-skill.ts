import type { Editor } from '@tiptap/core'
import type { AgentSkill } from '@genoffice/agent-core'
import {
  AGENT_SYSTEM_PROMPT,
  buildDocContext,
  getSelectionScope,
  type AiTrack,
  type NumIds,
} from './protocol'
import {
  AGENT_TOOLS,
  executeTool,
  markDocSeen,
  type AiCommentsAccess,
  type AiHeaderFooterAccess,
  type AiSectionAccess,
  type FrozenSelection,
} from './tools'

/**
 * The docx capability as an AgentSkill: document skeleton context, the five
 * document tools, and the local executor. Future apps register their own
 * skills (Excel / PPT) against the same agent loop.
 */

/** Appended to the system prompt while gsk cloud tools are unavailable. */
const GSK_TOOLS_OFF_NOTE =
  '\nCloud image generation (generate_image) is unavailable in this installation: ' +
  'for pictures use image_search and insert_image instead.'

export function createDocsSkill(
  getEditor: () => Editor,
  getNumIds: () => NumIds,
  getTrack?: () => AiTrack | undefined,
  getComments?: () => AiCommentsAccess | undefined,
  getHf?: () => AiHeaderFooterAccess | undefined,
  getSection?: () => AiSectionAccess | undefined,
  /** live predicate (gsk login && cloud-tools toggle); false hides generate_image */
  gskTools?: () => boolean,
): AgentSkill {
  // Selection frozen per run: tools act on the range the prompt described,
  // not on wherever the user's live selection has wandered mid-run. The doc
  // snapshot bounds the freeze's validity (see FrozenSelection).
  let frozen: FrozenSelection | null = null
  const cloudToolsOn = (): boolean => gskTools?.() ?? true
  return {
    id: 'docx',
    // live like tools: the loop re-reads both before every model request
    get systemPrompt() {
      return cloudToolsOn() ? AGENT_SYSTEM_PROMPT : AGENT_SYSTEM_PROMPT + GSK_TOOLS_OFF_NOTE
    },
    get tools() {
      return cloudToolsOn()
        ? AGENT_TOOLS
        : AGENT_TOOLS.filter((tool) => tool.name !== 'generate_image')
    },
    buildContext: () => {
      const editor = getEditor()
      markDocSeen(editor) // the context the model receives is the freshness baseline for index-addressed writes
      frozen = { scope: getSelectionScope(editor), doc: editor.state.doc }
      return buildDocContext(editor, frozen.scope, getComments?.()?.list(), getHf?.()?.read())
    },
    executeTool: (call, signal) =>
      executeTool(
        getEditor(),
        call,
        getNumIds(),
        getTrack?.(),
        signal,
        frozen,
        getComments?.(),
        getHf?.(),
        getSection?.(),
      ),
  }
}
