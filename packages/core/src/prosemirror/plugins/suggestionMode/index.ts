/**
 * Suggestion Mode Plugin
 *
 * When active, intercepts all text insertions and deletions, wrapping
 * them in tracked-change marks (insertion/deletion) instead of modifying
 * the document directly.
 *
 * - Typed text is marked as insertion (green underline)
 * - Deleted text is NOT removed — it's marked as deletion (red strikethrough)
 * - Text already marked as insertion by the current author is deleted
 *   normally (retracting your own suggestion)
 *
 * The implementation is split across this directory for readability:
 *   - `state.ts`       — plugin key, meta constants, shared types
 *   - `markAttrs.ts`   — fresh-attr minting + projection
 *   - `adjacency.ts`   — coalescing lookups (sibling, cross-block, cellMarker)
 *   - `handlers/`      — keyboard / input handlers (delete, insert, structural)
 *   - `commands.ts`    — toggle / set / isActive
 *   - this file        — `createSuggestionModePlugin` + public re-exports
 */

import { isHistoryTransaction } from 'prosemirror-history';
import { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

import { findAdjacentRevisionForRange } from './adjacency';
import { applySuggestionInsert } from './handlers/insert';
import { handleSuggestionDelete } from './handlers/delete';
import {
  handleSuggestionBackspaceAtStart,
  handleSuggestionDeleteAtEnd,
  handleSuggestionEnter,
} from './handlers/structural';
import { makeMarkAttrs } from './markAttrs';
import {
  suggestionModeKey,
  SUGGESTION_BYPASS_META,
  SUGGESTION_META,
  type SuggestionModeState,
} from './state';

/**
 * Create the suggestion-mode ProseMirror plugin. **Must be mounted on
 * the editor view for `setSuggestionMode` and `toggleSuggestionMode`
 * to do anything** — both adapters (`@eigenpal/docx-editor-react`,
 * `@eigenpal/docx-editor-vue`) auto-mount this inside the `DocxEditor`
 * component, so consumers using the bundled components don't need to
 * register it themselves.
 *
 * When active, typed text gets the `insertion` mark, deleted text gets
 * the `deletion` mark (text stays in the doc; the painter strikes it
 * through), Enter sets `pPrIns` on the originating paragraph, and
 * Backspace at paragraph start sets `pPrDel` on the previous paragraph.
 * Author + adjacent same-author marks coalesce into one tracked change.
 *
 * @param initialActive - Whether suggesting mode starts on. Default `false`.
 * @param author - Author name attached to every minted revision. Default `'User'`.
 *
 * @example
 * ```ts
 * import { createSuggestionModePlugin } from '@eigenpal/docx-editor-core/prosemirror/plugins';
 *
 * const plugin = createSuggestionModePlugin(false, 'Jane');
 * EditorState.create({ doc, plugins: [plugin, ...other] });
 * ```
 */
export function createSuggestionModePlugin(initialActive = false, author = 'User'): Plugin {
  return new Plugin({
    key: suggestionModeKey,

    state: {
      init(): SuggestionModeState {
        return { active: initialActive, author };
      },
      apply(tr, state): SuggestionModeState {
        const meta = tr.getMeta(suggestionModeKey);
        if (meta) {
          return { ...state, ...meta };
        }
        return state;
      },
    },

    props: {
      handleDOMEvents: {
        // Intercept text input at the DOM level. ProseMirror's handleTextInput
        // is NOT reliably called when the hidden PM has complex mark structures
        // (it requires the change to span exactly one text node). By handling
        // beforeinput directly, we ensure suggestion mode always processes input.
        beforeinput(view: EditorView, event: InputEvent) {
          const pluginState = suggestionModeKey.getState(view.state);
          if (!pluginState?.active) return false;

          if (event.inputType === 'insertText' && event.data) {
            event.preventDefault();
            const { from, to } = view.state.selection;
            return applySuggestionInsert(view, from, to, event.data, pluginState);
          }

          return false;
        },
      },
      // Intercept Backspace and Delete to mark as deletion.
      // Enter splits the paragraph and marks the FIRST paragraph's pPrIns.
      handleKeyDown(view: EditorView, event: KeyboardEvent): boolean {
        const pluginState = suggestionModeKey.getState(view.state);
        if (!pluginState?.active) return false;

        if (event.key === 'Enter' && !event.shiftKey) {
          return handleSuggestionEnter(view.state, view.dispatch);
        }
        if (event.key === 'Backspace') {
          // At paragraph start (non-first paragraph), track the pilcrow
          // deletion instead of joining or deleting a character.
          if (handleSuggestionBackspaceAtStart(view.state, view.dispatch)) return true;
          return handleSuggestionDelete(view.state, view.dispatch, 'backward');
        }
        if (event.key === 'Delete') {
          if (handleSuggestionDeleteAtEnd(view.state, view.dispatch)) return true;
          return handleSuggestionDelete(view.state, view.dispatch, 'forward');
        }
        return false;
      },

      // Backup: also handle via PM's handleTextInput for simple cases
      handleTextInput(view: EditorView, from: number, to: number, text: string): boolean {
        const pluginState = suggestionModeKey.getState(view.state);
        if (!pluginState?.active) return false;
        return applySuggestionInsert(view, from, to, text, pluginState);
      },
    },

    // Catch-all: mark any unhandled new content (e.g. paste) as insertion
    appendTransaction(transactions, _oldState, newState) {
      const pluginState = suggestionModeKey.getState(newState);
      if (!pluginState?.active) return null;

      // Skip the catch-all mark-as-insertion path for:
      //   - transactions we've already authored (`SUGGESTION_META`)
      //   - accept/reject command transactions (`SUGGESTION_BYPASS_META`)
      //   - undo/redo (`isHistoryTransaction`)
      // The bypass meta is set by `resolveById` so structural-revision joins
      // (e.g. `pPrIns` reject → `tr.split` + `tr.setNodeMarkup`) aren't
      // re-wrapped as user insertions.
      // History transactions are skipped because a tracked edit and its marks
      // are recorded in one history event, so undo/redo already restores them;
      // re-running the catch-all over the replayed steps mis-stamps an insertion
      // on the boundary character. See eigenpal/docx-editor#633.
      const userTr = transactions.find(
        (tr) =>
          tr.docChanged &&
          !tr.getMeta(SUGGESTION_META) &&
          !tr.getMeta(SUGGESTION_BYPASS_META) &&
          !isHistoryTransaction(tr)
      );
      if (!userTr) return null;

      const insertionType = newState.schema.marks.insertion;
      if (!insertionType) return null;

      const tr = newState.tr;
      tr.setMeta(SUGGESTION_META, true);

      const deletionType = newState.schema.marks.deletion;
      userTr.steps.forEach((step) => {
        const stepMap = step.getMap();
        stepMap.forEach((_oldFrom, _oldTo, newFrom, newTo) => {
          if (newTo > newFrom) {
            // Reuse an adjacent same-author insertion when present so a
            // paste right after typed text coalesces into one tracked change.
            const markAttrs =
              findAdjacentRevisionForRange(
                newState.doc,
                newFrom,
                newTo,
                'insertion',
                pluginState.author
              ) ?? makeMarkAttrs(pluginState);
            // Mark text AND inline atoms (image, shape) that don't already
            // carry a tracked-change mark, so a pasted/dropped picture becomes
            // a tracked insertion just like typed text. Marking the whole range
            // would overwrite other authors' marks, so we go node by node and
            // skip any that the schema disallows marks on.
            newState.doc.nodesBetween(newFrom, newTo, (node, pos) => {
              // Text is the short-circuit (a leaf text node's own markSet is
              // empty, so `allowsMarkType` is false even though the paragraph
              // permits the mark); inline atoms (image, shape) go through the
              // `allowsMarkType` arm. Dropping the text short-circuit silently
              // stops tracking pasted text.
              if (!node.isText && !(node.isInline && node.type.allowsMarkType(insertionType))) {
                return;
              }
              const hasTrackedMark = node.marks.some(
                (m) => m.type === insertionType || (deletionType && m.type === deletionType)
              );
              if (!hasTrackedMark) {
                const nodeStart = Math.max(pos, newFrom);
                const nodeEnd = Math.min(pos + node.nodeSize, newTo);
                tr.addMark(nodeStart, nodeEnd, insertionType.create(markAttrs));
              }
            });
          }
        });
      });

      return tr.steps.length > 0 ? tr : null;
    },
  });
}

// Public surface — keep import paths stable for external consumers.
export { suggestionModeKey, SUGGESTION_BYPASS_META } from './state';
export { toggleSuggestionMode, setSuggestionMode, isSuggestionModeActive } from './commands';
