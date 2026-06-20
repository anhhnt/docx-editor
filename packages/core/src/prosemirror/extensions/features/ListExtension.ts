/**
 * List Extension — list commands + keymaps
 *
 * No schema contribution — lists use paragraph attrs (numPr).
 * Provides: toggle bullet/number, indent/outdent, enter/backspace handling.
 */

import type { Command, EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { createExtension } from '../create';
import { Priority } from '../types';
import type { ExtensionRuntime } from '../types';
import { makeRevisionInfo } from '../../plugins/revisionIds';
import { SUGGESTION_META } from '../../plugins/suggestionMode/state';
import type { RevisionInfo } from '../../../types/content/trackedChange';

// ============================================================================
// CHAIN COMMANDS HELPER
// ============================================================================

function chainCommands(...commands: Command[]): Command {
  return (state, dispatch, view) => {
    for (const cmd of commands) {
      if (cmd(state, dispatch, view)) {
        return true;
      }
    }
    return false;
  };
}

// ============================================================================
// TRACKED PARAGRAPH-PROPERTY CHANGE (suggesting mode)
// ============================================================================

/**
 * Append a tracked paragraph-property change (`w:pPrChange`) recording the
 * paragraph's formatting before an edit, so Reject restores it. This is
 * exactly how Word encodes a paragraph-property change under track changes:
 * a `<w:pPrChange>` whose prior `<w:pPr>` holds the previous properties.
 *
 * `previousFormatting` carries only the fields the edit changed (reject merges
 * them back — see `applyPriorParagraphFormattingToAttrs`). Generic on purpose
 * so other suggesting-mode property edits (alignment, indent, …) can reuse it;
 * today only the list toggle does, capturing `{ numPr }` (issue #634).
 *
 * All paragraphs touched by one edit share `rev` (one id = one change), and
 * later Enter-splits copy the entry onto new paragraphs via PM's attr
 * inheritance — so rejecting reverts every affected paragraph.
 */
function appendParagraphPropertyChange(
  attrs: Record<string, unknown>,
  previousFormatting: Record<string, unknown>,
  rev: RevisionInfo
): Record<string, unknown> {
  const existing = Array.isArray(attrs.pPrChange) ? attrs.pPrChange : [];
  return {
    ...attrs,
    pPrChange: [
      ...existing,
      {
        type: 'paragraphPropertyChange',
        info: { id: rev.revisionId, author: rev.author, date: rev.date ?? undefined },
        previousFormatting,
      },
    ],
  };
}

// ============================================================================
// LIST COMMANDS
// ============================================================================

/** List-rendering attrs that must be reset when toggling lists via the toolbar. */
const CLEARED_LIST_RENDERING_ATTRS = {
  listStartOverride: null,
  listAbstractNumId: null,
  listLevelNumFmts: null,
  listMarkerHidden: null,
  listMarkerFontFamily: null,
  listMarkerFontSize: null,
  listMarkerSuffix: null,
  numPrFromStyle: null,
} as const;

function clearedListAttrs(): Record<string, unknown> {
  return {
    numPr: null,
    listIsBullet: null,
    listNumFmt: null,
    listMarker: null,
    ...CLEARED_LIST_RENDERING_ATTRS,
  };
}

type ListKind = 'bullet' | 'numbered';

function isBulletListItem(attrs: Record<string, unknown>): boolean {
  return !!(attrs.numPr as { numId?: number } | null)?.numId && attrs.listIsBullet === true;
}

function isNumberedListItem(attrs: Record<string, unknown>): boolean {
  return !!(attrs.numPr as { numId?: number } | null)?.numId && attrs.listIsBullet === false;
}

function isListItemOfKind(attrs: Record<string, unknown>, kind: ListKind): boolean {
  return kind === 'bullet' ? isBulletListItem(attrs) : isNumberedListItem(attrs);
}

function collectUsedNumIds(doc: PMNode): Set<number> {
  const used = new Set<number>();
  doc.descendants((node) => {
    if (node.type.name !== 'paragraph') return true;
    const numId = (node.attrs.numPr as { numId?: number } | null)?.numId;
    if (typeof numId === 'number' && numId > 0) used.add(numId);
    return true;
  });
  return used;
}

function allocateNumId(used: Set<number>): number {
  let id = 1;
  while (used.has(id)) id += 1;
  return id;
}

/** Immediate previous sibling when it is a paragraph; tables/other blocks break the chain. */
function getPreviousSiblingParagraph(doc: PMNode, pos: number): PMNode | null {
  const $pos = doc.resolve(pos);
  const index = $pos.index($pos.depth);
  if (index === 0) return null;
  const prev = $pos.parent.child(index - 1);
  return prev.type.name === 'paragraph' ? prev : null;
}

function continuesPriorList(doc: PMNode, pos: number, kind: ListKind): { numId: number } | null {
  const prev = getPreviousSiblingParagraph(doc, pos);
  if (!prev || !isListItemOfKind(prev.attrs as Record<string, unknown>, kind)) return null;
  const numId = (prev.attrs.numPr as { numId: number }).numId;
  return { numId };
}

function areAdjacentSiblings(prevPos: number, prevNode: PMNode, pos: number): boolean {
  return prevPos + prevNode.nodeSize === pos;
}

function listAttrsForSegment(
  kind: ListKind,
  numId: number,
  ilvl: number,
  startOverride: number | null = null
): Record<string, unknown> {
  const isBullet = kind === 'bullet';
  return {
    numPr: { numId, ilvl },
    listIsBullet: isBullet,
    listNumFmt: isBullet ? null : 'decimal',
    listMarker: null,
    ...CLEARED_LIST_RENDERING_ATTRS,
    listStartOverride: startOverride,
  };
}

function toggleListKind(kind: ListKind): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    const paragraph = $from.parent;
    if (paragraph.type.name !== 'paragraph') return false;

    const isInSameKindList = isListItemOfKind(paragraph.attrs as Record<string, unknown>, kind);

    if (!dispatch) return true;

    const paragraphs: { pos: number; node: PMNode }[] = [];
    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === 'paragraph') paragraphs.push({ pos, node });
    });
    paragraphs.sort((a, b) => a.pos - b.pos);

    let tr = state.tr;
    const touched = new Set<number>();

    // In suggesting mode, track the numbering change as a paragraph-property
    // revision so it can be reviewed and Reject reverts it. One toggle = one
    // revision id shared across every paragraph it touches (null when off).
    const rev = makeRevisionInfo(state);
    const usedNumIds = collectUsedNumIds(state.doc);
    let batchNumId: number | null = null;

    for (let i = 0; i < paragraphs.length; i += 1) {
      const { pos, node } = paragraphs[i];
      if (touched.has(pos)) continue;
      touched.add(pos);

      const priorNumPr = node.attrs.numPr;
      let nextAttrs: Record<string, unknown>;

      if (isInSameKindList) {
        nextAttrs = { ...node.attrs, ...clearedListAttrs() };
      } else {
        const ilvl = (node.attrs.numPr as { ilvl?: number } | null)?.ilvl ?? 0;
        let numId: number;
        let startOverride: number | null = null;

        const prevInBatch = i > 0 ? paragraphs[i - 1] : null;
        if (prevInBatch && areAdjacentSiblings(prevInBatch.pos, prevInBatch.node, pos)) {
          // Adjacent paragraphs toggled together share one new list instance.
          numId = batchNumId ?? allocateNumId(usedNumIds);
          if (batchNumId == null) {
            usedNumIds.add(numId);
            startOverride = 1;
          }
        } else {
          const continued = continuesPriorList(state.doc, pos, kind);
          if (continued) {
            // Extend an existing run (e.g. numbered item 2 → plain line 3).
            numId = continued.numId;
          } else {
            // New list after a break (plain paragraph, table, …). Word models
            // this as a fresh <w:num> with <w:startOverride w:val="1"/>.
            numId = allocateNumId(usedNumIds);
            usedNumIds.add(numId);
            startOverride = 1;
          }
        }

        batchNumId = numId;
        nextAttrs = { ...node.attrs, ...listAttrsForSegment(kind, numId, ilvl, startOverride) };
      }

      if (rev) {
        nextAttrs = appendParagraphPropertyChange(nextAttrs, { numPr: priorNumPr ?? null }, rev);
      }
      tr = tr.setNodeMarkup(pos, undefined, nextAttrs);
    }

    // Authored edit — keep the suggesting-mode catch-all from re-processing it.
    if (rev) tr.setMeta(SUGGESTION_META, true);

    dispatch(tr.scrollIntoView());
    return true;
  };
}

const toggleBulletList: Command = (state, dispatch) => {
  return toggleListKind('bullet')(state, dispatch);
};

const toggleNumberedList: Command = (state, dispatch) => {
  return toggleListKind('numbered')(state, dispatch);
};

const increaseListLevel: Command = (state, dispatch) => {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return false;
  if (!paragraph.attrs.numPr) return false;

  const currentLevel = paragraph.attrs.numPr.ilvl || 0;
  if (currentLevel >= 8) return false;

  if (!dispatch) return true;

  const paragraphPos = $from.before($from.depth);

  dispatch(
    state.tr
      .setNodeMarkup(paragraphPos, undefined, {
        ...paragraph.attrs,
        numPr: { ...paragraph.attrs.numPr, ilvl: currentLevel + 1 },
        // Clear explicit indentation so layout engine computes from new level
        indentLeft: null,
        indentFirstLine: null,
        hangingIndent: null,
      })
      .scrollIntoView()
  );

  return true;
};

const decreaseListLevel: Command = (state, dispatch) => {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return false;
  if (!paragraph.attrs.numPr) return false;

  const currentLevel = paragraph.attrs.numPr.ilvl || 0;

  if (!dispatch) return true;

  const paragraphPos = $from.before($from.depth);

  if (currentLevel <= 0) {
    dispatch(
      state.tr
        .setNodeMarkup(paragraphPos, undefined, {
          ...paragraph.attrs,
          ...clearedListAttrs(),
          indentLeft: null,
          indentFirstLine: null,
          hangingIndent: null,
        })
        .scrollIntoView()
    );
  } else {
    dispatch(
      state.tr
        .setNodeMarkup(paragraphPos, undefined, {
          ...paragraph.attrs,
          numPr: { ...paragraph.attrs.numPr, ilvl: currentLevel - 1 },
          indentLeft: null,
          indentFirstLine: null,
          hangingIndent: null,
        })
        .scrollIntoView()
    );
  }

  return true;
};

const removeList: Command = (state, dispatch) => {
  const { $from, $to } = state.selection;

  if (!dispatch) return true;

  let tr = state.tr;
  const seen = new Set<number>();

  state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
    if (node.type.name === 'paragraph' && node.attrs.numPr && !seen.has(pos)) {
      seen.add(pos);
      tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...clearedListAttrs() });
    }
  });

  dispatch(tr.scrollIntoView());
  return true;
};

// ============================================================================
// LIST QUERY HELPERS (exported for toolbar)
// ============================================================================

export function isInList(state: EditorState): boolean {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return false;
  return !!paragraph.attrs.numPr?.numId;
}

export function getListInfo(state: EditorState): { numId: number; ilvl: number } | null {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return null;
  if (!paragraph.attrs.numPr?.numId) return null;

  return {
    numId: paragraph.attrs.numPr.numId,
    ilvl: paragraph.attrs.numPr.ilvl || 0,
  };
}

// ============================================================================
// KEYMAP COMMANDS
// ============================================================================

function exitListOnEmptyEnter(): Command {
  return (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty) return false;

    const paragraph = $from.parent;
    if (paragraph.type.name !== 'paragraph') return false;

    const numPr = paragraph.attrs.numPr;
    if (!numPr) return false;

    if (paragraph.textContent.length > 0) return false;

    if (dispatch) {
      const tr = state.tr.setNodeMarkup($from.before(), undefined, {
        ...paragraph.attrs,
        ...clearedListAttrs(),
      });
      dispatch(tr);
    }
    return true;
  };
}

function splitListItem(): Command {
  return (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty) return false;

    const paragraph = $from.parent;
    if (paragraph.type.name !== 'paragraph') return false;

    const numPr = paragraph.attrs.numPr;
    if (!numPr) return false;

    if (dispatch) {
      const { tr } = state;
      const pos = $from.pos;

      tr.split(pos, 1, [{ type: state.schema.nodes.paragraph, attrs: { ...paragraph.attrs } }]);

      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

function backspaceExitList(): Command {
  return (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty) return false;

    if ($from.parentOffset !== 0) return false;

    const paragraph = $from.parent;
    if (paragraph.type.name !== 'paragraph') return false;

    const numPr = paragraph.attrs.numPr;
    if (!numPr) return false;

    if (dispatch) {
      const tr = state.tr.setNodeMarkup($from.before(), undefined, {
        ...paragraph.attrs,
        ...clearedListAttrs(),
      });
      dispatch(tr);
    }
    return true;
  };
}

function increaseListIndent(): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    // Collect all list paragraphs in the selection range
    const positions: { pos: number; attrs: Record<string, unknown> }[] = [];
    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === 'paragraph' && node.attrs.numPr) {
        const currentLevel = (node.attrs.numPr as { ilvl?: number }).ilvl ?? 0;
        if (currentLevel < 8) {
          positions.push({ pos, attrs: node.attrs as Record<string, unknown> });
        }
      }
    });

    if (positions.length === 0) return false;

    if (dispatch) {
      let tr = state.tr;
      for (const { pos, attrs } of positions) {
        const numPr = attrs.numPr as { ilvl?: number; numId?: number };
        tr = tr.setNodeMarkup(pos, undefined, {
          ...attrs,
          numPr: { ...numPr, ilvl: (numPr.ilvl ?? 0) + 1 },
          indentLeft: null,
          indentFirstLine: null,
          hangingIndent: null,
        });
      }
      dispatch(tr);
    }
    return true;
  };
}

function decreaseListIndent(): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    // Collect all list paragraphs in the selection range
    const positions: { pos: number; attrs: Record<string, unknown> }[] = [];
    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === 'paragraph' && node.attrs.numPr) {
        positions.push({ pos, attrs: node.attrs as Record<string, unknown> });
      }
    });

    if (positions.length === 0) return false;

    if (dispatch) {
      let tr = state.tr;
      for (const { pos, attrs } of positions) {
        const numPr = attrs.numPr as { ilvl?: number; numId?: number };
        const currentLevel = numPr.ilvl ?? 0;
        if (currentLevel <= 0) {
          tr = tr.setNodeMarkup(pos, undefined, {
            ...attrs,
            ...clearedListAttrs(),
            indentLeft: null,
            indentFirstLine: null,
            hangingIndent: null,
          });
        } else {
          tr = tr.setNodeMarkup(pos, undefined, {
            ...attrs,
            numPr: { ...numPr, ilvl: currentLevel - 1 },
            indentLeft: null,
            indentFirstLine: null,
            hangingIndent: null,
          });
        }
      }
      dispatch(tr);
    }
    return true;
  };
}

function insertTab(): Command {
  return (state, dispatch) => {
    const { schema } = state;
    const tabType = schema.nodes.tab;

    if (!tabType) {
      return false;
    }

    if (dispatch) {
      const tr = state.tr.replaceSelectionWith(tabType.create());
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

// Import goToNextCell/goToPrevCell from table extension for chaining
import { goToNextCell, goToPrevCell } from '../nodes/TableExtension';

// ============================================================================
// EXTENSION
// ============================================================================

export const ListExtension = createExtension({
  name: 'list',
  priority: Priority.High, // Must be before base keymap
  onSchemaReady(): ExtensionRuntime {
    return {
      commands: {
        toggleBulletList: () => toggleBulletList,
        toggleNumberedList: () => toggleNumberedList,
        increaseListLevel: () => increaseListLevel,
        decreaseListLevel: () => decreaseListLevel,
        removeList: () => removeList,
      },
      keyboardShortcuts: {
        Tab: chainCommands(goToNextCell(), increaseListIndent(), insertTab()),
        'Shift-Tab': chainCommands(goToPrevCell(), decreaseListIndent()),
        'Shift-Enter': () => false, // Let base keymap handle this
        Enter: chainCommands(exitListOnEmptyEnter(), splitListItem()),
        Backspace: backspaceExitList(),
      },
    };
  },
});
