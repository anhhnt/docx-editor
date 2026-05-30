/**
 * PM paragraph → Document Paragraph conversion.
 *
 * Owns `extractParagraphContent` — the run-coalescing state machine that
 * walks each child node, dispatching to the run/hyperlink/field/sdt
 * factories and tracking the current run + current hyperlink so adjacent
 * text with the same mark set gets folded into a single Run. Tracked-change
 * marks (insertion/deletion/moveFrom/moveTo) split the run and emit their
 * own wrapper content. `createInlineSdtFromNode` lives here (not in
 * ./runs.ts) because it recurses back through this walker.
 */

import type { Node as PMNode } from 'prosemirror-model';
import type {
  Paragraph,
  Run,
  ParagraphFormatting,
  ParagraphContent,
  Hyperlink,
  NoteReferenceContent,
  InlineSdt,
  SdtProperties,
  TrackedChangeInfo,
} from '../../../types/document';
import type { ParagraphAttrs } from '../../schema/nodes';
import { getLinkKey, getMarksKey, marksToTextFormatting } from './marks';
import {
  createHyperlink,
  addNodeToHyperlink,
  createRunFromText,
  appendTextToRun,
  createBreakRun,
  createTabRun,
  createFieldFromNode,
  createMathFromNode,
  createImageRun,
  createShapeRun,
} from './runs';

/**
 * Convert a ProseMirror paragraph node to our Paragraph type
 */
export function convertPMParagraph(node: PMNode): Paragraph {
  const attrs = node.attrs as ParagraphAttrs;
  let content = insertCommentRanges(extractParagraphContent(node), node);

  // Emit BookmarkStart/End from bookmarks attr (for TOC anchors, cross-references)
  const bookmarks = attrs.bookmarks as Array<{ id: number; name: string }> | undefined;
  if (bookmarks && bookmarks.length > 0) {
    const starts: import('../../../types/content').ParagraphContent[] = bookmarks.map((b) => ({
      type: 'bookmarkStart' as const,
      id: b.id,
      name: b.name,
    }));
    const ends: import('../../../types/content').ParagraphContent[] = bookmarks.map((b) => ({
      type: 'bookmarkEnd' as const,
      id: b.id,
    }));
    content = [...starts, ...content, ...ends];
  }

  const paragraph: Paragraph = {
    type: 'paragraph',
    paraId: attrs.paraId || undefined,
    textId: attrs.textId || undefined,
    formatting: paragraphAttrsToFormatting(attrs),
    content,
  };

  // Preserve `<w:lastRenderedPageBreak/>` so a save+reload doesn't silently
  // drop the break Word recorded for paginating this paragraph.
  if (attrs.renderedPageBreakBefore) {
    paragraph.renderedPageBreakBefore = true;
  }

  // Round-trip paragraph-mark tracked-change attrs.
  if (attrs.pPrIns) {
    paragraph.pPrIns = {
      id: attrs.pPrIns.revisionId,
      author: attrs.pPrIns.author,
      ...(attrs.pPrIns.date ? { date: attrs.pPrIns.date } : {}),
    };
  }
  if (attrs.pPrDel) {
    paragraph.pPrDel = {
      id: attrs.pPrDel.revisionId,
      author: attrs.pPrDel.author,
      ...(attrs.pPrDel.date ? { date: attrs.pPrDel.date } : {}),
    };
  }
  // Round-trip paragraph-property-change history.
  if (attrs.pPrChange && attrs.pPrChange.length > 0) {
    paragraph.propertyChanges = attrs.pPrChange;
  }

  // Restore full section properties (round-trip) or fallback to break type only
  if (attrs._sectionProperties) {
    paragraph.sectionProperties =
      attrs._sectionProperties as import('../../../types/content').SectionProperties;
  } else if (attrs.sectionBreakType) {
    paragraph.sectionProperties = {
      sectionStart: attrs.sectionBreakType as import('../../../types/content').SectionStart,
    };
  }

  return paragraph;
}

/**
 * Scan paragraph PM node for comment marks and insert commentRangeStart/End
 * markers in the content array for round-trip serialization.
 */
function insertCommentRanges(content: ParagraphContent[], paragraph: PMNode): ParagraphContent[] {
  // Collect which comment IDs appear as marks on child nodes
  const commentIds = new Set<number>();
  paragraph.forEach((node) => {
    for (const mark of node.marks) {
      if (mark.type.name === 'comment') {
        commentIds.add(mark.attrs.commentId as number);
      }
    }
  });

  if (commentIds.size === 0) return content;

  // For each comment ID, find the first and last content item that belongs to it
  // and wrap with commentRangeStart/End
  const result: ParagraphContent[] = [];
  const openedComments = new Set<number>();
  let nodeIndex = 0;

  paragraph.forEach((node) => {
    const nodeCommentIds = new Set<number>();
    for (const mark of node.marks) {
      if (mark.type.name === 'comment') {
        nodeCommentIds.add(mark.attrs.commentId as number);
      }
    }

    // Close comments that are no longer active BEFORE pushing current content,
    // so commentRangeEnd lands after the last marked node, not after the first unmarked one
    for (const cid of [...openedComments]) {
      if (!nodeCommentIds.has(cid)) {
        result.push({ type: 'commentRangeEnd', id: cid });
        openedComments.delete(cid);
      }
    }

    // Open new comments
    for (const cid of nodeCommentIds) {
      if (!openedComments.has(cid)) {
        result.push({ type: 'commentRangeStart', id: cid });
        openedComments.add(cid);
      }
    }

    // Push the actual content item
    if (nodeIndex < content.length) {
      result.push(content[nodeIndex]);
    }

    nodeIndex++;
  });

  // Close any remaining open comments
  for (const cid of openedComments) {
    result.push({ type: 'commentRangeEnd', id: cid });
  }

  return result;
}

function paragraphAttrsToFormatting(attrs: ParagraphAttrs): ParagraphFormatting | undefined {
  // If we have the original inline formatting from the DOCX, use it as a base
  // for lossless round-trip. This preserves properties like contextualSpacing,
  // widowControl, beforeAutospacing, runProperties, etc. that aren't tracked
  // as individual PM attrs. It also avoids "inlining" style-inherited values
  // (spacing, indentation, numPr) which would override style definitions
  // and break rendering in Word/Pages/Google Docs.
  //
  // We then apply overrides for any properties the user may have changed
  // via editor commands (alignment, list toggle, etc.).
  if (attrs._originalFormatting) {
    const orig = attrs._originalFormatting;
    const result = { ...orig };

    // Override properties that user may have changed via editor commands.
    // Only override if the PM attr differs from the original value.
    if (attrs.alignment !== (orig.alignment || undefined)) {
      result.alignment = attrs.alignment || undefined;
    }
    if (attrs.numPr !== orig.numPr) {
      // Use JSON comparison since these are objects
      if (JSON.stringify(attrs.numPr) !== JSON.stringify(orig.numPr)) {
        result.numPr = attrs.numPr || undefined;
      }
    }
    if (attrs.styleId !== (orig.styleId || undefined)) {
      result.styleId = attrs.styleId || undefined;
    }
    if (attrs.pageBreakBefore !== (orig.pageBreakBefore || undefined)) {
      result.pageBreakBefore = attrs.pageBreakBefore || undefined;
    }
    if (attrs.bidi !== (orig.bidi || undefined)) {
      result.bidi = attrs.bidi || undefined;
    }

    return result;
  }

  // Fallback: reconstruct formatting from individual attrs (e.g. for
  // newly created paragraphs that don't have _originalFormatting)
  const hasFormatting =
    attrs.alignment ||
    attrs.spaceBefore ||
    attrs.spaceAfter ||
    attrs.lineSpacing ||
    attrs.indentLeft ||
    attrs.indentRight ||
    attrs.indentFirstLine ||
    attrs.numPr ||
    attrs.styleId ||
    attrs.borders ||
    attrs.shading ||
    attrs.tabs ||
    attrs.outlineLevel != null ||
    attrs.contextualSpacing ||
    attrs.bidi;

  if (!hasFormatting) {
    return undefined;
  }

  return {
    alignment: attrs.alignment || undefined,
    spaceBefore: attrs.spaceBefore || undefined,
    spaceAfter: attrs.spaceAfter || undefined,
    lineSpacing: attrs.lineSpacing || undefined,
    lineSpacingRule: attrs.lineSpacingRule || undefined,
    indentLeft: attrs.indentLeft || undefined,
    indentRight: attrs.indentRight || undefined,
    indentFirstLine: attrs.indentFirstLine || undefined,
    hangingIndent: attrs.hangingIndent || undefined,
    numPr: attrs.numPr || undefined,
    styleId: attrs.styleId || undefined,
    borders: attrs.borders || undefined,
    shading: attrs.shading || undefined,
    tabs: attrs.tabs || undefined,
    outlineLevel: attrs.outlineLevel ?? undefined,
    contextualSpacing: attrs.contextualSpacing || undefined,
    bidi: attrs.bidi || undefined,
  };
}

/**
 * Extract paragraph content (runs, hyperlinks) from ProseMirror paragraph
 *
 * Coalesces consecutive text with the same marks into single Runs
 * for efficient DOCX representation.
 */
function extractParagraphContent(paragraph: PMNode): ParagraphContent[] {
  const content: ParagraphContent[] = [];

  // Track current run being built
  let currentRun: Run | null = null;
  let currentMarksKey: string | null = null;
  let currentHyperlink: Hyperlink | null = null;

  paragraph.forEach((node) => {
    // Check for footnote/endnote reference mark
    const noteRefMark = node.marks.find((m) => m.type.name === 'footnoteRef');
    if (noteRefMark) {
      // Finish any current content
      if (currentRun) {
        content.push(currentRun);
        currentRun = null;
        currentMarksKey = null;
      }
      if (currentHyperlink) {
        content.push(currentHyperlink);
        currentHyperlink = null;
      }
      const noteType = noteRefMark.attrs.noteType === 'endnote' ? 'endnoteRef' : 'footnoteRef';
      const noteRef: NoteReferenceContent = {
        type: noteType,
        id: parseInt(noteRefMark.attrs.id, 10) || 0,
      };
      content.push({
        type: 'run',
        content: [noteRef],
      });
      return;
    }

    // Check for tracked change marks (insertion/deletion)
    const insertionMark = node.marks.find((m) => m.type.name === 'insertion');
    const deletionMark = node.marks.find((m) => m.type.name === 'deletion');
    if (insertionMark || deletionMark) {
      // Finish any current content
      if (currentRun) {
        content.push(currentRun);
        currentRun = null;
        currentMarksKey = null;
      }
      if (currentHyperlink) {
        content.push(currentHyperlink);
        currentHyperlink = null;
      }

      const changeMark = (insertionMark || deletionMark)!;
      // Build the run for whatever this node is — text, image, or shape — so a
      // tracked picture/shape round-trips inside the `<w:ins>`/`<w:del>` wrapper
      // instead of collapsing to an empty run (the prior text-only path).
      let run: Run;
      if (node.type.name === 'image') {
        run = createImageRun(node);
      } else if (node.type.name === 'shape') {
        run = createShapeRun(node);
      } else {
        // Filter out the tracked change mark for text formatting extraction
        const otherMarks = node.marks.filter(
          (m) => m.type.name !== 'insertion' && m.type.name !== 'deletion'
        );
        const formatting = marksToTextFormatting(otherMarks);
        run = {
          type: 'run',
          content: node.isText && node.text ? [{ type: 'text', text: node.text }] : [],
          ...(Object.keys(formatting).length > 0 ? { formatting } : {}),
        };
      }

      const info: TrackedChangeInfo = {
        id: changeMark.attrs.revisionId as number,
        author: (changeMark.attrs.author as string) || 'Unknown',
        date: (changeMark.attrs.date as string) || undefined,
      };
      // Only treat as a move pair when the parser explicitly flagged the
      // mark as having originated from `<w:moveFrom>`/`<w:moveTo>`. `w:id`
      // is not required to be unique per ECMA-376, so id coincidence
      // between an insertion and a deletion is not a reliable signal —
      // the suggestion-mode "replace" flow legitimately shares a date but
      // mints distinct ids precisely to dodge this trap.
      const isMovePair = changeMark.attrs.isMovePair === true;

      if (insertionMark) {
        if (isMovePair) {
          content.push({ type: 'moveTo', info, content: [run] });
        } else {
          content.push({ type: 'insertion', info, content: [run] });
        }
      } else {
        if (isMovePair) {
          content.push({ type: 'moveFrom', info, content: [run] });
        } else {
          content.push({ type: 'deletion', info, content: [run] });
        }
      }
      return;
    }

    // Check for hyperlink mark
    const linkMark = node.marks.find((m) => m.type.name === 'hyperlink');

    if (linkMark) {
      // Start or continue hyperlink
      const linkKey = getLinkKey(linkMark);

      const currentKey =
        currentHyperlink?.href || (currentHyperlink?.anchor ? `#${currentHyperlink.anchor}` : '');
      if (currentHyperlink && currentKey === linkKey) {
        // Continue current hyperlink
        addNodeToHyperlink(currentHyperlink, node);
      } else {
        // Finish previous content
        if (currentRun) {
          content.push(currentRun);
          currentRun = null;
          currentMarksKey = null;
        }
        if (currentHyperlink) {
          content.push(currentHyperlink);
        }

        // Start new hyperlink
        currentHyperlink = createHyperlink(linkMark);
        addNodeToHyperlink(currentHyperlink, node);
      }
      return;
    }

    // Not in hyperlink - finish any current hyperlink
    if (currentHyperlink) {
      content.push(currentHyperlink);
      currentHyperlink = null;
    }

    // Handle node types
    if (node.isText) {
      const marksKey = getMarksKey(node.marks);

      if (currentRun && currentMarksKey === marksKey) {
        // Append to current run
        appendTextToRun(currentRun, node.text || '');
      } else {
        // Start new run
        if (currentRun) {
          content.push(currentRun);
        }
        currentRun = createRunFromText(node.text || '', node.marks);
        currentMarksKey = marksKey;
      }
    } else if (node.type.name === 'hardBreak') {
      // Hard break ends current run
      if (currentRun) {
        content.push(currentRun);
        currentRun = null;
        currentMarksKey = null;
      }
      content.push(createBreakRun());
    } else if (node.type.name === 'image') {
      // Image ends current run
      if (currentRun) {
        content.push(currentRun);
        currentRun = null;
        currentMarksKey = null;
      }
      content.push(createImageRun(node));
    } else if (node.type.name === 'shape') {
      // Shape ends current run
      if (currentRun) {
        content.push(currentRun);
        currentRun = null;
        currentMarksKey = null;
      }
      content.push(createShapeRun(node));
    } else if (node.type.name === 'tab') {
      // Tab ends current run
      if (currentRun) {
        content.push(currentRun);
        currentRun = null;
        currentMarksKey = null;
      }
      content.push(createTabRun());
    } else if (node.type.name === 'field') {
      // Field ends current run and emits a field content item
      if (currentRun) {
        content.push(currentRun);
        currentRun = null;
        currentMarksKey = null;
      }
      content.push(createFieldFromNode(node, node.marks));
    } else if (node.type.name === 'sdt') {
      // SDT ends current run and emits an InlineSdt content item
      if (currentRun) {
        content.push(currentRun);
        currentRun = null;
        currentMarksKey = null;
      }
      content.push(createInlineSdtFromNode(node));
    } else if (node.type.name === 'math') {
      // Math ends current run and emits a MathEquation content item
      if (currentRun) {
        content.push(currentRun);
        currentRun = null;
        currentMarksKey = null;
      }
      content.push(createMathFromNode(node));
    }
  });

  // Don't forget the last run/hyperlink
  if (currentRun) {
    content.push(currentRun);
  }
  if (currentHyperlink) {
    content.push(currentHyperlink);
  }

  return content;
}

/**
 * Create an InlineSdt from a PM sdt node. Lives here (not in ./runs.ts)
 * because it recurses through `extractParagraphContent` — putting it in
 * runs.ts would create an import cycle.
 */
function createInlineSdtFromNode(node: PMNode): InlineSdt {
  const attrs = node.attrs as Record<string, unknown>;

  const properties: SdtProperties = {
    sdtType: (attrs.sdtType as SdtProperties['sdtType']) ?? 'richText',
    alias: (attrs.alias as string) ?? undefined,
    tag: (attrs.tag as string) ?? undefined,
    lock: (attrs.lock as SdtProperties['lock']) ?? undefined,
    placeholder: (attrs.placeholder as string) ?? undefined,
    showingPlaceholder: (attrs.showingPlaceholder as boolean) ?? undefined,
    dateFormat: (attrs.dateFormat as string) ?? undefined,
    listItems: attrs.listItems ? JSON.parse(attrs.listItems as string) : undefined,
    checked: attrs.checked != null ? (attrs.checked as boolean) : undefined,
  };

  // Extract content from the sdt node's children. OOXML allows runs,
  // hyperlinks, simple/complex fields, nested SDTs, and math here — keep
  // all of them so docProps-bound fields and similar template content
  // survive a round-trip through the editor.
  const sdtContent = extractParagraphContent(node);
  const content = sdtContent.filter(
    (c): c is InlineSdt['content'][number] =>
      c.type === 'run' ||
      c.type === 'hyperlink' ||
      c.type === 'simpleField' ||
      c.type === 'complexField' ||
      c.type === 'inlineSdt' ||
      c.type === 'mathEquation'
  );

  return {
    type: 'inlineSdt',
    properties,
    content,
  };
}
