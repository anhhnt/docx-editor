/**
 * Regression — toolbar-created numbered lists must start at 1., not 2.
 *
 * Separate list runs after a plain paragraph must restart at 1. (Word uses a
 * fresh <w:num> + <w:startOverride w:val="1"/> per segment). A single shared
 * numId made the second list continue as 5., 6., …
 */

import { describe, test, expect } from 'bun:test';
import { schema } from '../../prosemirror/schema';
import { toFlowBlocks } from '../toFlowBlocks';
import { computeListMarker } from '../toFlowBlocks/listMarkers';

function numberedPara(text: string, numId: number, extra: Record<string, unknown> = {}) {
  return schema.nodes.paragraph.create(
    {
      numPr: { numId, ilvl: 0 },
      listIsBullet: false,
      listNumFmt: 'decimal',
      listMarker: null,
      ...extra,
    },
    schema.text(text)
  );
}

function markersOf(doc: ReturnType<typeof schema.nodes.doc.create>): string[] {
  return toFlowBlocks(doc, {})
    .filter((b) => b.kind === 'paragraph')
    .map((b) => b.attrs?.listMarker)
    .filter((m): m is string => !!m);
}

describe('numbered list starts at 1', () => {
  test('single list segment renders 1., 2., 3.', () => {
    const doc = schema.nodes.doc.create(null, [
      numberedPara('one', 2),
      numberedPara('two', 2),
      numberedPara('three', 2),
    ]);
    expect(markersOf(doc)).toEqual(['1.', '2.', '3.']);
  });

  test('second list restarts at 1 after a plain paragraph break', () => {
    const doc = schema.nodes.doc.create(null, [
      numberedPara('one', 2),
      numberedPara('two', 2),
      numberedPara('three', 2),
      numberedPara('four', 2),
      schema.nodes.paragraph.create({}, schema.text('this is a paragraph')),
      numberedPara('one', 3, { listStartOverride: 1 }),
      numberedPara('two', 3),
      numberedPara('three', 3),
      numberedPara('four', 3),
    ]);
    expect(markersOf(doc)).toEqual(['1.', '2.', '3.', '4.', '1.', '2.', '3.', '4.']);
  });

  test('toolbar listStartOverride=1 restarts without listAbstractNumId', () => {
    const counters = new Map<number, number[]>();
    const seen = new Set<string>();
    const marker = computeListMarker(
      {
        numPr: { numId: 5, ilvl: 0 },
        listIsBullet: false,
        listNumFmt: 'decimal',
        listStartOverride: 1,
      },
      counters,
      seen
    );
    expect(marker).toBe('1.');
  });

  test('orphan listStartOverride without listAbstractNumId is ignored', () => {
    const counters = new Map<number, number[]>();
    const seen = new Set<string>();
    const marker = computeListMarker(
      {
        numPr: { numId: 2, ilvl: 0 },
        listIsBullet: false,
        listNumFmt: 'decimal',
        listStartOverride: 2,
      },
      counters,
      seen
    );
    expect(marker).toBe('1.');
  });

  test('listStartOverride with listAbstractNumId still applies', () => {
    const counters = new Map<number, number[]>();
    const seen = new Set<string>();
    const marker = computeListMarker(
      {
        numPr: { numId: 2, ilvl: 0 },
        listIsBullet: false,
        listNumFmt: 'decimal',
        listMarker: '%1.',
        listLevelNumFmts: ['decimal'],
        listAbstractNumId: 4,
        listStartOverride: 5,
      },
      counters,
      seen
    );
    expect(marker).toBe('5.');
  });
});
