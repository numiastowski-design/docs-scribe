/** The row layout shared by the Word and PowerPoint control readings (`source-reading-table.ts#placeRowCells`): a cell starts at the
 *  next column no rowspan above covers, a gap standing for the covered columns, a range running down at most to the last row.
 *  Run: npx vitest run --project=main src/pipeline/source-reading-table.test.ts */
import { describe, expect, it } from "vitest";

import { placeRowCells, type CoveredRange, type LayoutCell } from "./source-reading-table";

const cell = (span: number, filled: boolean, rowspan = 1): LayoutCell => ({ span, filled, rowspan });

/** Lays a whole table out row by row, as the readings do. */
const layOut = (rows: LayoutCell[][]): { span: number; filled: boolean }[][] => {
  let covered: CoveredRange[] = [];
  return rows.map((cells, r) => {
    const row = placeRowCells(covered, cells, r, rows.length);
    covered = row.covered;
    return row.placed;
  });
};

describe("placeRowCells", () => {
  it("lays cells side by side when nothing is covered", () => {
    expect(layOut([[cell(2, true), cell(1, false)]])).toEqual([
      [
        { span: 2, filled: true },
        { span: 1, filled: false },
      ],
    ]);
  });

  it("stands an unfilled gap for the columns a rowspan above covers", () => {
    expect(layOut([[cell(2, true, 2), cell(1, true)], [cell(1, true)]])).toEqual([
      [
        { span: 2, filled: true },
        { span: 1, filled: true },
      ],
      [
        { span: 2, filled: false },
        { span: 1, filled: true },
      ],
    ]);
  });

  it("puts the gap between cells when the covered columns are in the middle", () => {
    expect(
      layOut([
        [cell(1, true), cell(1, true, 2), cell(1, true)],
        [cell(1, true), cell(1, true)],
      ])[1],
    ).toEqual([
      { span: 1, filled: true },
      { span: 1, filled: false },
      { span: 1, filled: true },
    ]);
  });

  it("stops covering after the last row the rowspan reaches", () => {
    expect(layOut([[cell(1, true, 2)], [cell(1, true)], [cell(1, true)]])[2]).toEqual([{ span: 1, filled: true }]);
  });

  it("caps a rowspan running past the last row at the last row", () => {
    const rows = layOut([[cell(1, true, 9)], [cell(1, true)], [cell(1, true)]]);
    expect(rows[1]).toEqual([
      { span: 1, filled: false },
      { span: 1, filled: true },
    ]);
    expect(rows[2]).toEqual([
      { span: 1, filled: false },
      { span: 1, filled: true },
    ]);
    // The cap itself (`rows[1]` and `rows[2]` are the same with or without it): the range ends at the last row.
    expect(placeRowCells([], [cell(1, true, 9)], 0, 3).covered[0]?.until).toBe(2);
  });

  it("a rowspan of 1 or less covers nothing", () => {
    expect(layOut([[cell(1, true, 1)], [cell(1, true)]])[1]).toEqual([{ span: 1, filled: true }]);
    expect(layOut([[cell(1, true, 0)], [cell(1, true)]])[1]).toEqual([{ span: 1, filled: true }]);
  });
});
