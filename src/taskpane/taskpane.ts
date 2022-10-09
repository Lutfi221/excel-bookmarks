/* global console, document, Excel, Office */

import { Mark } from "../types";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

/**
 *
 * @param context
 * @param depth Depth of heading to search to.
 *              For example, a depth of three means it will only search for subheadings
 *              at most three levels deep.
 * @returns An array of cells grouped by their depth.
 *          For example, ```output[1][4]``` points to the fifth second-hierarchical heading.
 */
async function getHeadingCells(context: Excel.RequestContext, depth = 3): Promise<Excel.Range[][]> {
  const workbook = context.workbook;
  const depthToHeadings: Excel.Range[][] = Array(depth).fill([]);

  workbook.worksheets.load("items");
  await context.sync();

  for (let i = 0; i < depth; i++) {
    const searchStr = "#".repeat(i + 1) + " ";

    for (let sheet of workbook.worksheets.items) {
      let rangeCollection: Excel.RangeCollection;
      rangeCollection = sheet.findAll(searchStr, { completeMatch: false }).areas;
      rangeCollection.load("items");
      try {
        // eslint-disable-next-line office-addins/no-context-sync-in-loop
        await context.sync();
      } catch (e) {
        if (e.code === "ItemNotFound") continue;
        console.error(e);
      }

      depthToHeadings[i] = rangeCollection.items.filter((item) => item.text[0][0].startsWith(searchStr));
    }
  }
  return depthToHeadings;
}

/**
 * Sorts marks so it can be displayed as a table
 * of contents.
 * @param marks Marks
 * @returns Sorted marks
 */
function sortMarks(marks: Mark[]): Mark[] {
  const sortedMarks: Mark[] = [];
  const depthToMarks: (Mark[] | undefined)[] = [];

  for (let mark of marks) {
    if (typeof depthToMarks[mark.order] === "undefined") {
      depthToMarks[mark.order] = [];
    }
    depthToMarks[mark.order].push(mark);
  }

  /**
   * Collapse children
   * @param mark Mark
   * @returns Collapsed mark and it's children
   */
  const collapseChildren = (mark: Mark): Mark[] => {
    const children = mark.getChildren(depthToMarks[mark.order + 1] || []);
    const output = [mark];
    for (let child of children) {
      output.push(...collapseChildren(child));
    }
    return output;
  };

  for (let mark of depthToMarks[0]) {
    sortedMarks.push(...collapseChildren(mark));
  }

  return sortedMarks;
}

async function createTableOfContests(
  marks: Mark[],
  range: Excel.Range,
  context: Excel.RequestContext,
  indentStr = "  "
) {
  range.load("rowCount");
  await context.sync();

  const cc = range.getCell(2, 1);
  cc.load(["text", "hyperlink"]);
  await context.sync();
  console.log(cc.hyperlink);
  console.log(cc.text);

  const rowCount = range.rowCount;
  let i = 0;

  for (let mark of marks) {
    if (i >= rowCount) {
      console.error(new Error("Make a bigger selection"));
      break;
    }

    const text = indentStr.repeat(mark.order) + mark.name;

    range.getCell(i, 0).set({ values: [[text]], hyperlink: { documentReference: mark.address, textToDisplay: text } });
    i++;
  }
}

export async function run() {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      workbook.worksheets.load("items");
      await context.sync();

      const worksheetNames = workbook.worksheets.items.map((worksheet) => worksheet.name);

      const depthToHeadingCells = await getHeadingCells(context);
      const marks: Mark[] = [];

      for (let headingCells of depthToHeadingCells) {
        for (let headingCell of headingCells) {
          marks.push(new Mark(headingCell, worksheetNames));
        }
      }

      for (let mark of marks) {
        mark.findParent(marks);
      }

      const sortedMarks = sortMarks(marks);
      await createTableOfContests(sortedMarks, workbook.getSelectedRange(), context);
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
