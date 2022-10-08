/* global console, document, Excel, Office */

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

export async function run() {
  try {
    await Excel.run(async (context) => {
      const depthToHeadingCells = await getHeadingCells(context);
      for (let headingCells of depthToHeadingCells) {
        for (let headingCell of headingCells) {
          console.log(headingCell.text[0][0]);
        }
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
