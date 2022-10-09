/* global Excel*/

export class Mark {
  /**
   * A1-style address to the heading cell.
   */
  address: string;
  /**
   * Name or title of the heading.
   */
  name: string;
  /**
   * Zero-based order of the heading.
   */
  order: number;
  /**
   * Parent heading.
   */
  parentMark?: Mark;
  /**
   * Parental distance.
   */
  parentalDistance?: number;
  /**
   * Zero-based coordinates of the heading cell.
   */
  coordinates: [number, number];
  /**
   * Worksheet name containing the heading cell
   */
  worksheetName: string;
  headingCell: Excel.Range;
  constructor(headingCell: Excel.Range, worksheetName?: string) {
    this.address = headingCell.address;
    this.name = headingCell.text[0][0].replace(/^#+ /, "");

    // Counts the number of "#" characters at the start, then subtract 1.
    this.order = headingCell.text[0][0].split(" ")[0].length - 1;
    this.coordinates = [headingCell.columnIndex, headingCell.rowIndex];

    if (worksheetName) {
      this.worksheetName = worksheetName;
    } else {
      this.worksheetName = headingCell.worksheet.name;
    }

    this.headingCell = headingCell;
  }

  /**
   * Finds the most likely parent, and updates
   * ```parentMark``` and ```parentalDistance``` property.
   * @param pool Array of potential parents
   * @returns The parent mark or null
   */
  public findParent(pool: Mark[]): Mark | null {
    if (this.order === 0) return null;
    const parentOrder = this.order - 1;
    const filteredPool = pool.filter((mark) => mark.order === parentOrder);

    let closestParent: Mark | null = null;
    let distanceToClosestParent = Infinity;

    for (let mark of filteredPool) {
      const d = this.parentalDistanceTo(mark);
      if (d < distanceToClosestParent) {
        closestParent = mark;
        distanceToClosestParent = d;
      }
    }

    this.parentMark = closestParent;
    this.parentalDistance = distanceToClosestParent;
    return closestParent;
  }

  /**
   * Calculates the parental distance
   * @param mark Potential parent
   * @returns Parental distance (the lower it is, the more likely
   *          that is the potential parent)
   */
  private parentalDistanceTo(mark: Mark): number {
    if (mark.worksheetName !== this.worksheetName) return Infinity;
    let offset = 0;
    const deltaRows = this.coordinates[1] - mark.coordinates[1];
    const deltaColumns = this.coordinates[0] - mark.coordinates[0];

    // If the potential parent mark is to the bottom.
    if (deltaRows < 0) {
      offset += 10000;
    }
    // If the potential parent mark is to the right.
    if (deltaColumns < 0) {
      offset += 10000;
    }

    if (deltaColumns > 0) {
      return this.coordinates[1] * deltaColumns + mark.coordinates[1] + offset;
    }
    return deltaRows + offset;
  }
}
