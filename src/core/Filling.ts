import Element from "../components/Element";

class Filling extends Element<null> {
  layout: Excel.LayoutInfo;
  cells: Excel.Cell.CellInstance[][];
  fixedColWidth: number;
  fixedRowHeight: number;
  constructor(
    layout: Excel.LayoutInfo,
    cells: Excel.Cell.CellInstance[][],
    fixedColWidth: number,
    fixedRowHeight: number
  ) {
    super("");
    this.layout = layout;
    this.cells = cells;
    this.fixedColWidth = fixedColWidth;
    this.fixedRowHeight = fixedRowHeight;
  }

  render(
    ctx: CanvasRenderingContext2D,
    fillingCells: Excel.Sheet.CellRange | null,
    scrollX: number,
    scrollY: number
  ) {}
}

export default Filling;
