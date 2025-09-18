import Element from "../components/Element";
import {
  DEFAULT_CELL_FILLING_BORDER_COLOR,
  DEFAULT_CELL_FILLING_BORDER_WIDTH,
  DEFAULT_CELL_FILLING_LINE_DASH,
} from "../config/index";
import drawBorder from "../utils/drawBorder";
import getRangeBorderInfo from "../utils/getRangeBorderInfo";

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
    selectedCells: Excel.Sheet.CellRange | null,
    fillingCells: Excel.Sheet.CellRange | null,
    scrollX: number,
    scrollY: number
  ) {
    if (!fillingCells) {
      return;
    }
    ctx.save();
    ctx.strokeStyle = DEFAULT_CELL_FILLING_BORDER_COLOR;
    const {
      minX,
      minY,
      maxX,
      maxY,
      leftX,
      rightX,
      topY,
      bottomY,
      topBorderShow,
      bottomBorderShow,
      leftBorderShow,
      rightBorderShow,
    } = getRangeBorderInfo(
      fillingCells,
      scrollX,
      scrollY,
      this.layout,
      this.cells,
      this.fixedColWidth,
      this.fixedRowHeight
    );

    if (topBorderShow) {
      drawBorder(
        ctx,
        leftX,
        minY - scrollY,
        rightX,
        minY - scrollY,
        DEFAULT_CELL_FILLING_BORDER_COLOR,
        DEFAULT_CELL_FILLING_BORDER_WIDTH,
        DEFAULT_CELL_FILLING_LINE_DASH
      );
    }
    if (bottomBorderShow) {
      drawBorder(
        ctx,
        leftX,
        maxY - scrollY,
        rightX,
        maxY - scrollY,
        DEFAULT_CELL_FILLING_BORDER_COLOR,
        DEFAULT_CELL_FILLING_BORDER_WIDTH,
        DEFAULT_CELL_FILLING_LINE_DASH
      );
    }
    if (leftBorderShow) {
      drawBorder(
        ctx,
        minX - scrollX,
        topY,
        minX - scrollX,
        bottomY,
        DEFAULT_CELL_FILLING_BORDER_COLOR,
        DEFAULT_CELL_FILLING_BORDER_WIDTH,
        DEFAULT_CELL_FILLING_LINE_DASH
      );
    }
    if (rightBorderShow) {
      drawBorder(
        ctx,
        maxX - scrollX,
        topY,
        maxX - scrollX,
        bottomY,
        DEFAULT_CELL_FILLING_BORDER_COLOR,
        DEFAULT_CELL_FILLING_BORDER_WIDTH,
        DEFAULT_CELL_FILLING_LINE_DASH
      );
    }
    ctx.restore();
  }
}

export default Filling;
