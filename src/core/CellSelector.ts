import Element from "../components/Element";
import {
  DEFAULT_CELL_SELECTED_BACKGROUND_COLOR,
  DEFAULT_CELL_SELECTED_BORDER_COLOR,
  DEFAULT_CELL_SELECTED_FIXED_CELL_LINE_WIDTH,
} from "../config/index";
import drawBorder from "../utils/drawBorder";
import getRangeBorderInfo from "../utils/getRangeBorderInfo";

class CellSelector extends Element<null> {
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
  private checkStartCellInMergedRange(
    startCell: Excel.Cell.CellInstance | null,
    mergedCells: Excel.Sheet.CellRange[] | null
  ) {
    if (startCell && mergedCells) {
      for (const range of mergedCells) {
        if (
          startCell.rowIndex! >= range[0] &&
          startCell.rowIndex! <= range[1] &&
          startCell.colIndex! >= range[2] &&
          startCell.colIndex! <= range[3]
        ) {
          return true;
        }
      }
    }
    return false;
  }

  render(
    ctx: CanvasRenderingContext2D,
    selectedCells: Excel.Sheet.CellRange | null,
    startCell: Excel.Cell.CellInstance | null,
    scrollX: number,
    scrollY: number,
    mergedCells: Excel.Sheet.CellRange[] | null
  ) {
    if (selectedCells) {
      ctx.save();
      ctx.strokeStyle = DEFAULT_CELL_SELECTED_BORDER_COLOR;
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
        selectedCells,
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
          DEFAULT_CELL_SELECTED_BORDER_COLOR
        );
      }
      if (bottomBorderShow) {
        drawBorder(
          ctx,
          leftX,
          maxY - scrollY,
          rightX,
          maxY - scrollY,
          DEFAULT_CELL_SELECTED_BORDER_COLOR
        );
      }
      if (leftBorderShow) {
        drawBorder(
          ctx,
          minX - scrollX,
          topY,
          minX - scrollX,
          bottomY,
          DEFAULT_CELL_SELECTED_BORDER_COLOR
        );
      }
      if (rightBorderShow) {
        drawBorder(
          ctx,
          maxX - scrollX,
          topY,
          maxX - scrollX,
          bottomY,
          DEFAULT_CELL_SELECTED_BORDER_COLOR
        );
      }

      const w = rightX - leftX;
      const h = bottomY - topY;
      if (w > 0 && h > 0) {
        ctx.save();
        ctx.fillStyle = DEFAULT_CELL_SELECTED_BACKGROUND_COLOR;
        ctx.fillRect(leftX, topY, rightX - leftX, bottomY - topY);
        ctx.restore();
      }

      if (w > 0) {
        ctx.save();
        ctx.translate(0, -DEFAULT_CELL_SELECTED_FIXED_CELL_LINE_WIDTH / 2);
        drawBorder(
          ctx,
          leftX,
          this.fixedRowHeight,
          rightX,
          this.fixedRowHeight,
          DEFAULT_CELL_SELECTED_BORDER_COLOR,
          DEFAULT_CELL_SELECTED_FIXED_CELL_LINE_WIDTH
        );
        ctx.restore();
        ctx.save();
        ctx.fillStyle = DEFAULT_CELL_SELECTED_BACKGROUND_COLOR;
        ctx.fillRect(leftX, 0, rightX - leftX, this.fixedRowHeight);
        ctx.restore();
      }

      if (h > 0) {
        ctx.save();
        ctx.translate(-DEFAULT_CELL_SELECTED_FIXED_CELL_LINE_WIDTH / 2, 0);
        drawBorder(
          ctx,
          this.fixedColWidth,
          topY,
          this.fixedColWidth,
          bottomY,
          DEFAULT_CELL_SELECTED_BORDER_COLOR,
          DEFAULT_CELL_SELECTED_FIXED_CELL_LINE_WIDTH
        );
        ctx.restore();
        ctx.save();
        ctx.fillStyle = DEFAULT_CELL_SELECTED_BACKGROUND_COLOR;
        ctx.fillRect(0, topY, this.fixedColWidth, bottomY - topY);
        ctx.restore();
      }
      ctx.restore();
    }
    if (startCell) {
      if (this.checkStartCellInMergedRange(startCell, mergedCells)) {
        return;
      }
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
        [
          startCell.rowIndex!,
          startCell.rowIndex!,
          startCell.colIndex!,
          startCell.colIndex!,
        ],
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
          DEFAULT_CELL_SELECTED_BORDER_COLOR,
          2
        );
      }
      if (bottomBorderShow) {
        drawBorder(
          ctx,
          leftX,
          maxY - scrollY,
          rightX,
          maxY - scrollY,
          DEFAULT_CELL_SELECTED_BORDER_COLOR,
          2
        );
      }
      if (leftBorderShow) {
        drawBorder(
          ctx,
          minX - scrollX,
          topY,
          minX - scrollX,
          bottomY,
          DEFAULT_CELL_SELECTED_BORDER_COLOR,
          2
        );
      }
      if (rightBorderShow) {
        drawBorder(
          ctx,
          maxX - scrollX,
          topY,
          maxX - scrollX,
          bottomY,
          DEFAULT_CELL_SELECTED_BORDER_COLOR,
          2
        );
      }
    }
  }
}

export default CellSelector;
