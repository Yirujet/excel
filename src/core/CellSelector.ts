import Element from "../components/Element";
import {
  DEFAULT_CELL_SELECTED_BACKGROUND_COLOR,
  DEFAULT_CELL_SELECTED_COLOR,
  DEFAULT_CELL_SELECTED_FIXED_CELL_LINE_WIDTH,
} from "../config/index";

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

  getCellBorderInfo(
    range: Excel.Sheet.CellRange,
    scrollX: number,
    scrollY: number
  ) {
    const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] = range;
    const minX = this.cells[minRowIndex][minColIndex].position.leftTop.x!;
    const minY = this.cells[minRowIndex][minColIndex].position.leftTop.y!;
    const maxX = this.cells[maxRowIndex][maxColIndex].position.rightBottom.x!;
    const maxY = this.cells[maxRowIndex][maxColIndex].position.rightBottom.y!;
    const leftX = Math.max(minX - scrollX, this.fixedColWidth);
    const rightX = Math.min(maxX - scrollX, this.layout!.width);
    const topY = Math.max(minY - scrollY, this.fixedRowHeight);
    const bottomY = Math.min(maxY - scrollY, this.layout!.height);

    const topBorderShow =
      minY - scrollY >= this.fixedRowHeight &&
      minY - scrollY <= this.layout!.height &&
      leftX < rightX;
    const bottomBorderShow =
      maxY - scrollY <= this.layout!.height &&
      maxY - scrollY >= this.fixedRowHeight &&
      leftX < rightX;
    const leftBorderShow =
      minX - scrollX >= this.fixedColWidth &&
      minX - scrollX <= this.layout!.width &&
      topY < bottomY;
    const rightBorderShow =
      maxX - scrollX <= this.layout!.width &&
      maxX - scrollX >= this.fixedColWidth &&
      topY < bottomY;

    return {
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
    };
  }

  private drawBorder(
    ctx: CanvasRenderingContext2D,
    startX: number,
    startY: number,
    endX: number,
    endY: number,
    lineWidth: number = 1
  ) {
    ctx.save();
    ctx.strokeStyle = DEFAULT_CELL_SELECTED_COLOR;
    ctx.lineWidth = lineWidth;
    ctx.beginPath();
    ctx.moveTo(startX, startY);
    ctx.lineTo(endX, endY);
    ctx.closePath();
    ctx.stroke();
    ctx.restore();
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
      ctx.strokeStyle = DEFAULT_CELL_SELECTED_COLOR;
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
      } = this.getCellBorderInfo(selectedCells, scrollX, scrollY);

      if (topBorderShow) {
        this.drawBorder(ctx, leftX, minY - scrollY, rightX, minY - scrollY);
      }
      if (bottomBorderShow) {
        this.drawBorder(ctx, leftX, maxY - scrollY, rightX, maxY - scrollY);
      }
      if (leftBorderShow) {
        this.drawBorder(ctx, minX - scrollX, topY, minX - scrollX, bottomY);
      }
      if (rightBorderShow) {
        this.drawBorder(ctx, maxX - scrollX, topY, maxX - scrollX, bottomY);
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
        this.drawBorder(
          ctx,
          leftX,
          this.fixedRowHeight,
          rightX,
          this.fixedRowHeight,
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
        this.drawBorder(
          ctx,
          this.fixedColWidth,
          topY,
          this.fixedColWidth,
          bottomY,
          DEFAULT_CELL_SELECTED_FIXED_CELL_LINE_WIDTH
        );
        ctx.restore();
        ctx.save();
        ctx.fillStyle = DEFAULT_CELL_SELECTED_BACKGROUND_COLOR;
        ctx.fillRect(0, topY, this.fixedColWidth, bottomY - topY);
        ctx.restore();
      }
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
      } = this.getCellBorderInfo(
        [
          startCell.rowIndex!,
          startCell.rowIndex!,
          startCell.colIndex!,
          startCell.colIndex!,
        ],
        scrollX,
        scrollY
      );

      if (topBorderShow) {
        this.drawBorder(ctx, leftX, minY - scrollY, rightX, minY - scrollY, 2);
      }
      if (bottomBorderShow) {
        this.drawBorder(ctx, leftX, maxY - scrollY, rightX, maxY - scrollY, 2);
      }
      if (leftBorderShow) {
        this.drawBorder(ctx, minX - scrollX, topY, minX - scrollX, bottomY, 2);
      }
      if (rightBorderShow) {
        this.drawBorder(ctx, maxX - scrollX, topY, maxX - scrollX, bottomY, 2);
      }
    }
  }
}

export default CellSelector;
