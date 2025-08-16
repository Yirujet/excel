import Element from "../components/Element";
import Sheet from "./Sheet";

class CellSelector extends Element {
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
    selectedCells: [number, number, number, number] | null,
    scrollX: number,
    scrollY: number
  ) {
    if (selectedCells) {
      ctx.save();
      ctx.strokeStyle = Sheet.DEFAULT_CELL_SELECTED_COLOR;
      const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
        selectedCells;
      const minX = this.cells[minRowIndex][minColIndex].position.leftTop.x!;
      const minY = this.cells[minRowIndex][minColIndex].position.leftTop.y!;
      const maxX = this.cells[maxRowIndex][maxColIndex].position.rightBottom.x!;
      const maxY = this.cells[maxRowIndex][maxColIndex].position.rightBottom.y!;

      const drawBorder = (
        startX: number,
        startY: number,
        endX: number,
        endY: number
      ) => {
        ctx.save();
        ctx.strokeStyle = Sheet.DEFAULT_CELL_SELECTED_COLOR;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(startX, startY);
        ctx.lineTo(endX, endY);
        ctx.closePath();
        ctx.stroke();
        ctx.restore();
      };

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

      if (topBorderShow) {
        drawBorder(leftX, minY - scrollY, rightX, minY - scrollY);
      }
      if (bottomBorderShow) {
        drawBorder(leftX, maxY - scrollY, rightX, maxY - scrollY);
      }
      if (leftBorderShow) {
        drawBorder(minX - scrollX, topY, minX - scrollX, bottomY);
      }
      if (rightBorderShow) {
        drawBorder(maxX - scrollX, topY, maxX - scrollX, bottomY);
      }

      const w = rightX - leftX;
      const h = bottomY - topY;
      if (w > 0 && h > 0) {
        ctx.save();
        ctx.fillStyle = Sheet.DEFAULT_CELL_SELECTED_BACKGROUND_COLOR;
        ctx.fillRect(leftX, topY, rightX - leftX, bottomY - topY);
        ctx.restore();
      }
    }
  }
}

export default CellSelector;
