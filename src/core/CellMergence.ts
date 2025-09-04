import Element from "../components/Element";
import { DEFAULT_CELL_LINE_DASH } from "../config/index";

class CellMergence extends Element<null> {
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
    mergedCells: Excel.Sheet.CellRange[] | null,
    scrollX: number,
    scrollY: number
  ) {
    if (mergedCells) {
      mergedCells.forEach((e) => {
        const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] = e;
        const minX = this.cells[minRowIndex][minColIndex].position.leftTop.x!;
        const minY = this.cells[minRowIndex][minColIndex].position.leftTop.y!;
        const maxX =
          this.cells[maxRowIndex][maxColIndex].position.rightBottom.x!;
        const maxY =
          this.cells[maxRowIndex][maxColIndex].position.rightBottom.y!;
        const leftX = Math.max(minX - scrollX, this.fixedColWidth);
        const rightX = Math.min(maxX - scrollX, this.layout!.width);
        const topY = Math.max(minY - scrollY, this.fixedRowHeight);
        const bottomY = Math.min(maxY - scrollY, this.layout!.height);
        const leftTopCell = this.cells[minRowIndex][minColIndex];
        const rightBottomCell = this.cells[maxRowIndex][maxColIndex];
        const w =
          rightBottomCell.position.rightBottom.x! -
          leftTopCell.position.leftTop.x!;
        const h =
          rightBottomCell.position.rightBottom.y! -
          leftTopCell.position.leftTop.y!;
        if (leftX < rightX && topY < bottomY) {
          ctx.save();
          ctx.fillStyle = leftTopCell.textStyle.backgroundColor;
          ctx.beginPath();
          ctx.moveTo(leftX, topY);
          ctx.lineTo(rightX, topY);
          ctx.lineTo(rightX, bottomY);
          ctx.lineTo(leftX, bottomY);
          ctx.closePath();
          ctx.fill();
          ctx.restore();

          ctx.save();
          if (!leftTopCell.border.top.solid) {
            ctx.setLineDash(DEFAULT_CELL_LINE_DASH);
          } else {
            ctx.setLineDash([]);
          }
          ctx.strokeStyle = leftTopCell.border.top.color;
          if (leftTopCell.border.top.bold) {
            ctx.lineWidth = 2;
          } else {
            ctx.lineWidth = 1;
          }
          ctx.beginPath();
          ctx.moveTo(leftX, topY);
          ctx.lineTo(rightX, topY);
          ctx.closePath();
          ctx.stroke();
          ctx.restore();

          ctx.save();
          if (!leftTopCell.border.right.solid) {
            ctx.setLineDash(DEFAULT_CELL_LINE_DASH);
          } else {
            ctx.setLineDash([]);
          }
          ctx.strokeStyle = leftTopCell.border.right.color;
          if (leftTopCell.border.right.bold) {
            ctx.lineWidth = 2;
          } else {
            ctx.lineWidth = 1;
          }
          ctx.beginPath();
          ctx.moveTo(rightX, topY);
          ctx.lineTo(rightX, bottomY);
          ctx.closePath();
          ctx.stroke();
          ctx.restore();

          ctx.save();
          if (!leftTopCell.border.bottom.solid) {
            ctx.setLineDash(DEFAULT_CELL_LINE_DASH);
          } else {
            ctx.setLineDash([]);
          }
          ctx.strokeStyle = leftTopCell.border.bottom.color;
          if (leftTopCell.border.bottom.bold) {
            ctx.lineWidth = 2;
          } else {
            ctx.lineWidth = 1;
          }
          ctx.beginPath();
          ctx.moveTo(rightX, bottomY);
          ctx.lineTo(leftX, bottomY);
          ctx.closePath();
          ctx.stroke();
          ctx.restore();

          ctx.save();
          if (!leftTopCell.border.left.solid) {
            ctx.setLineDash(DEFAULT_CELL_LINE_DASH);
          } else {
            ctx.setLineDash([]);
          }
          ctx.strokeStyle = leftTopCell.border.left.color;
          if (leftTopCell.border.left.bold) {
            ctx.lineWidth = 2;
          } else {
            ctx.lineWidth = 1;
          }
          ctx.beginPath();
          ctx.moveTo(leftX, topY);
          ctx.lineTo(leftX, bottomY);
          ctx.closePath();
          ctx.stroke();
          ctx.restore();

          ctx.save();
          ctx.font = `${leftTopCell.textStyle.italic ? "italic" : ""} ${
            leftTopCell.textStyle.bold ? "bold" : "normal"
          } ${leftTopCell.textStyle.fontSize}px ${
            leftTopCell.textStyle.fontFamily
          }`;
          ctx.textBaseline = "middle";
          ctx.textAlign = leftTopCell.textStyle.align as CanvasTextAlign;
          ctx.fillStyle = leftTopCell.textStyle.color;
          const textAlignOffsetX = leftTopCell.getTextAlignOffsetX(w);
          ctx.fillText(
            leftTopCell.value,
            leftTopCell.x! + textAlignOffsetX - scrollX,
            leftTopCell.y! + h / 2 - scrollY
          );
          ctx.restore();
        }
      });
    }
  }
}

export default CellMergence;
