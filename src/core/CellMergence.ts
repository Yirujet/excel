import Element from "../components/Element";
import { DEFAULT_CELL_LINE_DASH, DEFAULT_CELL_PADDING } from "../config/index";
import getImgDrawInfoByFillMode from "../utils/getImgDrawInfoByFillMode";
import getTextMetrics from "../utils/getTextMetrics";

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
  drawDataCell(
    ctx: CanvasRenderingContext2D,
    cell: Excel.Cell.CellInstance,
    x: number,
    y: number,
    width: number,
    height: number,
    cellWidth: number,
    cellHeight: number,
    scrollX: number,
    scrollY: number
  ) {
    switch (cell.meta?.type) {
      case "text":
        const textAlignOffsetX = cell.getTextAlignOffsetX(width);
        this.drawDataCellText(
          ctx,
          cell,
          textAlignOffsetX,
          x,
          y,
          width,
          height,
          scrollX,
          scrollY
        );
        if (cell.textStyle.underline) {
          this.drawDataCellUnderline(
            ctx,
            cell,
            textAlignOffsetX,
            x,
            y,
            width,
            height,
            scrollX,
            scrollY
          );
        }
        break;
      case "image":
        this.drawDataCellImage(
          ctx,
          cell,
          x,
          y,
          width,
          height,
          cellWidth,
          cellHeight,
          scrollX,
          scrollY
        );
        break;
    }
  }
  drawDataCellText(
    ctx: CanvasRenderingContext2D,
    cell: Excel.Cell.CellInstance,
    textAlignOffsetX: number,
    x: number,
    y: number,
    width: number,
    height: number,
    scrollX: number,
    scrollY: number
  ) {
    ctx.save();
    cell.setTextStyle(ctx);
    ctx.fillText(
      cell.value,
      cell.position.leftTop.x! + textAlignOffsetX - scrollX,
      cell.position.leftTop.y! + height! / 2 - scrollY
    );
    ctx.restore();
  }

  drawDataCellUnderline(
    ctx: CanvasRenderingContext2D,
    cell: Excel.Cell.CellInstance,
    textAlignOffsetX: number,
    x: number,
    y: number,
    width: number,
    height: number,
    scrollX: number,
    scrollY: number
  ) {
    const { width: wordWidth, height: wordHeight } = getTextMetrics(
      cell.value,
      cell.textStyle.fontSize
    );
    const underlineOffset = cell.getTextAlignOffsetX(wordWidth);
    ctx.save();
    ctx.translate(0, 0.5);
    ctx.lineWidth = 0.5;
    ctx.strokeStyle = cell.textStyle.color;
    ctx.beginPath();
    ctx.moveTo(
      cell.position.leftTop.x! + textAlignOffsetX - scrollX - underlineOffset,
      cell.position.leftTop.y! + height! / 2 - scrollY + wordHeight / 2
    );
    ctx.lineTo(
      cell.position.leftTop.x! +
        textAlignOffsetX -
        scrollX -
        underlineOffset +
        wordWidth,
      cell.position.leftTop.y! + height! / 2 - scrollY + wordHeight / 2
    );
    ctx.closePath();
    ctx.stroke();
    ctx.restore();
  }

  drawDataCellImage(
    ctx: CanvasRenderingContext2D,
    cell: Excel.Cell.CellInstance,
    cellX: number,
    cellY: number,
    cellWidth: number,
    cellHeight: number,
    viewWidth: number,
    viewHeight: number,
    scrollX: number,
    scrollY: number
  ) {
    const { x, y, width, height } = getImgDrawInfoByFillMode(
      cell.meta!.data as Excel.Cell.CellImageMetaData,
      {
        x: cell.position.leftTop.x! - scrollX + DEFAULT_CELL_PADDING,
        y: cell.position.leftTop.y! - scrollY + DEFAULT_CELL_PADDING,
        width: cellWidth - DEFAULT_CELL_PADDING * 2,
        height: cellHeight - DEFAULT_CELL_PADDING * 2,
      }
    )!;
    ctx.save();
    const range = new Path2D();
    range.rect(cellX - scrollX, cellY - scrollY, viewWidth, viewHeight);
    ctx.clip(range);
    ctx.drawImage(
      (cell.meta!.data as Excel.Cell.CellImageMetaData).img,
      x,
      y,
      width,
      height
    );
    ctx.restore();
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

          this.drawDataCell(
            ctx,
            leftTopCell,
            leftX + scrollX,
            topY + scrollY,
            w,
            h,
            rightX - leftX,
            bottomY - topY,
            scrollX,
            scrollY
          );
        }
      });
    }
  }
}

export default CellMergence;
