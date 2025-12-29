import Element from "../components/Element";
import {
  DEFAULT_CELL_DIAGONAL_LINE_COLOR,
  DEFAULT_CELL_DIAGONAL_LINE_WIDTH,
  DEFAULT_CELL_DIAGONAL_TEXT_COLOR,
  DEFAULT_CELL_DIAGONAL_TEXT_FONT_SIZE,
  DEFAULT_CELL_LINE_DASH,
  DEFAULT_CELL_PADDING,
} from "../config/index";
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
        let textList: string[] =
          cell.valueSlices!.length > 0
            ? cell.valueSlices!
            : [cell.value as string];
        textList.forEach((text, i) => {
          const textAlignOffsetX = cell.getTextAlignOffsetX(width);
          this.drawDataCellText(
            ctx,
            cell,
            text,
            textAlignOffsetX,
            x,
            y,
            width,
            height,
            cellWidth,
            cellHeight,
            scrollX,
            scrollY,
            cell.position.leftTop.y! +
              (height / (textList.length + 1)) * (i + 1) -
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
        });

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
      case "diagonal":
        this.drawDataCellDiagonal(
          ctx,
          cell,
          x,
          y,
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
    text: string,
    textAlignOffsetX: number,
    x: number,
    y: number,
    width: number,
    height: number,
    viewWidth: number,
    viewHeight: number,
    scrollX: number,
    scrollY: number,
    textY: number
  ) {
    ctx.save();
    const path = new Path2D();
    path.rect(x - scrollX, y - scrollY, viewWidth, viewHeight);
    ctx.clip(path);
    cell.setTextStyle(ctx);
    ctx.fillText(
      text,
      cell.position.leftTop.x! + textAlignOffsetX - scrollX,
      textY
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
      Math.round(
        cell.position.leftTop.x! + textAlignOffsetX - scrollX - underlineOffset
      ),
      Math.round(
        cell.position.leftTop.y! + height! / 2 - scrollY + wordHeight / 2
      )
    );
    ctx.lineTo(
      Math.round(
        cell.position.leftTop.x! +
          textAlignOffsetX -
          scrollX -
          underlineOffset +
          wordWidth
      ),
      Math.round(
        cell.position.leftTop.y! + height! / 2 - scrollY + wordHeight / 2
      )
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
    const path = new Path2D();
    path.rect(cellX - scrollX, cellY - scrollY, viewWidth, viewHeight);
    ctx.clip(path);
    ctx.drawImage(
      (cell.meta!.data as Excel.Cell.CellImageMetaData).img,
      x,
      y,
      width,
      height
    );
    ctx.restore();
  }

  drawDiagonalText(
    ctx: CanvasRenderingContext2D,
    prePoint: [number, number],
    curPoint: [number, number],
    text: string,
    cellX: number,
    cellY: number,
    scrollX: number,
    scrollY: number
  ) {
    const preAngle =
      Math.abs(prePoint[1] - cellY) / Math.abs(prePoint[0] - cellX);
    const curAngle =
      Math.abs(curPoint[1] - cellY) / Math.abs(curPoint[0] - cellX);
    const angle =
      Math.atan(preAngle) + (Math.atan(curAngle) - Math.atan(preAngle)) / 2;

    ctx.save();
    ctx.translate(cellX - scrollX, cellY - scrollY);
    ctx.rotate(angle);

    const midPoint = [
      (prePoint[0] + curPoint[0]) / 2,
      (prePoint[1] + curPoint[1]) / 2,
    ];
    const d = Math.sqrt(
      Math.pow(midPoint[0] - cellX, 2) + Math.pow(midPoint[1] - cellY, 2)
    );

    ctx.fillStyle = DEFAULT_CELL_DIAGONAL_TEXT_COLOR;
    ctx.font = `${DEFAULT_CELL_DIAGONAL_TEXT_FONT_SIZE}px sans-serif`;
    ctx.textBaseline = "middle";

    const textWidth = getTextMetrics(
      text,
      DEFAULT_CELL_DIAGONAL_TEXT_FONT_SIZE
    ).width;
    ctx.fillText(text, d / 2 - textWidth / 2, 0);

    ctx.restore();
  }

  drawDataCellDiagonal(
    ctx: CanvasRenderingContext2D,
    cell: Excel.Cell.CellInstance,
    cellX: number,
    cellY: number,
    cellWidth: number,
    cellHeight: number,
    scrollX: number,
    scrollY: number
  ) {
    const { direction, value } = cell.meta!
      .data as Excel.Cell.CellDiagonalMetaData;
    const times =
      value.length & 1 ? Math.floor(value.length / 2) : value.length / 2 - 1;

    let endPoints: [number, number][] = [];
    for (let i = 1; i <= times; i++) {
      endPoints.push([
        cellX + cellWidth,
        cellY + i * (cellHeight / (times + 1)),
      ]);
    }
    for (let i = times; i >= 1; i--) {
      endPoints.push([
        cellX + i * (cellWidth / (times + 1)),
        cellY + cellHeight,
      ]);
    }
    if (!(value.length & 1)) {
      endPoints.splice(times, 0, [cellX + cellWidth, cellY + cellHeight]);
    }

    ctx.save();
    ctx.strokeStyle = DEFAULT_CELL_DIAGONAL_LINE_COLOR;
    ctx.lineWidth = DEFAULT_CELL_DIAGONAL_LINE_WIDTH;
    ctx.beginPath();

    const startX = Math.round(cellX - scrollX);
    const startY = Math.round(cellY - scrollY);

    endPoints.forEach(([x, y]) => {
      ctx.moveTo(startX, startY);
      ctx.lineTo(Math.round(x - scrollX), Math.round(y - scrollY));
    });

    ctx.stroke();
    ctx.restore();

    endPoints.forEach(([x, y], i) => {
      const prePoint: [number, number] =
        i > 0 ? endPoints[i - 1] : [cellX + cellWidth, cellY];
      this.drawDiagonalText(
        ctx,
        prePoint,
        [x, y],
        value[i],
        cellX,
        cellY,
        scrollX,
        scrollY
      );
    });

    this.drawDiagonalText(
      ctx,
      endPoints[endPoints.length - 1],
      [cellX, cellY + cellHeight],
      value[value.length - 1],
      cellX,
      cellY,
      scrollX,
      scrollY
    );
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
        const showTopBorder =
          (minY - scrollY > this.fixedRowHeight && scrollY > 0) ||
          scrollY === 0;
        const showLeftBorder =
          (minX - scrollX > this.fixedColWidth && scrollX > 0) || scrollX === 0;
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

          if (leftTopCell.border.top && showTopBorder) {
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
            ctx.moveTo(Math.round(leftX), Math.round(topY));
            ctx.lineTo(Math.round(rightX), Math.round(topY));
            ctx.closePath();
            ctx.stroke();
            ctx.restore();
          }

          if (leftTopCell.border.right) {
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
            ctx.moveTo(Math.round(rightX), Math.round(topY));
            ctx.lineTo(Math.round(rightX), Math.round(bottomY));
            ctx.closePath();
            ctx.stroke();
            ctx.restore();
          }

          if (leftTopCell.border.bottom) {
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
            ctx.moveTo(Math.round(rightX), Math.round(bottomY));
            ctx.lineTo(Math.round(leftX), Math.round(bottomY));
            ctx.closePath();
            ctx.stroke();
            ctx.restore();
          }

          if (leftTopCell.border.left && showLeftBorder) {
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
            ctx.moveTo(Math.round(leftX), Math.round(topY));
            ctx.lineTo(Math.round(leftX), Math.round(bottomY));
            ctx.closePath();
            ctx.stroke();
            ctx.restore();
          }

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
