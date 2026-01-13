import Element from "../components/Element";
import {
  DEFAULT_RESIZER_LINE_COLOR,
  DEFAULT_RESIZER_LINE_DASH,
  DEFAULT_RESIZER_LINE_WIDTH,
} from "../config/index";
import HorizontalScrollbar from "../core/Scrollbar/HorizontalScrollbar";
import VerticalScrollbar from "../core/Scrollbar/VerticalScrollbar";

class CellResizer extends Element<null> {
  declare cellResizer: CellResizer | null;
  declare resizeInfo: Excel.Cell.CellAction["resize"];
  declare cells: Excel.Cell.CellInstance[][];
  declare realWidth: number;
  declare realHeight: number;
  declare horizontalScrollBar: HorizontalScrollbar | null;
  declare verticalScrollBar: VerticalScrollbar | null;
  private declare pointInAbsFixedCell: (e: MouseEvent) => boolean;
  private declare pointInColResize: (e: MouseEvent) => boolean;
  private declare pointInRowResize: (e: MouseEvent) => boolean;
  private declare handleCellAction: (
    e: MouseEvent,
    isInX: boolean,
    isInY: boolean,
    triggerEvent: (
      resize: Excel.Cell.CellAction[Excel.Cell.Action],
      isEnd?: boolean
    ) => void
  ) => void;
  declare draw: () => void;
  layout: Excel.LayoutInfo;

  constructor(layout: Excel.LayoutInfo) {
    super("");
    this.layout = layout;
  }

  handleCellResize(resize: Excel.Cell.CellAction["resize"], isEnd = false) {
    if (resize.value) {
      this.resizeInfo = resize;
    }
    if (isEnd) {
      if (this.resizeInfo.x) {
        this.cells.forEach((row) => {
          row.forEach((cell, colIndex) => {
            if (colIndex === this.resizeInfo.colIndex!) {
              cell.width = cell.width! + this.resizeInfo.value!;
              cell.updatePosition();
            }
            if (colIndex > this.resizeInfo.colIndex!) {
              cell.x = cell.x! + this.resizeInfo.value!;
              cell.updatePosition();
            }
          });
        });
        this.layout!.bodyRealWidth += this.resizeInfo.value!;
        this.realWidth += this.resizeInfo.value!;
        this.horizontalScrollBar?.updateScrollbarInfo();
        this.horizontalScrollBar?.updatePosition();
      }
      if (this.resizeInfo.y) {
        this.cells.forEach((row, rowIndex) => {
          row.forEach((cell) => {
            if (rowIndex === this.resizeInfo.rowIndex!) {
              cell.height = cell.height! + this.resizeInfo.value!;
              cell.updatePosition();
            }
            if (rowIndex > this.resizeInfo.rowIndex!) {
              cell.y = cell.y! + this.resizeInfo.value!;
              cell.updatePosition();
            }
          });
        });
        this.layout!.bodyRealHeight += this.resizeInfo.value!;
        this.realHeight += this.resizeInfo.value!;
        this.verticalScrollBar?.updateScrollbarInfo();
        this.verticalScrollBar?.updatePosition();
      }
      this.resizeInfo = {
        x: false,
        y: false,
        rowIndex: null,
        colIndex: null,
        value: null,
      };
    }
    this.draw();
  }

  resize(e: MouseEvent) {
    if (!this.cellResizer) return;
    const isInAbsFixedCell = this.pointInAbsFixedCell(e);
    const isInColResize = this.pointInColResize(e);
    const isInRowResize = this.pointInRowResize(e);
    if (isInAbsFixedCell) return;
    this.handleCellAction(
      e,
      isInColResize,
      isInRowResize,
      this.cellResizer.handleCellResize.bind(this)
    );
  }

  render(
    ctx: CanvasRenderingContext2D,
    cellInfo: Excel.Cell.CellInstance,
    resizeInfo: Excel.Cell.CellAction["resize"],
    scrollInfo: Excel.PositionPoint
  ) {
    ctx.save();
    ctx.setLineDash(DEFAULT_RESIZER_LINE_DASH);
    ctx.lineWidth = DEFAULT_RESIZER_LINE_WIDTH;
    ctx.strokeStyle = DEFAULT_RESIZER_LINE_COLOR;
    if (resizeInfo.x) {
      ctx.beginPath();
      ctx.moveTo(
        Math.round(
          cellInfo.position.rightTop.x + resizeInfo.value! - scrollInfo.x
        ),
        0
      );
      ctx.lineTo(
        Math.round(
          cellInfo.position.rightTop.x + resizeInfo.value! - scrollInfo.x
        ),
        this.layout.height
      );
      ctx.closePath();
      ctx.stroke();
      ctx.beginPath();
      ctx.moveTo(Math.round(cellInfo.position.leftTop.x - scrollInfo.x), 0);
      ctx.lineTo(
        Math.round(cellInfo.position.leftTop.x - scrollInfo.x),
        this.layout.height
      );
      ctx.closePath();
      ctx.stroke();
    }
    if (resizeInfo.y) {
      ctx.beginPath();
      ctx.moveTo(
        0,
        Math.round(
          cellInfo.position.leftBottom.y + resizeInfo.value! - scrollInfo.y
        )
      );
      ctx.lineTo(
        this.layout.width,
        Math.round(
          cellInfo.position.leftBottom.y + resizeInfo.value! - scrollInfo.y
        )
      );
      ctx.closePath();
      ctx.stroke();
      ctx.beginPath();
      ctx.moveTo(0, Math.round(cellInfo.position.leftTop.y - scrollInfo.y));
      ctx.lineTo(
        this.layout.width,
        Math.round(cellInfo.position.leftTop.y - scrollInfo.y)
      );
      ctx.closePath();
      ctx.stroke();
    }
    ctx.restore();
  }
}

export default CellResizer;
