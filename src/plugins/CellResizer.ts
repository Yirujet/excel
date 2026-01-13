import Element from "../components/Element";
import {
  DEFAULT_RESIZER_LINE_COLOR,
  DEFAULT_RESIZER_LINE_DASH,
  DEFAULT_RESIZER_LINE_WIDTH,
} from "../config/index";
import HorizontalScrollbar from "../core/Scrollbar/HorizontalScrollbar";
import VerticalScrollbar from "../core/Scrollbar/VerticalScrollbar";
import drawBorder from "../utils/drawBorder";

class CellResizer extends Element<null> {
  declare cellResizer: CellResizer | null;
  // declare resizeInfo: Excel.Cell.CellAction["resize"];
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
  resizeInfo: Excel.Cell.CellAction["resize"] = {
    x: false,
    y: false,
    rowIndex: null,
    colIndex: null,
    value: null,
  };

  constructor(layout: Excel.LayoutInfo) {
    super("");
    this.layout = layout;
  }

  handleCellResize(resize: Excel.Cell.CellAction["resize"], isEnd = false) {
    if (resize.value) {
      this.cellResizer!.resizeInfo = resize;
    }
    if (isEnd) {
      if (this.cellResizer!.resizeInfo.x) {
        this.cells.forEach((row) => {
          row.forEach((cell, colIndex) => {
            if (colIndex === this.cellResizer!.resizeInfo.colIndex!) {
              cell.width = cell.width! + this.cellResizer!.resizeInfo.value!;
              cell.updatePosition();
            }
            if (colIndex > this.cellResizer!.resizeInfo.colIndex!) {
              cell.x = cell.x! + this.cellResizer!.resizeInfo.value!;
              cell.updatePosition();
            }
          });
        });
        this.layout!.bodyRealWidth += this.cellResizer!.resizeInfo.value!;
        this.realWidth += this.cellResizer!.resizeInfo.value!;
        this.horizontalScrollBar?.updateScrollbarInfo();
        this.horizontalScrollBar?.updatePosition();
      }
      if (this.cellResizer!.resizeInfo.y) {
        this.cells.forEach((row, rowIndex) => {
          row.forEach((cell) => {
            if (rowIndex === this.cellResizer!.resizeInfo.rowIndex!) {
              cell.height = cell.height! + this.cellResizer!.resizeInfo.value!;
              cell.updatePosition();
            }
            if (rowIndex > this.cellResizer!.resizeInfo.rowIndex!) {
              cell.y = cell.y! + this.cellResizer!.resizeInfo.value!;
              cell.updatePosition();
            }
          });
        });
        this.layout!.bodyRealHeight += this.cellResizer!.resizeInfo.value!;
        this.realHeight += this.cellResizer!.resizeInfo.value!;
        this.verticalScrollBar?.updateScrollbarInfo();
        this.verticalScrollBar?.updatePosition();
      }
      this.cellResizer!.resizeInfo = {
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
    scrollInfo: Excel.PositionPoint
  ) {
    if (this.resizeInfo.x) {
      drawBorder(
        ctx,
        Math.round(
          cellInfo.position.rightTop.x + this.resizeInfo.value! - scrollInfo.x
        ),
        0,
        Math.round(
          cellInfo.position.rightTop.x + this.resizeInfo.value! - scrollInfo.x
        ),
        this.layout.height,
        DEFAULT_RESIZER_LINE_COLOR,
        DEFAULT_RESIZER_LINE_WIDTH,
        DEFAULT_RESIZER_LINE_DASH
      );
      drawBorder(
        ctx,
        Math.round(cellInfo.position.leftTop.x - scrollInfo.x),
        0,
        Math.round(cellInfo.position.leftTop.x - scrollInfo.x),
        this.layout.height,
        DEFAULT_RESIZER_LINE_COLOR,
        DEFAULT_RESIZER_LINE_WIDTH,
        DEFAULT_RESIZER_LINE_DASH
      );
    }
    if (this.resizeInfo.y) {
      drawBorder(
        ctx,
        0,
        Math.round(
          cellInfo.position.leftBottom.y + this.resizeInfo.value! - scrollInfo.y
        ),
        this.layout.width,
        Math.round(
          cellInfo.position.leftBottom.y + this.resizeInfo.value! - scrollInfo.y
        ),
        DEFAULT_RESIZER_LINE_COLOR,
        DEFAULT_RESIZER_LINE_WIDTH,
        DEFAULT_RESIZER_LINE_DASH
      );
      drawBorder(
        ctx,
        0,
        Math.round(cellInfo.position.leftTop.y - scrollInfo.y),
        this.layout.width,
        Math.round(cellInfo.position.leftTop.y - scrollInfo.y),
        DEFAULT_RESIZER_LINE_COLOR,
        DEFAULT_RESIZER_LINE_WIDTH,
        DEFAULT_RESIZER_LINE_DASH
      );
    }
  }
}

export default CellResizer;
