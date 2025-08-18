import Element from "../components/Element";
import Sheet from "./Sheet";

class CellResizer extends Element {
  layout: Excel.LayoutInfo;
  fixedColWidth: number;
  fixedRowHeight: number;
  constructor(
    layout: Excel.LayoutInfo,
    fixedColWidth: number,
    fixedRowHeight: number
  ) {
    super("");
    this.layout = layout;
    this.fixedColWidth = fixedColWidth;
    this.fixedRowHeight = fixedRowHeight;
  }

  render(
    ctx: CanvasRenderingContext2D,
    cellInfo: Excel.Cell.CellInstance,
    resizeInfo: Excel.Cell.CellResize,
    scrollInfo: Excel.Sheet.ScrollInfo
  ) {
    ctx.save();
    ctx.setLineDash(Sheet.DEFAULT_RESIZER_LINE_DASH);
    ctx.lineWidth = Sheet.DEFAULT_RESIZER_LINE_WIDTH;
    ctx.strokeStyle = Sheet.DEFAULT_RESIZER_LINE_COLOR;
    if (resizeInfo.x) {
      ctx.beginPath();
      ctx.moveTo(
        cellInfo.position.rightTop.x + resizeInfo.value! - scrollInfo.x,
        this.fixedRowHeight
      );
      ctx.lineTo(
        cellInfo.position.rightTop.x + resizeInfo.value! - scrollInfo.x,
        this.fixedRowHeight + this.layout.height
      );
      ctx.closePath();
      ctx.stroke();
      ctx.beginPath();
      ctx.moveTo(
        cellInfo.position.leftTop.x - scrollInfo.x,
        this.fixedRowHeight
      );
      ctx.lineTo(
        cellInfo.position.leftTop.x - scrollInfo.x,
        this.fixedRowHeight + this.layout.height
      );
      ctx.closePath();
      ctx.stroke();
    }
    if (resizeInfo.y) {
      ctx.beginPath();
      ctx.moveTo(
        this.fixedColWidth,
        cellInfo.position.leftBottom.y + resizeInfo.value! - scrollInfo.y
      );
      ctx.lineTo(
        this.fixedColWidth + this.layout.width,
        cellInfo.position.leftBottom.y + resizeInfo.value! - scrollInfo.y
      );
      ctx.closePath();
      ctx.stroke();
      ctx.beginPath();
      ctx.moveTo(
        this.fixedColWidth,
        cellInfo.position.leftTop.y - scrollInfo.y
      );
      ctx.lineTo(
        this.fixedColWidth + this.layout.width,
        cellInfo.position.leftTop.y - scrollInfo.y
      );
      ctx.closePath();
      ctx.stroke();
    }
    ctx.restore();
  }
}

export default CellResizer;
