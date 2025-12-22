import Element from "../components/Element";
import {
  DEFAULT_RESIZER_LINE_COLOR,
  DEFAULT_RESIZER_LINE_DASH,
  DEFAULT_RESIZER_LINE_WIDTH,
} from "../config/index";

class CellResizer extends Element<null> {
  layout: Excel.LayoutInfo;
  constructor(layout: Excel.LayoutInfo) {
    super("");
    this.layout = layout;
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
