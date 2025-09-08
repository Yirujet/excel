import Element from "../components/Element";
import {
  DEFAULT_FILL_HANDLE_BACKGROUND_COLOR,
  DEFAULT_FILL_HANDLE_BORDER_COLOR,
  DEFAULT_FILL_HANDLE_BORDER_WIDTH,
  DEFAULT_FILL_HANDLE_SIZE,
} from "../config/index";
import debounce from "../utils/debounce";
import globalObj from "./globalObj";

class FillHandle
  extends Element<null>
  implements Excel.FillHandle.FillHandleInstance
{
  eventObserver: Excel.Event.ObserverInstance;
  layout: Excel.LayoutInfo;
  cells: Excel.Cell.CellInstance[][];
  fixedColWidth: number;
  fixedRowHeight: number;
  width: number | null = DEFAULT_FILL_HANDLE_SIZE;
  height: number | null = DEFAULT_FILL_HANDLE_SIZE;
  position: Excel.Position = {
    leftTop: {
      x: 0,
      y: 0,
    },
    rightTop: {
      x: 0,
      y: 0,
    },
    rightBottom: {
      x: 0,
      y: 0,
    },
    leftBottom: {
      x: 0,
      y: 0,
    },
  };
  constructor(
    eventObserver: Excel.Event.ObserverInstance,
    layout: Excel.LayoutInfo,
    cells: Excel.Cell.CellInstance[][],
    fixedColWidth: number,
    fixedRowHeight: number
  ) {
    super("");
    this.eventObserver = eventObserver;
    this.layout = layout;
    this.cells = cells;
    this.fixedColWidth = fixedColWidth;
    this.fixedRowHeight = fixedRowHeight;
    this.init();
  }

  init() {
    this.updatePosition();
    this.initEvents();
  }

  checkHit(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    const scrollX = globalObj.SCROLL_X;
    const scrollY = globalObj.SCROLL_Y;
    if (
      !(
        offsetX < this.position!.leftTop.x - scrollX ||
        offsetX > this.position!.rightTop.x - scrollX ||
        offsetY < this.position!.leftTop.y - scrollY ||
        offsetY > this.position!.leftBottom.y - scrollY
      )
    ) {
      globalObj.SET_CURSOR("crosshair");
    } else {
      globalObj.SET_CURSOR("default");
    }
  }

  initEvents() {
    const onMouseMove = debounce((e: MouseEvent) => {
      this.checkHit(e);
    }, 100);

    const defaultEventListeners = {
      mousemove: onMouseMove,
    };
    this.registerListenerFromOnProp(
      defaultEventListeners,
      this.eventObserver,
      this
    );
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
    ctx.strokeStyle = DEFAULT_FILL_HANDLE_BORDER_COLOR;
    ctx.lineWidth = lineWidth;
    ctx.beginPath();
    ctx.moveTo(startX, startY);
    ctx.lineTo(endX, endY);
    ctx.closePath();
    ctx.stroke();
    ctx.restore();
  }

  updatePosition() {
    this.position = {
      leftTop: {
        x: this.x!,
        y: this.y!,
      },
      rightTop: {
        x: this.x! + this.width!,
        y: this.y!,
      },
      rightBottom: {
        x: this.x! + this.width!,
        y: this.y! + this.height!,
      },
      leftBottom: {
        x: this.x!,
        y: this.y! + this.height!,
      },
    };
  }

  render(
    ctx: CanvasRenderingContext2D,
    selectedCells: Excel.Sheet.CellRange | null,
    scrollX: number,
    scrollY: number
  ) {
    if (selectedCells) {
      ctx.save();
      ctx.strokeStyle = DEFAULT_FILL_HANDLE_BORDER_COLOR;
      const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
        selectedCells;
      const minX =
        this.cells[maxRowIndex][maxColIndex].position.rightBottom.x! -
        DEFAULT_FILL_HANDLE_SIZE / 2 -
        DEFAULT_FILL_HANDLE_BORDER_WIDTH;
      const minY =
        this.cells[maxRowIndex][maxColIndex].position.rightBottom.y! -
        DEFAULT_FILL_HANDLE_SIZE / 2 -
        DEFAULT_FILL_HANDLE_BORDER_WIDTH;
      const maxX =
        this.cells[maxRowIndex][maxColIndex].position.rightBottom.x! +
        DEFAULT_FILL_HANDLE_SIZE / 2 +
        DEFAULT_FILL_HANDLE_BORDER_WIDTH;
      const maxY =
        this.cells[maxRowIndex][maxColIndex].position.rightBottom.y! +
        DEFAULT_FILL_HANDLE_SIZE / 2 +
        DEFAULT_FILL_HANDLE_BORDER_WIDTH;
      this.x = minX;
      this.y = minY;
      this.width = maxX - minX;
      this.height = maxY - minY;
      this.updatePosition();
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
        this.drawBorder(
          ctx,
          leftX,
          minY - scrollY,
          rightX,
          minY - scrollY,
          DEFAULT_FILL_HANDLE_BORDER_WIDTH
        );
      }
      if (bottomBorderShow) {
        this.drawBorder(
          ctx,
          leftX,
          maxY - scrollY,
          rightX,
          maxY - scrollY,
          DEFAULT_FILL_HANDLE_BORDER_WIDTH
        );
      }
      if (leftBorderShow) {
        this.drawBorder(
          ctx,
          minX - scrollX,
          topY,
          minX - scrollX,
          bottomY,
          DEFAULT_FILL_HANDLE_BORDER_WIDTH
        );
      }
      if (rightBorderShow) {
        this.drawBorder(
          ctx,
          maxX - scrollX,
          topY,
          maxX - scrollX,
          bottomY,
          DEFAULT_FILL_HANDLE_BORDER_WIDTH
        );
      }

      const w = rightX - leftX;
      const h = bottomY - topY;
      if (w > 0 && h > 0) {
        ctx.save();
        ctx.fillStyle = DEFAULT_FILL_HANDLE_BACKGROUND_COLOR;
        ctx.fillRect(leftX, topY, rightX - leftX, bottomY - topY);
        ctx.restore();
      }
    }
  }
}

export default FillHandle;
