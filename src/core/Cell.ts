import getTextMetrics from "../utils/getTextMetrics";
import Sheet from "./Sheet";
import Element from "../components/Element";
import debounce from "../utils/debounce";

class Cell extends Element implements Excel.Cell.CellInstance {
  width: number | null = null;
  height: number | null = null;
  rowIndex: number | null = null;
  colIndex: number | null = null;
  selected = false;
  cellName: string = "";
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
  textStyle = {
    fontFamily: "sans-serif",
    fontSize: 12,
    bold: false,
    italic: false,
    underline: false,
    backgroundColor: "",
    color: "#000",
    align: "left",
  };
  borderStyle = {
    solid: false,
    color: "#ccc",
    bold: false,
  };
  meta = null;
  value = "";
  fn = null;
  fixed = false;
  hidden = false;
  scrollX = 0;
  scrollY = 0;
  eventObserver: Excel.Event.ObserverInstance;

  constructor(eventObserver: Excel.Event.ObserverInstance) {
    super("", true);
    this.eventObserver = eventObserver;
    this.init();
  }

  init() {}

  initEvents() {
    const onMouseMove = debounce((e: MouseEvent) => {
      this.checkHit(e);
      if (this.fixed) return;
      if (!this.mouseEntered) return;
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

  checkHit(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    if (
      !(
        offsetX < this.position!.leftTop.x - Sheet.SCROLL_X ||
        offsetX > this.position!.rightTop.x - Sheet.SCROLL_X ||
        offsetY < this.position!.leftTop.y - Sheet.SCROLL_Y ||
        offsetY > this.position!.leftBottom.y - Sheet.SCROLL_Y
      )
    ) {
      console.log(this.value);
      this.mouseEntered = true;
      if (this.fixed) {
        this.cursor = "s-resize";
      } else {
        this.cursor = "cell";
      }
    } else {
      this.mouseEntered = false;
    }
  }

  setBorderStyle(ctx: CanvasRenderingContext2D) {
    if (!this.borderStyle.solid) {
      ctx.setLineDash(Sheet.DEFAULT_CELL_LINE_DASH);
    } else {
      ctx.setLineDash([]);
    }
    ctx.strokeStyle = this.borderStyle.color;
    if (this.borderStyle.bold) {
      ctx.lineWidth = 2;
    } else {
      ctx.lineWidth = 1;
    }
  }

  setTextStyle(ctx: CanvasRenderingContext2D) {
    ctx.font = `${this.textStyle.italic ? "italic" : ""} ${
      this.textStyle.bold ? "bold" : "normal"
    } ${this.textStyle.fontSize}px ${this.textStyle.fontFamily}`;
    ctx.textBaseline = "middle";
    ctx.textAlign = this.textStyle.align as CanvasTextAlign;
    ctx.fillStyle = this.textStyle.color;
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

  getTextAlignOffsetX(baseWidth: number) {
    if (this.textStyle.align === "left") {
      return 0;
    }
    if (this.textStyle.align === "center") {
      return baseWidth / 2;
    }
    return baseWidth;
  }

  render(
    ctx: CanvasRenderingContext2D,
    scrollX: number,
    scrollY: number,
    isEnd: boolean
  ) {
    if (isEnd) {
      this.clearEvents!(this.eventObserver, this);
      this.initEvents();
    }
    this.scrollX = scrollX;
    this.scrollY = scrollY;
    if (this.fixed) {
      this.drawFixedCell(ctx);
    } else {
      this.drawDataCell(ctx);
    }
    if (!this.hidden) {
      const textAlignOffsetX = this.getTextAlignOffsetX(this.width!);
      this.drawDataCellText(ctx, textAlignOffsetX);
      if (this.textStyle.underline) {
        this.drawDataCellUnderline(ctx, textAlignOffsetX);
      }
    }
  }

  drawFixedCell(ctx: CanvasRenderingContext2D) {
    ctx.save();
    this.setBorderStyle(ctx);
    ctx.strokeRect(
      this.x! - this.scrollX,
      this.y! - this.scrollY,
      this.width!,
      this.height!
    );
    ctx.restore();
  }

  drawDataCell(ctx: CanvasRenderingContext2D) {
    ctx.save();
    this.setBorderStyle(ctx);
    ctx.beginPath();
    ctx.moveTo(
      this.position.rightTop.x - this.scrollX,
      this.position.rightTop.y - this.scrollY
    );
    ctx.lineTo(
      this.position.rightBottom.x - this.scrollX,
      this.position.rightBottom.y - this.scrollY
    );
    ctx.closePath();
    ctx.stroke();
    ctx.restore();
    ctx.save();
    this.setBorderStyle(ctx);
    ctx.beginPath();
    ctx.moveTo(
      this.position.leftBottom.x - this.scrollX,
      this.position.leftBottom.y - this.scrollY
    );
    ctx.lineTo(
      this.position.rightBottom.x - this.scrollX,
      this.position.rightBottom.y - this.scrollY
    );
    ctx.closePath();
    ctx.stroke();
    ctx.restore();
  }

  drawDataCellText(ctx: CanvasRenderingContext2D, textAlignOffsetX: number) {
    ctx.save();
    this.setTextStyle(ctx);
    ctx.fillText(
      this.value,
      this.x! + textAlignOffsetX - this.scrollX,
      this.y! + this.height! / 2 - this.scrollY
    );
    ctx.restore();
  }

  drawDataCellUnderline(
    ctx: CanvasRenderingContext2D,
    textAlignOffsetX: number
  ) {
    const { width: wordWidth, height: wordHeight } = getTextMetrics(
      this.value,
      this.textStyle.fontSize
    );
    const underlineOffset = this.getTextAlignOffsetX(wordWidth);
    ctx.save();
    ctx.translate(0, 0.5);
    ctx.lineWidth = 0.5;
    ctx.strokeStyle = this.textStyle.color;
    ctx.beginPath();
    ctx.moveTo(
      this.x! + textAlignOffsetX - this.scrollX - underlineOffset,
      this.y! + this.height! / 2 - this.scrollY + wordHeight / 2
    );
    ctx.lineTo(
      this.x! + textAlignOffsetX - this.scrollX - underlineOffset + wordWidth,
      this.y! + this.height! / 2 - this.scrollY + wordHeight / 2
    );
    ctx.closePath();
    ctx.stroke();
    ctx.restore();
  }
}

export default Cell;
