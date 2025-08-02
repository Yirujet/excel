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
  border: Excel.Cell.Border = {
    top: {
      solid: false,
      color: "rgb(230, 230, 230)",
      bold: false,
    },
    bottom: {
      solid: false,
      color: "rgb(230, 230, 230)",
      bold: false,
    },
    left: {
      solid: false,
      color: "rgb(230, 230, 230)",
      bold: false,
    },
    right: {
      solid: false,
      color: "rgb(230, 230, 230)",
      bold: false,
    },
  };
  meta = null;
  value = "";
  fn = null;
  fixed = {
    x: false,
    y: false,
  };
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
    const scrollX =
      this.fixed.x || this.fixed.y ? this.scrollX : Sheet.SCROLL_X;
    const scrollY =
      this.fixed.x || this.fixed.y ? this.scrollY : Sheet.SCROLL_Y;
    if (
      !(
        offsetX < this.position!.leftTop.x - scrollX ||
        offsetX > this.position!.rightTop.x - scrollX ||
        offsetY < this.position!.leftTop.y - scrollY ||
        offsetY > this.position!.leftBottom.y - scrollY
      )
    ) {
      this.mouseEntered = true;
      // console.log(this.value);
      if (this.fixed.y) {
        Sheet.SET_CURSOR("s-resize");
      } else if (this.fixed.x) {
        Sheet.SET_CURSOR("w-resize");
      } else {
        Sheet.SET_CURSOR("cell");
      }
    } else {
      this.mouseEntered = false;
    }
  }

  setBorderStyle(ctx: CanvasRenderingContext2D, side: Excel.Cell.BorderSide) {
    if (!this.border[side].solid) {
      ctx.setLineDash(Sheet.DEFAULT_CELL_LINE_DASH);
    } else {
      ctx.setLineDash([]);
    }
    ctx.strokeStyle = this.border[side].color;
    if (this.border[side].bold) {
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
    this.drawCellBg(ctx);
    this.drawCellBorder(ctx);
    if (!this.hidden) {
      const textAlignOffsetX = this.getTextAlignOffsetX(this.width!);
      this.drawDataCellText(ctx, textAlignOffsetX);
      if (this.textStyle.underline) {
        this.drawDataCellUnderline(ctx, textAlignOffsetX);
      }
    }
  }

  drawCellBg(ctx: CanvasRenderingContext2D) {
    if (this.textStyle.backgroundColor) {
      ctx.fillStyle = this.textStyle.backgroundColor;
      ctx.fillRect(
        this.x! - this.scrollX,
        this.y! - this.scrollY,
        this.width!,
        this.height!
      );
    }
  }

  drawCellBorder(ctx: CanvasRenderingContext2D) {
    if (this.fixed.x || this.fixed.y) {
      ctx.save();
      this.setBorderStyle(ctx, "top");
      ctx.beginPath();
      ctx.moveTo(
        this.position.leftTop.x - this.scrollX,
        this.position.leftTop.y - this.scrollY
      );
      ctx.lineTo(
        this.position.rightTop.x - this.scrollX,
        this.position.rightTop.y - this.scrollY
      );
      ctx.closePath();
      ctx.stroke();
      ctx.restore();
      ctx.save();
      this.setBorderStyle(ctx, "left");
      ctx.beginPath();
      ctx.moveTo(
        this.position.leftBottom.x - this.scrollX,
        this.position.leftBottom.y - this.scrollY
      );
      ctx.lineTo(
        this.position.leftTop.x - this.scrollX,
        this.position.leftTop.y - this.scrollY
      );
      ctx.closePath();
      ctx.stroke();
      ctx.restore();
    }
    ctx.save();
    this.setBorderStyle(ctx, "right");
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
    this.setBorderStyle(ctx, "bottom");
    ctx.beginPath();
    ctx.moveTo(
      this.position.rightBottom.x - this.scrollX,
      this.position.rightBottom.y - this.scrollY
    );
    ctx.lineTo(
      this.position.leftBottom.x - this.scrollX,
      this.position.leftBottom.y - this.scrollY
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
