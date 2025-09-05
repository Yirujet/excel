import getTextMetrics from "../utils/getTextMetrics";
import Element from "../components/Element";
import debounce from "../utils/debounce";
import throttle from "../utils/throttle";
import {
  DEFAULT_CELL_LINE_BOLD,
  DEFAULT_CELL_LINE_COLOR,
  DEFAULT_CELL_LINE_DASH,
  DEFAULT_CELL_LINE_SOLID,
  DEFAULT_CELL_TEXT_ALIGN,
  DEFAULT_CELL_TEXT_BACKGROUND_COLOR,
  DEFAULT_CELL_TEXT_BOLD,
  DEFAULT_CELL_TEXT_COLOR,
  DEFAULT_CELL_TEXT_FONT_FAMILY,
  DEFAULT_CELL_TEXT_FONT_SIZE,
  DEFAULT_CELL_TEXT_ITALIC,
  DEFAULT_CELL_TEXT_UNDERLINE,
  RESIZE_COL_SIZE,
  RESIZE_ROW_SIZE,
} from "../config/index";
import globalObj from "./globalObj";

class Cell extends Element<null> implements Excel.Cell.CellInstance {
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
    fontFamily: DEFAULT_CELL_TEXT_FONT_FAMILY,
    fontSize: DEFAULT_CELL_TEXT_FONT_SIZE,
    bold: DEFAULT_CELL_TEXT_BOLD,
    italic: DEFAULT_CELL_TEXT_ITALIC,
    underline: DEFAULT_CELL_TEXT_UNDERLINE,
    backgroundColor: DEFAULT_CELL_TEXT_BACKGROUND_COLOR,
    color: DEFAULT_CELL_TEXT_COLOR,
    align: DEFAULT_CELL_TEXT_ALIGN,
  };
  border: Excel.Cell.Border = {
    top: {
      solid: DEFAULT_CELL_LINE_SOLID,
      color: DEFAULT_CELL_LINE_COLOR,
      bold: DEFAULT_CELL_LINE_BOLD,
    },
    bottom: {
      solid: DEFAULT_CELL_LINE_SOLID,
      color: DEFAULT_CELL_LINE_COLOR,
      bold: DEFAULT_CELL_LINE_BOLD,
    },
    left: {
      solid: DEFAULT_CELL_LINE_SOLID,
      color: DEFAULT_CELL_LINE_COLOR,
      bold: DEFAULT_CELL_LINE_BOLD,
    },
    right: {
      solid: DEFAULT_CELL_LINE_SOLID,
      color: DEFAULT_CELL_LINE_COLOR,
      bold: DEFAULT_CELL_LINE_BOLD,
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
  action: Excel.Cell.CellAction = {
    select: {
      x: false,
      y: false,
      rowIndex: null,
      colIndex: null,
      value: null,
    },
    resize: {
      x: false,
      y: false,
      rowIndex: null,
      colIndex: null,
      value: null,
    },
  };
  moveStartValue = 0;
  actionEvent: Excel.Event.FnType | null = null;
  acting = false;

  constructor(eventObserver: Excel.Event.ObserverInstance) {
    super("", true);
    this.eventObserver = eventObserver;
    this.init();
  }

  init() {}

  initEvents() {
    const onMouseMove = debounce((e: MouseEvent) => {
      if (this.acting) return;
      this.checkHit(e);
      this.checkResizeOrSelect(e);
    }, 100);

    const handleStartAction = (e: MouseEvent, type: Excel.Cell.Action) => {
      this.acting = true;
      const offsetProp = this.action[type].x ? "x" : "y";
      this.moveStartValue = e[offsetProp];
      this.initActionEvent(offsetProp, type);
      const onEndAction = () => {
        this.triggerEvent(type, this.action[type], true);
        this.acting = false;
        this.resetAction();
        this.moveStartValue = 0;
        window.removeEventListener("mousemove", this.actionEvent!);
        this.actionEvent = null;
        window.removeEventListener("mouseup", onEndAction);
      };
      if (this.acting) {
        window.addEventListener("mousemove", this.actionEvent!);
        window.addEventListener("mouseup", onEndAction);
      }
    };

    const onMouseDown = (e: MouseEvent) => {
      if (!(this.action.resize.x || this.action.resize.y)) {
        if (!(this.action.select.x || this.action.select.y)) {
          return;
        } else {
          handleStartAction(e, "select");
        }
      } else {
        handleStartAction(e, "resize");
      }
    };

    const defaultEventListeners = {
      mousemove: onMouseMove,
      mousedown: onMouseDown,
    };
    this.registerListenerFromOnProp(
      defaultEventListeners,
      this.eventObserver,
      this
    );
  }

  initActionEvent(offsetProp: "x" | "y", type: Excel.Cell.Action) {
    if (this.acting) {
      this.actionEvent = throttle((e: MouseEvent) => {
        if (!this.acting) return;
        this.action[type].value = e[offsetProp] - this.moveStartValue;
        this.action[type].mouseX = e.x;
        this.action[type].mouseY = e.y;
        this.triggerEvent(type, this.action[type]);
      }, 100);
    }
  }

  checkHit(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    const scrollX =
      this.fixed.x || this.fixed.y ? this.scrollX : globalObj.SCROLL_X;
    const scrollY =
      this.fixed.x || this.fixed.y ? this.scrollY : globalObj.SCROLL_Y;
    if (
      !(
        offsetX < this.position!.leftTop.x - scrollX ||
        offsetX > this.position!.rightTop.x - scrollX ||
        offsetY < this.position!.leftTop.y - scrollY ||
        offsetY > this.position!.leftBottom.y - scrollY
      )
    ) {
      this.mouseEntered = true;
      if (this.fixed.y) {
        globalObj.SET_CURSOR("s-resize");
      } else if (this.fixed.x) {
        globalObj.SET_CURSOR("w-resize");
      } else {
        globalObj.SET_CURSOR("cell");
      }
    } else {
      this.mouseEntered = false;
    }
  }

  resetAction() {
    this.action = {
      select: {
        x: false,
        y: false,
        rowIndex: null,
        colIndex: null,
        value: null,
      },
      resize: {
        x: false,
        y: false,
        rowIndex: null,
        colIndex: null,
        value: null,
      },
    };
  }

  checkResizeOrSelect(e: MouseEvent) {
    if (
      !(
        (this.fixed.x && this.colIndex === 0) ||
        (this.fixed.y && this.rowIndex === 0)
      )
    ) {
      this.resetAction();
      return;
    }
    const { offsetX, offsetY } = e;
    const scrollX =
      this.fixed.x || this.fixed.y ? this.scrollX : globalObj.SCROLL_X;
    const scrollY =
      this.fixed.x || this.fixed.y ? this.scrollY : globalObj.SCROLL_Y;

    this.resetAction();

    if (this.fixed.y) {
      if (
        !(
          offsetX < this.position!.rightTop.x - scrollX - RESIZE_COL_SIZE ||
          offsetX > this.position!.rightTop.x - scrollX ||
          offsetY < this.position!.leftTop.y - scrollY ||
          offsetY > this.position!.leftBottom.y - scrollY
        )
      ) {
        globalObj.SET_CURSOR("col-resize");
        this.action.resize = {
          x: true,
          y: false,
          rowIndex: this.rowIndex!,
          colIndex: this.colIndex!,
        };
      } else {
        if (
          !(
            offsetX < this.position!.leftTop.x - scrollX ||
            offsetX > this.position!.rightTop.x - scrollX - RESIZE_COL_SIZE ||
            offsetY < this.position!.leftTop.y - scrollY ||
            offsetY > this.position!.leftBottom.y - scrollY
          )
        ) {
          this.action.select = {
            x: true,
            y: false,
            rowIndex: this.rowIndex!,
            colIndex: this.colIndex!,
          };
        }
      }
    } else if (this.fixed.x) {
      if (
        !(
          offsetX < this.position!.leftTop.x - scrollX ||
          offsetX > this.position!.rightTop.x - scrollX ||
          offsetY < this.position!.leftBottom.y - scrollY - RESIZE_ROW_SIZE ||
          offsetY > this.position!.leftBottom.y - scrollY
        )
      ) {
        globalObj.SET_CURSOR("row-resize");
        this.action.resize = {
          x: false,
          y: true,
          rowIndex: this.rowIndex!,
          colIndex: this.colIndex!,
        };
      } else {
        if (
          !(
            offsetX < this.position!.leftTop.x - scrollX ||
            offsetX > this.position!.rightTop.x - scrollX ||
            offsetY < this.position!.leftTop.y - scrollY ||
            offsetY > this.position!.leftBottom.y - scrollY - RESIZE_ROW_SIZE
          )
        ) {
          this.action.select = {
            x: false,
            y: true,
            rowIndex: this.rowIndex!,
            colIndex: this.colIndex!,
          };
        }
      }
    }
  }

  setBorderStyle(ctx: CanvasRenderingContext2D, side: Excel.Cell.BorderSide) {
    if (!this.border[side].solid) {
      ctx.setLineDash(DEFAULT_CELL_LINE_DASH);
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
    if (
      this.fixed.x ||
      this.fixed.y ||
      this.border.top.solid ||
      this.border.left.solid
    ) {
      if (this.fixed.x || this.fixed.y || this.border.top.solid) {
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
      }
      if (this.fixed.x || this.fixed.y || this.border.left.solid) {
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
