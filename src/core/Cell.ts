import getTextMetrics from "../utils/getTextMetrics";
import Element from "../components/Element";
import {
  DEFAULT_CELL_LINE_BOLD,
  DEFAULT_CELL_LINE_COLOR,
  DEFAULT_CELL_LINE_DASH,
  DEFAULT_CELL_LINE_SOLID,
  DEFAULT_CELL_PADDING,
  DEFAULT_CELL_TEXT_ALIGN,
  DEFAULT_CELL_TEXT_BACKGROUND_COLOR,
  DEFAULT_CELL_TEXT_BOLD,
  DEFAULT_CELL_TEXT_COLOR,
  DEFAULT_CELL_TEXT_FONT_FAMILY,
  DEFAULT_CELL_TEXT_FONT_SIZE,
  DEFAULT_CELL_TEXT_ITALIC,
  DEFAULT_CELL_TEXT_UNDERLINE,
} from "../config/index";

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
  meta: Excel.Cell.Meta = null;
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
      return DEFAULT_CELL_PADDING;
    }
    if (this.textStyle.align === "center") {
      return baseWidth / 2;
    }
    return baseWidth - DEFAULT_CELL_PADDING;
  }

  render(ctx: CanvasRenderingContext2D, scrollX: number, scrollY: number) {
    this.scrollX = scrollX;
    this.scrollY = scrollY;
    this.drawCellBg(ctx);
    this.drawCellBorder(ctx);
    if (!this.hidden) {
      this.drawDataCell(ctx);
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

  drawDataCell(ctx: CanvasRenderingContext2D) {
    switch (this.meta?.type) {
      case "text":
        const textAlignOffsetX = this.getTextAlignOffsetX(this.width!);
        this.drawDataCellText(ctx, textAlignOffsetX);
        if (this.textStyle.underline) {
          this.drawDataCellUnderline(ctx, textAlignOffsetX);
        }
        break;
      case "image":
        this.drawDataCellImage(ctx);
        break;
    }
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

  drawDataCellImage(ctx: CanvasRenderingContext2D) {
    ctx.drawImage(
      (this.meta!.data as Excel.Cell.CellImageMetaData).img,
      this.position.leftTop.x - this.scrollX + DEFAULT_CELL_PADDING,
      this.position.leftTop.y - this.scrollY + DEFAULT_CELL_PADDING,
      this.width! - DEFAULT_CELL_PADDING * 2,
      this.height! - DEFAULT_CELL_PADDING * 2
    );
  }
}

export default Cell;
