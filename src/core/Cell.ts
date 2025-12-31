import getTextMetrics from "../utils/getTextMetrics";
import Element from "../components/Element";
import {
  DEFAULT_CELL_DIAGONAL_LINE_COLOR,
  DEFAULT_CELL_DIAGONAL_LINE_WIDTH,
  DEFAULT_CELL_DIAGONAL_TEXT_COLOR,
  DEFAULT_CELL_DIAGONAL_TEXT_FONT_SIZE,
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
import getImgDrawInfoByFillMode from "../utils/getImgDrawInfoByFillMode";
import drawBorder from "../utils/drawBorder";

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
  valueSlices = [];
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
    if (!this.border[side]!.solid) {
      ctx.setLineDash(DEFAULT_CELL_LINE_DASH);
    } else {
      ctx.setLineDash([]);
    }
    ctx.strokeStyle = this.border[side]!.color;
    if (this.border[side]!.bold) {
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

  isInMergedCells(mergedCells: Excel.Sheet.CellRange[]): boolean {
    return mergedCells.some((e) => {
      const [
        minRowIndexMerged,
        maxRowIndexMerged,
        minColIndexMerged,
        maxColIndexMerged,
      ] = e;
      return (
        minRowIndexMerged <= this.rowIndex! &&
        this.rowIndex! <= maxRowIndexMerged &&
        minColIndexMerged <= this.colIndex! &&
        this.colIndex! <= maxColIndexMerged
      );
    });
  }

  render(
    ctx: CanvasRenderingContext2D,
    scrollX: number,
    scrollY: number,
    mergedCells: Excel.Sheet.CellRange[]
  ) {
    if (this.isInMergedCells(mergedCells)) {
      return;
    }
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
    // 上边框
    if (this.border.top) {
      drawBorder(
        ctx,
        Math.round(this.position.leftTop.x - this.scrollX),
        Math.round(this.position.leftTop.y - this.scrollY),
        Math.round(this.position.rightTop.x - this.scrollX),
        Math.round(this.position.rightTop.y - this.scrollY),
        this.border.top!.color,
        this.border.top!.bold ? 2 : 1,
        !this.border.top!.solid ? DEFAULT_CELL_LINE_DASH : []
      );
    }

    // 左边框
    if (this.border.left) {
      drawBorder(
        ctx,
        Math.round(this.position.leftBottom.x - this.scrollX),
        Math.round(this.position.leftBottom.y - this.scrollY),
        Math.round(this.position.leftTop.x - this.scrollX),
        Math.round(this.position.leftTop.y - this.scrollY),
        this.border.left!.color,
        this.border.left!.bold ? 2 : 1,
        !this.border.left!.solid ? DEFAULT_CELL_LINE_DASH : []
      );
    }

    // 右边框
    if (this.border.right) {
      drawBorder(
        ctx,
        Math.round(this.position.rightTop.x - this.scrollX),
        Math.round(this.position.rightTop.y - this.scrollY),
        Math.round(this.position.rightTop.x - this.scrollX),
        Math.round(this.position.rightBottom.y - this.scrollY),
        this.border.right!.color,
        this.border.right!.bold ? 2 : 1,
        !this.border.right!.solid ? DEFAULT_CELL_LINE_DASH : []
      );
    }

    // 下边框
    if (this.border.bottom) {
      drawBorder(
        ctx,
        Math.round(this.position.rightBottom.x - this.scrollX),
        Math.round(this.position.rightBottom.y - this.scrollY),
        Math.round(this.position.leftBottom.x - this.scrollX),
        Math.round(this.position.leftBottom.y - this.scrollY),
        this.border.bottom!.color,
        this.border.bottom!.bold ? 2 : 1,
        !this.border.bottom!.solid ? DEFAULT_CELL_LINE_DASH : []
      );
    }
  }

  drawDataCell(ctx: CanvasRenderingContext2D) {
    switch (this.meta?.type) {
      case "text":
        let textList: string[] =
          this.valueSlices.length > 0 ? this.valueSlices : [this.value];
        textList.forEach((text, i: number) => {
          const textAlignOffsetX = this.getTextAlignOffsetX(this.width!);
          this.drawDataCellText(
            ctx,
            text,
            textAlignOffsetX,
            this.y! +
              (this.height! / (textList.length + 1)) * (i + 1) -
              this.scrollY
          );
          if (this.textStyle.underline) {
            this.drawDataCellUnderline(ctx, textAlignOffsetX);
          }
        });
        break;
      case "image":
        this.drawDataCellImage(ctx);
        break;
      case "diagonal":
        this.drawDataCellDiagonal(ctx);
        break;
    }
  }

  drawDataCellText(
    ctx: CanvasRenderingContext2D,
    text: string,
    textAlignOffsetX: number,
    y: number
  ) {
    ctx.save();
    this.setTextStyle(ctx);
    ctx.fillText(text, this.x! + textAlignOffsetX - this.scrollX, y);
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
    drawBorder(
      ctx,
      Math.round(this.x! + textAlignOffsetX - this.scrollX - underlineOffset),
      Math.round(this.y! + this.height! / 2 - this.scrollY + wordHeight / 2),
      Math.round(
        this.x! + textAlignOffsetX - this.scrollX - underlineOffset + wordWidth
      ),
      Math.round(this.y! + this.height! / 2 - this.scrollY + wordHeight / 2),
      this.textStyle.color,
      0.5
    );
    ctx.restore();
  }

  drawDataCellImage(ctx: CanvasRenderingContext2D) {
    if ((this.meta!.data as Excel.Cell.CellImageMetaData).img) {
      const { x, y, width, height } = getImgDrawInfoByFillMode(
        this.meta!.data as Excel.Cell.CellImageMetaData,
        {
          x: this.position.leftTop.x - this.scrollX + DEFAULT_CELL_PADDING,
          y: this.position.leftTop.y - this.scrollY + DEFAULT_CELL_PADDING,
          width: this.width! - DEFAULT_CELL_PADDING * 2,
          height: this.height! - DEFAULT_CELL_PADDING * 2,
        }
      )!;
      ctx.drawImage(
        (this.meta!.data as Excel.Cell.CellImageMetaData).img,
        x,
        y,
        width,
        height
      );
    }
  }

  drawDiagonalText(
    ctx: CanvasRenderingContext2D,
    prePoint: [number, number],
    curPoint: [number, number],
    text: string
  ) {
    const preAngle =
      Math.abs(prePoint[1] - this.position.leftTop.y) /
      Math.abs(prePoint[0] - this.position.leftTop.x);
    const curAngle =
      Math.abs(curPoint[1] - this.position.leftTop.y) /
      Math.abs(curPoint[0] - this.position.leftTop.x);
    const angle =
      Math.atan(preAngle) + (Math.atan(curAngle) - Math.atan(preAngle)) / 2;

    ctx.save();
    ctx.translate(
      this.position.leftTop.x - this.scrollX,
      this.position.leftTop.y - this.scrollY
    );
    ctx.rotate(angle);

    const midPoint = [
      (prePoint[0] + curPoint[0]) / 2,
      (prePoint[1] + curPoint[1]) / 2,
    ];
    const d = Math.sqrt(
      Math.pow(midPoint[0] - this.position.leftTop.x, 2) +
        Math.pow(midPoint[1] - this.position.leftTop.y, 2)
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

  drawDataCellDiagonal(ctx: CanvasRenderingContext2D) {
    const { direction, value } = this.meta!
      .data as Excel.Cell.CellDiagonalMetaData;
    const times =
      value.length & 1 ? Math.floor(value.length / 2) : value.length / 2 - 1;

    let endPoints: [number, number][] = [];
    for (let i = 1; i <= times; i++) {
      endPoints.push([
        this.position.rightTop.x,
        this.position.rightTop.y + i * (this.height! / (times + 1)),
      ]);
    }
    for (let i = times; i >= 1; i--) {
      endPoints.push([
        this.position.leftBottom.x + i * (this.width! / (times + 1)),
        this.position.leftBottom.y,
      ]);
    }
    if (!(value.length & 1)) {
      endPoints.splice(times, 0, [
        this.position.rightBottom.x,
        this.position.rightBottom.y,
      ]);
    }

    ctx.save();
    ctx.strokeStyle = DEFAULT_CELL_DIAGONAL_LINE_COLOR;
    ctx.lineWidth = DEFAULT_CELL_DIAGONAL_LINE_WIDTH;
    ctx.beginPath();

    const startX = Math.round(this.position.leftTop.x - this.scrollX);
    const startY = Math.round(this.position.leftTop.y - this.scrollY);

    endPoints.forEach(([x, y]) => {
      ctx.moveTo(startX, startY);
      ctx.lineTo(Math.round(x - this.scrollX), Math.round(y - this.scrollY));
    });

    ctx.stroke();
    ctx.restore();

    endPoints.forEach(([x, y], i) => {
      const prePoint: [number, number] =
        i > 0
          ? endPoints[i - 1]
          : [this.position.rightTop.x, this.position.rightTop.y];
      this.drawDiagonalText(ctx, prePoint, [x, y], value[i]);
    });

    this.drawDiagonalText(
      ctx,
      endPoints[endPoints.length - 1],
      [this.position.leftBottom.x, this.position.leftBottom.y],
      value[value.length - 1]
    );
  }
}

export default Cell;
