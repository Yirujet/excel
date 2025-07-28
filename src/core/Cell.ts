import Sheet from "./Sheet";

class Cell implements Excel.Cell.CellInstance {
  width: number | null = null;
  height: number | null = null;
  rowIndex: number | null = null;
  colIndex: number | null = null;
  selected = false;
  cellName: string = "";
  x: number | null = null;
  y: number | null = null;
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

  getTextAlignOffsetX() {
    if (this.textStyle.align === "left") {
      return 0;
    }
    if (this.textStyle.align === "center") {
      return this.width! / 2;
    }
    return this.width!;
  }

  render(ctx: CanvasRenderingContext2D, scrollX: number, scrollY: number) {
    if (this.fixed) {
      ctx.save();
      this.setBorderStyle(ctx);
      ctx.strokeRect(
        this.x! - scrollX,
        this.y! - scrollY,
        this.width!,
        this.height!
      );
      ctx.restore();
    } else {
      ctx.save();
      this.setBorderStyle(ctx);
      ctx.beginPath();
      ctx.moveTo(
        this.position.rightTop.x - scrollX,
        this.position.rightTop.y - scrollY
      );
      ctx.lineTo(
        this.position.rightBottom.x - scrollX,
        this.position.rightBottom.y - scrollY
      );
      ctx.closePath();
      ctx.stroke();
      ctx.restore();
      ctx.save();
      this.setBorderStyle(ctx);
      ctx.beginPath();
      ctx.moveTo(
        this.position.leftBottom.x - scrollX,
        this.position.leftBottom.y - scrollY
      );
      ctx.lineTo(
        this.position.rightBottom.x - scrollX,
        this.position.rightBottom.y - scrollY
      );
      ctx.closePath();
      ctx.stroke();
      ctx.restore();
    }
    if (!this.hidden) {
      ctx.save();
      this.setTextStyle(ctx);
      const textAlignOffsetX = this.getTextAlignOffsetX();
      ctx.fillText(
        this.value,
        this.x! + textAlignOffsetX - scrollX,
        this.y! + this.height! / 2 - scrollY
      );
      ctx.restore();
    }
  }
}

export default Cell;
