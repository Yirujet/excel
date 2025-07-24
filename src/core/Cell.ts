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

  render(ctx: CanvasRenderingContext2D, scrollX: number, scrollY: number) {
    if (this.fixed) {
      ctx.save();
      ctx.strokeStyle = "#ccc";
      ctx.strokeRect(
        this.x! - scrollX,
        this.y! - scrollY,
        this.width!,
        this.height!
      );
      ctx.restore();
    } else {
      ctx.save();
      ctx.setLineDash([2, 4]);
      ctx.strokeStyle = "#ccc";
      ctx.textBaseline = "middle";
      ctx.textAlign = "center";
      ctx.strokeRect(
        this.x! - scrollX,
        this.y! - scrollY,
        this.width!,
        this.height!
      );
      ctx.restore();
    }
    ctx.save();
    ctx.fillStyle = "#000";
    ctx.textBaseline = "middle";
    ctx.textAlign = "center";
    ctx.fillText(
      this.value,
      this.x! + this.width! / 2 - scrollX,
      this.y! + this.height! / 2 - scrollY
    );
    ctx.restore();
  }
}

export default Cell;
