class Shadow {
  x: number;
  y: number;
  width: number;
  height: number;
  color: [string, string];
  direction: Excel.Scrollbar.Type;
  constructor(
    x: number,
    y: number,
    width: number,
    height: number,
    color: [string, string],
    direction: Excel.Scrollbar.Type
  ) {
    this.x = x;
    this.y = y;
    this.width = width;
    this.height = height;
    this.color = color;
    this.direction = direction;
  }

  render(ctx: CanvasRenderingContext2D) {
    const gradient = ctx.createLinearGradient(
      this.direction === "vertical" ? this.width / 2 : this.x,
      this.direction === "vertical" ? this.y : this.height / 2,
      this.direction === "vertical" ? this.width / 2 : this.x + this.width,
      this.direction === "vertical" ? this.y + this.height : this.height / 2
    );
    gradient.addColorStop(0, this.color[0]);
    gradient.addColorStop(1, this.color[1]);
    ctx.save();
    ctx.fillStyle = gradient;
    ctx.fillRect(this.x, this.y, this.width, this.height);
    ctx.restore();
  }
}

export default Shadow;
