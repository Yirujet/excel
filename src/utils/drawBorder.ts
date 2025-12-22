export default (
  ctx: CanvasRenderingContext2D,
  startX: number,
  startY: number,
  endX: number,
  endY: number,
  color: string,
  lineWidth: number = 1,
  lineDashOffset?: number[]
) => {
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = lineWidth;
  if (lineDashOffset) {
    ctx.setLineDash(lineDashOffset);
  }
  ctx.beginPath();
  ctx.moveTo(Math.round(startX), Math.round(startY));
  ctx.lineTo(Math.round(endX), Math.round(endY));
  ctx.closePath();
  ctx.stroke();
  ctx.restore();
};
