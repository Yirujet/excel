export default (
  meta: Excel.Cell.CellImageMetaData,
  cellInfo: {
    x: number;
    y: number;
    width: number;
    height: number;
  }
) => {
  const { x, y, width, height } = cellInfo;
  const { width: imgWidth, height: imgHeight } = meta;
  let imgInfo;
  let scale;
  switch (meta.fit) {
    case "fill":
      imgInfo = {
        x,
        y,
        width,
        height,
      };
      break;
    case "contain":
      scale = Math.min(width / imgWidth, height / imgHeight);
      imgInfo = {
        x,
        y,
        width: imgWidth * scale,
        height: imgHeight * scale,
      };
      break;
    case "cover":
      scale = Math.max(width / imgWidth, height / imgHeight);
      imgInfo = {
        x,
        y,
        width: imgWidth * scale,
        height: imgHeight * scale,
      };
      break;
    case "none":
      imgInfo = {
        x,
        y,
        width: imgWidth,
        height: imgHeight,
      };
      break;
    default:
      break;
  }
  return imgInfo;
};
