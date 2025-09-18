export default (
  range: Excel.Sheet.CellRange,
  scrollX: number,
  scrollY: number,
  layout: Excel.LayoutInfo,
  cells: Excel.Cell.CellInstance[][],
  fixedColWidth: number,
  fixedRowHeight: number
) => {
  const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] = range;
  const minX = cells[minRowIndex][minColIndex].position.leftTop.x!;
  const minY = cells[minRowIndex][minColIndex].position.leftTop.y!;
  const maxX = cells[maxRowIndex][maxColIndex].position.rightBottom.x!;
  const maxY = cells[maxRowIndex][maxColIndex].position.rightBottom.y!;
  const leftX = Math.max(minX - scrollX, fixedColWidth);
  const rightX = Math.min(maxX - scrollX, layout!.width);
  const topY = Math.max(minY - scrollY, fixedRowHeight);
  const bottomY = Math.min(maxY - scrollY, layout!.height);

  const topBorderShow =
    minY - scrollY >= fixedRowHeight &&
    minY - scrollY <= layout!.height &&
    leftX < rightX;
  const bottomBorderShow =
    maxY - scrollY <= layout!.height &&
    maxY - scrollY >= fixedRowHeight &&
    leftX < rightX;
  const leftBorderShow =
    minX - scrollX >= fixedColWidth &&
    minX - scrollX <= layout!.width &&
    topY < bottomY;
  const rightBorderShow =
    maxX - scrollX <= layout!.width &&
    maxX - scrollX >= fixedColWidth &&
    topY < bottomY;

  return {
    minX,
    minY,
    maxX,
    maxY,
    leftX,
    rightX,
    topY,
    bottomY,
    topBorderShow,
    bottomBorderShow,
    leftBorderShow,
    rightBorderShow,
  };
};
