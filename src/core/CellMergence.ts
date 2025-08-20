import Element from "../components/Element";

class CellMergence extends Element<null> {
  layout: Excel.LayoutInfo;
  cells: Excel.Cell.CellInstance[][];
  fixedColWidth: number;
  fixedRowHeight: number;
  constructor(
    layout: Excel.LayoutInfo,
    cells: Excel.Cell.CellInstance[][],
    fixedColWidth: number,
    fixedRowHeight: number
  ) {
    super("");
    this.layout = layout;
    this.cells = cells;
    this.fixedColWidth = fixedColWidth;
    this.fixedRowHeight = fixedRowHeight;
  }
  render(
    ctx: CanvasRenderingContext2D,
    mergedCells: Excel.Sheet.CellRange[] | null,
    scrollX: number,
    scrollY: number
  ) {
    if (mergedCells) {
      mergedCells.forEach((e) => {
        const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] = e;
        const minX = this.cells[minRowIndex][minColIndex].position.leftTop.x!;
        const minY = this.cells[minRowIndex][minColIndex].position.leftTop.y!;
        const maxX =
          this.cells[maxRowIndex][maxColIndex].position.rightBottom.x!;
        const maxY =
          this.cells[maxRowIndex][maxColIndex].position.rightBottom.y!;
        const leftX = Math.max(minX - scrollX, this.fixedColWidth);
        const rightX = Math.min(maxX - scrollX, this.layout!.width);
        const topY = Math.max(minY - scrollY, this.fixedRowHeight);
        const bottomY = Math.min(maxY - scrollY, this.layout!.height);
        const leftTopCell = this.cells[minRowIndex][minColIndex];
        const rightBottomCell = this.cells[maxRowIndex][maxColIndex];
        const w =
          rightBottomCell.position.rightBottom.x! -
          leftTopCell.position.leftTop.x!;
        const h =
          rightBottomCell.position.rightBottom.y! -
          leftTopCell.position.leftTop.y!;
        if (leftX < rightX && topY < bottomY) {
          ctx.save();
          ctx.fillStyle = leftTopCell.textStyle.backgroundColor;
          ctx.beginPath();
          ctx.moveTo(leftX, topY);
          ctx.lineTo(rightX, topY);
          ctx.lineTo(rightX, bottomY);
          ctx.lineTo(leftX, bottomY);
          ctx.closePath();
          ctx.fill();
          ctx.save();
          ctx.font = `${leftTopCell.textStyle.italic ? "italic" : ""} ${
            leftTopCell.textStyle.bold ? "bold" : "normal"
          } ${leftTopCell.textStyle.fontSize}px ${
            leftTopCell.textStyle.fontFamily
          }`;
          ctx.textBaseline = "middle";
          ctx.textAlign = leftTopCell.textStyle.align as CanvasTextAlign;
          ctx.fillStyle = leftTopCell.textStyle.color;
          const textAlignOffsetX = leftTopCell.getTextAlignOffsetX(w);
          ctx.fillText(
            leftTopCell.value,
            leftTopCell.x! + textAlignOffsetX - scrollX,
            leftTopCell.y! + h / 2 - scrollY
          );
          ctx.restore();
        }
      });
    }
  }
}

export default CellMergence;
