import CellInput from "../CellInput";
import CellMergence from "../CellMergence";
import CellResizer from "../../plugins/CellResizer";
import CellSelector from "../CellSelector";
import FillHandle from "../FillHandle";
import Filling from "../Filling";
import HorizontalScrollbar from "../Scrollbar/HorizontalScrollbar";
import VerticalScrollbar from "../Scrollbar/VerticalScrollbar";
import Shadow from "../Shadow";

export default abstract class SheetApi {
  declare cells: Excel.Cell.CellInstance[][];
  declare mergedCells: Excel.Sheet.CellRange[];
  declare mode: Excel.Sheet.Mode;
  declare rowCount: number;
  declare colCount: number;
  declare layout: Excel.LayoutInfo | null;
  declare scroll: Excel.PositionPoint;
  declare realWidth: number;
  declare realHeight: number;
  declare selectedCells: Excel.Sheet.CellRange | null;
  declare fixedColWidth: number;
  declare fixedRowHeight: number;
  declare fixedColIndex: number;
  declare fixedRowIndex: number;
  declare fillingCells: Excel.Sheet.CellRange | null;
  declare horizontalScrollBar: HorizontalScrollbar | null;
  declare verticalScrollBar: VerticalScrollbar | null;
  declare horizontalScrollBarShadow: Shadow | null;
  declare verticalScrollBarShadow: Shadow | null;
  declare fillHandle: FillHandle | null;
  declare filling: Filling | null;
  declare cellResizer: CellResizer | null;
  declare cellSelector: CellSelector | null;
  declare cellMergence: CellMergence | null;
  declare cellInput: CellInput | null;
  private declare _ctx: CanvasRenderingContext2D | null;
  declare width: number;
  declare height: number;
  declare draw: () => void;
  declare initCells: (cells: Excel.Cell.CellInstance[][] | undefined) => void;
  declare render: (autoRegisteEvents: boolean) => void;

  getCell(rowIndex: number, colIndex: number) {
    return this.cells[rowIndex]?.[colIndex] || null;
  }

  clearCellMeta(cell: Excel.Cell.CellInstance) {
    cell.meta = null;
    cell.value = "";
    cell.valueSlices = [];
  }

  setCellStyle(cell: Excel.Cell.CellInstance, cellStyle: Excel.Cell.Style) {
    if (cellStyle.text) {
      cell.textStyle = {
        ...cell.textStyle,
        ...cellStyle.text,
      };
    }
    if (cellStyle.border) {
      cell.border.top = {
        ...cell.border.top!,
        ...cellStyle.border,
      };
      cell.border.left = {
        ...cell.border.left!,
        ...cellStyle.border,
      };
      cell.border.right = {
        ...cell.border.right!,
        ...cellStyle.border,
      };
      cell.border.bottom = {
        ...cell.border.bottom!,
        ...cellStyle.border,
      };
      const leftSiblingCell = this.getCell(cell.rowIndex!, cell.colIndex! - 1);
      if (leftSiblingCell && !leftSiblingCell.fixed.x) {
        leftSiblingCell.border.right = {
          ...leftSiblingCell.border.right!,
          ...cellStyle.border,
        };
      }
      const topSiblingCell = this.getCell(cell.rowIndex! - 1, cell.colIndex!);
      if (topSiblingCell && !topSiblingCell.fixed.y) {
        topSiblingCell.border.bottom = {
          ...topSiblingCell.border.bottom!,
          ...cellStyle.border,
        };
      }
      const rightSiblingCell = this.getCell(cell.rowIndex!, cell.colIndex! + 1);
      if (rightSiblingCell && !rightSiblingCell.fixed.x) {
        rightSiblingCell.border.left = {
          ...rightSiblingCell.border.left!,
          ...cellStyle.border,
        };
      }
      const bottomSiblingCell = this.getCell(
        cell.rowIndex! + 1,
        cell.colIndex!
      );
      if (bottomSiblingCell && !bottomSiblingCell.fixed.y) {
        bottomSiblingCell.border.top = {
          ...bottomSiblingCell.border.top!,
          ...cellStyle.border,
        };
      }
    }
  }

  setSelectionCellsStyle(
    selectedCells: Excel.Sheet.CellRange,
    cellStyle: Excel.Cell.Style
  ) {
    const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] = selectedCells;
    for (let i = minRowIndex; i <= maxRowIndex; i++) {
      for (let j = minColIndex; j <= maxColIndex; j++) {
        const cell = this.getCell(i, j);
        if (cell) {
          this.setCellStyle(cell, cellStyle);
        }
      }
    }
    this.draw();
  }

  setCellMeta(
    cell: Excel.Cell.CellInstance,
    cellMeta: Excel.Cell.Meta,
    needDraw: boolean = true
  ) {
    if (cellMeta) {
      cell.meta = cellMeta;
      cell.value = cellMeta.data;
    }
    if (needDraw) {
      this.draw();
    }
  }

  setCellImageMeta(cell: Excel.Cell.CellInstance, image: File) {
    const reader = new FileReader();
    reader.readAsDataURL(image);
    reader.onload = () => {
      const img = new Image();
      img.src = reader.result as string;
      img.onload = () => {
        this.setCellMeta(cell, {
          type: "image",
          data: {
            img,
            width: img.width,
            height: img.height,
            fit: "fill",
          },
        });
      };
    };
  }

  merge([
    minRowIndex,
    maxRowIndex,
    minColIndex,
    maxColIndex,
  ]: Excel.Sheet.CellRange) {
    this.mergedCells.push([minRowIndex, maxRowIndex, minColIndex, maxColIndex]);
    this.draw();
  }

  unmerge([
    minRowIndex,
    maxRowIndex,
    minColIndex,
    maxColIndex,
  ]: Excel.Sheet.CellRange) {
    this.mergedCells = this.mergedCells.filter(
      (range) =>
        !(
          range[0] >= minRowIndex &&
          range[1] <= maxRowIndex &&
          range[2] >= minColIndex &&
          range[3] <= maxColIndex
        )
    );
    this.draw();
  }

  adjust() {
    this.clear();
    this.destroy();
    if (this.mode === "view") {
      this.initCells(this.cells);
    } else {
      const contentCells = this.cells.slice(1).map((row) => row.slice(1));
      this.rowCount = contentCells.length;
      this.colCount = contentCells[0].length;
      this.initCells(contentCells);
    }
    this.render(false);
  }

  getCellPointByMousePosition(mouseX: number, mouseY: number) {
    const x = Math.max(
      Math.min(mouseX - this.layout!.x + (this.scroll.x || 0), this.realWidth),
      0
    );
    const y = Math.max(
      Math.min(mouseY - this.layout!.y + (this.scroll.y || 0), this.realHeight),
      0
    );
    return {
      x,
      y,
    };
  }

  getMergedRangesInSelectedCells() {
    const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
      this.selectedCells!;
    const mergedRangesInSelectedCells = this.mergedCells.filter((item) => {
      const [
        mergedMinRowIndex,
        mergedMaxRowIndex,
        mergedMinColIndex,
        mergedMaxColIndex,
      ] = item;
      return (
        mergedMinRowIndex >= minRowIndex &&
        mergedMaxRowIndex <= maxRowIndex &&
        mergedMinColIndex >= minColIndex &&
        mergedMaxColIndex <= maxColIndex
      );
    });
    return mergedRangesInSelectedCells;
  }

  findCellByPoint(
    x: number,
    y: number,
    ignoreFixedX = true,
    ignoreFixedY = true
  ) {
    let cell = null;
    if (ignoreFixedX) {
      x = Math.max(x, this.fixedColWidth);
    }
    if (ignoreFixedY) {
      y = Math.max(y, this.fixedRowHeight);
    }
    let rowIndex = this.cells.findIndex(
      (e) => e[0].position.leftTop.y <= y && e[0].position.leftBottom.y >= y
    );
    let colIndex = this.cells[0].findIndex(
      (e) => e.position.leftTop.x <= x && e.position.rightTop.x >= x
    );
    if (ignoreFixedX) {
      colIndex = Math.max(colIndex, this.fixedColIndex);
    }
    if (ignoreFixedY) {
      rowIndex = Math.max(rowIndex, this.fixedRowIndex);
    }
    colIndex = Math.max(colIndex, 0);
    rowIndex = Math.max(rowIndex, 0);
    cell = this.cells[rowIndex][colIndex];
    return cell;
  }

  clearSelectCells() {
    if (this.selectedCells) {
      this.selectedCells = null;
    }
  }

  clearFillingCells() {
    if (this.fillingCells) {
      this.fillingCells = null;
    }
  }

  destroy() {
    this.horizontalScrollBar = null;
    this.verticalScrollBar = null;
    this.horizontalScrollBarShadow = null;
    this.verticalScrollBarShadow = null;
    this.fillHandle = null;
    this.filling = null;
    // this.cellResizer = null;
    this.cellSelector = null;
    this.cellMergence = null;
    this.cellInput = null;
  }

  clear() {
    this._ctx!.clearRect(0, 0, this.width, this.height);
  }

  clearCells(fixedInX: boolean, fixedInY: boolean) {
    let w = this.width;
    let h = this.height;
    if (fixedInX) {
      w = this.fixedColWidth;
    }
    if (fixedInY) {
      h = this.fixedRowHeight;
    }
    this._ctx!.clearRect(0, 0, w, h);
  }
}
