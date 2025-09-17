import Element from "../components/Element";
import $10226 from "../utils/10226";
import EventObserver from "../utils/EventObserver";
import throttle from "../utils/throttle";
import Cell from "./Cell";
import CellResizer from "./CellResizer";
import HorizontalScrollbar from "./Scrollbar/HorizontalScrollbar";
import VerticalScrollbar from "./Scrollbar/VerticalScrollbar";
import CellSelector from "./CellSelector";
import CellMergence from "./CellMergence";
import Shadow from "./Shadow";
import {
  DEFAULT_CELL_COL_COUNT,
  DEFAULT_CELL_HEIGHT,
  DEFAULT_CELL_LINE_COLOR,
  DEFAULT_CELL_ROW_COUNT,
  DEFAULT_CELL_WIDTH,
  DEFAULT_FIXED_CELL_BACKGROUND_COLOR,
  DEFAULT_FIXED_CELL_COLOR,
  DEFAULT_GRADIENT_OFFSET,
  DEFAULT_GRADIENT_START_COLOR,
  DEFAULT_GRADIENT_STOP_COLOR,
  DEFAULT_INDEX_CELL_WIDTH,
  DEFAULT_SCROLLBAR_TRACK_SIZE,
  DEVIATION_COMPARE_VALUE,
  RESIZE_COL_SIZE,
  RESIZE_ROW_SIZE,
} from "../config/index";
import globalObj from "./globalObj";
import FillHandle from "./FillHandle";
import Scrollbar from "./Scrollbar/Scrollbar";
import Filling from "./Filling";
import debounce from "../utils/debounce";

class Sheet
  extends Element<HTMLCanvasElement>
  implements Excel.Sheet.SheetInstance
{
  private _ctx: CanvasRenderingContext2D | null = null;
  private _startCell: Excel.Cell.CellInstance | null = null;
  name = "";
  cells: Excel.Cell.CellInstance[][] = [];
  width = 0;
  height = 0;
  scroll: Excel.PositionPoint = { x: 0, y: 0 };
  horizontalScrollBar: HorizontalScrollbar | null = null;
  verticalScrollBar: VerticalScrollbar | null = null;
  cellResizer: CellResizer | null = null;
  cellSelector: CellSelector | null = null;
  cellMergence: CellMergence | null = null;
  horizontalScrollBarShadow: Shadow | null = null;
  verticalScrollBarShadow: Shadow | null = null;
  fillHandle: FillHandle | null = null;
  filling: Filling | null = null;
  sheetEventsObserver: Excel.Event.ObserverInstance = new EventObserver();
  globalEventsObserver: Excel.Event.ObserverInstance = new EventObserver();
  realWidth = 0;
  realHeight = 0;
  fixedRowIndex = 1;
  fixedColIndex = 1;
  fixedRowCells: Excel.Cell.CellInstance[][] = [];
  fixedColCells: Excel.Cell.CellInstance[][] = [];
  fixedCells: Excel.Cell.CellInstance[][] = [];
  fixedRowHeight = 0;
  fixedColWidth = 0;
  layout: Excel.LayoutInfo | null = null;
  resizeInfo: Excel.Cell.CellAction["resize"] = {
    x: false,
    y: false,
    rowIndex: null,
    colIndex: null,
    value: null,
  };
  selectInfo: Excel.Cell.CellAction["select"] = {
    x: false,
    y: false,
    rowIndex: null,
    colIndex: null,
    value: null,
  };
  selectedCells: Excel.Sheet.CellRange | null = null;
  mergedCells: Excel.Sheet.CellRange[] = [];
  isFilling = false;
  fillingCells: Excel.Sheet.CellRange | null = null;

  constructor(name: string, cells?: Excel.Cell.CellInstance[][]) {
    super("canvas");
    this.name = name;
    this.initCells(cells);
  }

  private exceed(x: number, y: number) {
    const exceedInfo = {
      x: {
        exceed: false,
        value: 0,
      },
      y: {
        exceed: false,
        value: 0,
      },
    };
    const offsetX = x - this.layout!.x;
    const offsetY = y - this.layout!.y;
    if (offsetX < this.fixedColWidth) {
      exceedInfo.x.exceed = true;
      exceedInfo.x.value = offsetX - this.fixedColWidth;
    } else if (offsetX > this.layout!.width) {
      exceedInfo.x.exceed = true;
      exceedInfo.x.value = offsetX - this.layout!.width;
    } else {
      exceedInfo.x.exceed = false;
      exceedInfo.x.value = 0;
    }
    if (offsetY < this.fixedRowHeight) {
      exceedInfo.y.exceed = true;
      exceedInfo.y.value = offsetY - this.fixedRowHeight;
    } else if (offsetY > this.layout!.height) {
      exceedInfo.y.exceed = true;
      exceedInfo.y.value = offsetY - this.layout!.height;
    } else {
      exceedInfo.y.exceed = false;
      exceedInfo.y.value = 0;
    }
    return exceedInfo;
  }

  private mergeIntersectMergedCells(
    mergedCells: Excel.Sheet.CellRange[],
    selectedCells: Excel.Sheet.CellRange
  ) {
    mergedCells.forEach((e, i) => {
      const [
        minRowIndexMerged,
        maxRowIndexMerged,
        minColIndexMerged,
        maxColIndexMerged,
      ] = e;
      const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
        selectedCells;
      if (
        (minRowIndex >= minRowIndexMerged &&
          minRowIndex <= maxRowIndexMerged &&
          !(
            minColIndex > maxColIndexMerged || maxColIndex < maxColIndexMerged
          )) ||
        (maxColIndex >= minColIndexMerged &&
          maxColIndex <= maxColIndexMerged &&
          !(
            minRowIndex > maxRowIndexMerged || maxRowIndex < minRowIndexMerged
          )) ||
        (maxRowIndex >= minRowIndexMerged &&
          maxRowIndex <= maxRowIndexMerged &&
          !(
            minColIndex > maxColIndexMerged || maxColIndex < maxColIndexMerged
          )) ||
        (minColIndex >= minColIndexMerged &&
          minColIndex <= maxColIndexMerged &&
          !(minRowIndex > maxRowIndexMerged || maxRowIndex < minRowIndexMerged))
      ) {
        selectedCells = this.mergeIntersectMergedCells(
          [...mergedCells.slice(0, i), ...mergedCells.slice(i + 1)],
          [
            Math.min(minRowIndex, minRowIndexMerged),
            Math.max(maxRowIndex, maxRowIndexMerged),
            Math.min(minColIndex, minColIndexMerged),
            Math.max(maxColIndex, maxColIndexMerged),
          ]
        );
      }
    });
    return selectedCells;
  }

  private autoScroll(x: number, y: number) {
    const exceedInfo = this.exceed(x, y);
    if (exceedInfo.x.exceed) {
      this.horizontalScrollBar!.value -= exceedInfo.x.value;
      if (
        this.horizontalScrollBar!.value +
          this.horizontalScrollBar!.track.width <=
        this.horizontalScrollBar!.thumb.width
      ) {
        this.horizontalScrollBar!.value =
          this.horizontalScrollBar!.thumb.width -
          this.horizontalScrollBar!.track.width;
        this.horizontalScrollBar!.isLast = true;
      }
      if (
        this.horizontalScrollBar!.value > 0 ||
        Math.abs(this.horizontalScrollBar!.value) <
          this.layout!.deviationCompareValue
      ) {
        this.horizontalScrollBar!.value = 0;
      }
      this.horizontalScrollBar!.percent =
        this.horizontalScrollBar!.value /
        (this.horizontalScrollBar!.thumb.width -
          this.horizontalScrollBar!.track.width);
      this.updateScroll(
        this.horizontalScrollBar!.percent,
        this.horizontalScrollBar!.type
      );
    }
    if (exceedInfo.y.exceed) {
      this.verticalScrollBar!.value -= exceedInfo.y.value;
      if (
        this.verticalScrollBar!.value + this.verticalScrollBar!.track.height <=
        this.verticalScrollBar!.thumb.height
      ) {
        this.verticalScrollBar!.value =
          this.verticalScrollBar!.thumb.height -
          this.verticalScrollBar!.track.height;
        this.verticalScrollBar!.isLast = true;
      }
      if (
        this.verticalScrollBar!.value > 0 ||
        Math.abs(this.verticalScrollBar!.value) <
          this.layout!.deviationCompareValue
      ) {
        this.verticalScrollBar!.value = 0;
      }
      this.verticalScrollBar!.percent =
        this.verticalScrollBar!.value /
        (this.verticalScrollBar!.thumb.height -
          this.verticalScrollBar!.track.height);
      this.updateScroll(
        this.verticalScrollBar!.percent,
        this.verticalScrollBar!.type
      );
    }
  }

  private getCellPointByMousePosition(mouseX: number, mouseY: number) {
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

  private selectCellRange(e: MouseEvent) {
    if (globalObj.EVENT_LOCKED) {
      return;
    }
    if (this.isFilling) {
      return;
    }
    const exceedInfo = this.exceed(e.x, e.y);
    if (exceedInfo.x.exceed || exceedInfo.y.exceed) {
      return;
    }
    globalObj.EVENT_LOCKED = true;
    const x = e.x - this.layout!.x + (this.scroll.x || 0);
    const y = e.y - this.layout!.y + (this.scroll.y || 0);
    this._startCell = this.findCellByPoint(x, y);
    if (this._startCell) {
      this.clearSelectCells();
      this.selectedCells = [
        this._startCell.rowIndex!,
        this._startCell.rowIndex!,
        this._startCell.colIndex!,
        this._startCell.colIndex!,
      ];
      const containedMergedCell = this.mergedCells.find((e) => {
        const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] = e;
        return (
          minRowIndex <= this._startCell!.rowIndex! &&
          maxRowIndex >= this._startCell!.rowIndex! &&
          minColIndex <= this._startCell!.colIndex! &&
          maxColIndex >= this._startCell!.colIndex!
        );
      });
      if (containedMergedCell) {
        this.selectedCells = [...containedMergedCell];
      }
    }
    const onSelectingCells = throttle(this.selectingCellRange.bind(this), 50);
    const onEndSelectCells = () => {
      if (this.selectedCells) {
        this.draw();
      }
      globalObj.EVENT_LOCKED = false;
      window.removeEventListener("mousemove", onSelectingCells);
      window.removeEventListener("mouseup", onEndSelectCells);
    };
    window.addEventListener("mousemove", onSelectingCells);
    window.addEventListener("mouseup", onEndSelectCells);
  }

  private selectingCellRange(e: MouseEvent) {
    if (this._startCell) {
      this.autoScroll(e.x, e.y);
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const endCell = this.findCellByPoint(x, y);
      this.clearSelectCells();
      if (endCell) {
        const minRowIndex = Math.min(
          this._startCell!.rowIndex!,
          endCell.rowIndex!
        );
        const maxRowIndex = Math.max(
          this._startCell!.rowIndex!,
          endCell.rowIndex!
        );
        const minColIndex = Math.min(
          this._startCell!.colIndex!,
          endCell.colIndex!
        );
        const maxColIndex = Math.max(
          this._startCell!.colIndex!,
          endCell.colIndex!
        );
        this.selectedCells = [
          minRowIndex,
          maxRowIndex,
          minColIndex,
          maxColIndex,
        ];
        this.selectedCells = this.mergeIntersectMergedCells(
          this.mergedCells,
          this.selectedCells!
        );
        this.draw();
      }
    }
  }

  private handleCellAction(
    e: MouseEvent,
    isInX: boolean,
    isInY: boolean,
    triggerEvent: (
      resize: Excel.Cell.CellAction[Excel.Cell.Action],
      isEnd?: boolean
    ) => void
  ) {
    if (!(isInX || isInY)) return;
    globalObj.EVENT_LOCKED = true;
    const offsetProp = isInX ? "x" : "y";
    const start = e[offsetProp];
    const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
    const pointX = isInY ? x - this.scroll.x : x;
    const pointY = isInX ? y - this.scroll.y : y;
    const cell = this.findCellByPoint(pointX, pointY, false, false);
    const onMove = throttle((e: MouseEvent) => {
      triggerEvent.call(this, {
        x: isInX,
        y: isInY,
        rowIndex: cell.rowIndex,
        colIndex: cell.colIndex,
        value: e[offsetProp] - start,
        mouseX: e.x,
        mouseY: e.y,
      });
    }, 100);
    const onEnd = (e: MouseEvent) => {
      triggerEvent.call(
        this,
        {
          x: isInX,
          y: isInY,
          rowIndex: cell.rowIndex,
          colIndex: cell.colIndex,
          value: e[offsetProp] - start,
          mouseX: e.x,
          mouseY: e.y,
        },
        true
      );
      globalObj.EVENT_LOCKED = false;
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onEnd);
    };
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onEnd);
  }

  private resizeCell(e: MouseEvent) {
    const isInColResize = this.pointInColResize(e);
    const isInRowResize = this.pointInRowResize(e);
    this.handleCellAction(
      e,
      isInColResize,
      isInRowResize,
      this.handleCellResize
    );
  }

  private selectFullCell(e: MouseEvent) {
    if (!this.pointInFixedCell(e)) return;
    const isInFixedYCell =
      this.pointInFixedXCell(e) && !this.pointInRowResize(e);
    const isInFixedXCell =
      this.pointInFixedYCell(e) && !this.pointInColResize(e);
    if (!(isInFixedXCell || isInFixedYCell)) return;
    this.handleCellAction(
      e,
      isInFixedXCell,
      isInFixedYCell,
      this.handleCellSelect
    );
  }

  private fill(e: MouseEvent) {
    if (!this.pointInFillHandle(e)) return;
    globalObj.EVENT_LOCKED = true;
    this.isFilling = true;
    const onFill = throttle(this.fillingCellRange.bind(this), 50);
    const onEndFill = () => {
      globalObj.EVENT_LOCKED = false;
      this.isFilling = false;
      this.clearFillingCells();
      window.removeEventListener("mousemove", onFill);
      window.removeEventListener("mouseup", onEndFill);
    };
    window.addEventListener("mousemove", onFill);
    window.addEventListener("mouseup", onEndFill);
  }

  private fillingCellRange(e: MouseEvent) {
    if (this.selectedCells) {
      this.autoScroll(e.x, e.y);
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const endCell = this.findCellByPoint(x, y);
      this.clearFillingCells();
      if (endCell) {
        console.log(endCell.value);
        this.draw();
      }
    }
  }

  private binaryQuery<T>(
    arr: T[],
    value: number,
    compare: (_binaryIndex: number, value: number, arr: T[]) => boolean,
    complete: (_binaryIndex: number, arr: T[]) => boolean,
    judge?: (_binaryIndex: number, value: number, arr: T[]) => boolean
  ) {
    let index = null;
    let binaryIndexStart = 0;
    let binaryIndexEnd = arr.length;
    let binaryIndex = Math.floor((binaryIndexStart + binaryIndexEnd) / 2);
    while (index === null) {
      if (complete(binaryIndex, arr)) {
        index = binaryIndex;
      } else {
        if (compare(binaryIndex, value, arr)) {
          if (!compare(binaryIndex - 1, value, arr)) {
            index = binaryIndex;
          } else {
            binaryIndexEnd = binaryIndex;
          }
        } else {
          if (judge) {
            if (judge(binaryIndex, value, arr)) {
              binaryIndexStart = binaryIndex;
            } else {
              binaryIndexEnd = binaryIndex;
            }
          } else {
            binaryIndexStart = binaryIndex;
          }
        }
        binaryIndex = Math.floor((binaryIndexStart + binaryIndexEnd) / 2);
      }
    }
    return index!;
  }

  private findCellByPoint(
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
    cell = this.cells[rowIndex][colIndex];
    return cell;
  }

  private pointInScrollbar(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    if (this.verticalScrollBar || this.horizontalScrollBar) {
      const checkInScrollbar = (
        offsetX: number,
        offsetY: number,
        scrollbar: Scrollbar | null
      ) => {
        if (scrollbar) {
          return !(
            offsetX < scrollbar.x ||
            offsetX > scrollbar.x + scrollbar.track.width ||
            offsetY < scrollbar.y ||
            offsetY > scrollbar.y + scrollbar.track.height
          );
        } else {
          return false;
        }
      };
      return (
        checkInScrollbar(offsetX, offsetY, this.verticalScrollBar) ||
        checkInScrollbar(offsetX, offsetY, this.horizontalScrollBar)
      );
    } else {
      return false;
    }
  }

  private pointInCellRange(e: MouseEvent) {
    const { x, y, offsetX, offsetY } = e;
    return (
      x >= this.layout!.x &&
      y >= this.layout!.y &&
      x <= this.layout!.x + this.layout!.width &&
      y <= this.layout!.y + this.layout!.height &&
      offsetX >= this.x - this.layout!.x &&
      offsetX <= this.layout!.width &&
      offsetY >= this.y - this.layout!.y &&
      offsetY <= this.layout!.height
    );
  }

  private pointInFixedCell(e: MouseEvent) {
    if (this.pointInCellRange(e)) {
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const cell = this.findCellByPoint(
        x - this.scroll.x,
        y - this.scroll.y,
        false,
        false
      );
      if (cell) {
        return cell.fixed.x || cell.fixed.y;
      }
    }
    return false;
  }

  private pointInFixedXCell(e: MouseEvent) {
    if (this.pointInFixedCell(e)) {
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const cell = this.findCellByPoint(
        x - this.scroll.x,
        y - this.scroll.y,
        false,
        false
      );
      if (cell) {
        return cell.fixed.x;
      }
    }
    return false;
  }

  private pointInFixedYCell(e: MouseEvent) {
    if (this.pointInFixedCell(e)) {
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const cell = this.findCellByPoint(
        x - this.scroll.x,
        y - this.scroll.y,
        false,
        false
      );
      if (cell) {
        return cell.fixed.y;
      }
    }
    return false;
  }

  private pointInNormalCell(e: MouseEvent) {
    if (this.pointInCellRange(e)) {
      return !this.pointInFixedCell(e);
    }
    return false;
  }

  private pointInFillHandle(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    if (this.fillHandle) {
      return (
        offsetX >= this.fillHandle.position.leftTop.x - this.scroll.x &&
        offsetX <= this.fillHandle.position.rightBottom.x - this.scroll.x &&
        offsetY >= this.fillHandle.position.leftTop.y - this.scroll.y &&
        offsetY <= this.fillHandle.position.rightBottom.y - this.scroll.y
      );
    } else {
      return false;
    }
  }

  private pointInRowResize(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    if (this.pointInFixedCell(e)) {
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const cell = this.findCellByPoint(x - this.scroll.x, y, false, false);
      if (cell) {
        if (cell.fixed.x) {
          return (
            offsetX >= cell.position!.leftTop.x &&
            offsetX <= cell.position!.rightTop.x &&
            offsetY >=
              cell.position!.leftBottom.y - this.scroll.y - RESIZE_ROW_SIZE &&
            offsetY <= cell.position!.leftBottom.y - this.scroll.y
          );
        } else {
          return false;
        }
      }
      return false;
    } else {
      return false;
    }
  }

  private pointInColResize(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    if (this.pointInFixedCell(e)) {
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const cell = this.findCellByPoint(x, y - this.scroll.y, false, false);
      if (cell) {
        if (cell.fixed.y) {
          return (
            offsetX >=
              cell.position!.rightTop.x - this.scroll.x - RESIZE_COL_SIZE &&
            offsetX <= cell.position!.rightTop.x - this.scroll.x &&
            offsetY >= cell.position!.leftTop.y &&
            offsetY <= cell.position!.leftBottom.y
          );
        } else {
          return false;
        }
      }
      return false;
    } else {
      return false;
    }
  }

  private updateCursor(e: MouseEvent) {
    if (globalObj.EVENT_LOCKED) return;
    if (!this.pointInCellRange(e)) {
      if (this.pointInScrollbar(e)) {
        globalObj.SET_CURSOR("default");
      } else {
        globalObj.SET_CURSOR("default");
      }
    } else {
      if (this.pointInFixedXCell(e)) {
        globalObj.SET_CURSOR("w-resize");
      }
      if (this.pointInFixedYCell(e)) {
        globalObj.SET_CURSOR("s-resize");
      }
      if (this.pointInNormalCell(e)) {
        globalObj.SET_CURSOR("cell");
      }
      if (this.pointInFillHandle(e)) {
        globalObj.SET_CURSOR("crosshair");
      }
      if (this.pointInRowResize(e)) {
        globalObj.SET_CURSOR("row-resize");
      }
      if (this.pointInColResize(e)) {
        globalObj.SET_CURSOR("col-resize");
      }
    }
  }

  private preventGlobalWheel(e: WheelEvent) {
    if (
      e.x >= this.x &&
      e.x <= this.x + this.width &&
      e.y >= this.y &&
      e.y <= this.y + this.height
    ) {
      e.stopPropagation();
      e.preventDefault();
    }
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

  getCell(rowIndex: number, colIndex: number) {
    return this.cells[rowIndex]?.[colIndex] || null;
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
        ...cell.border.top,
        ...cellStyle.border,
      };
      cell.border.left = {
        ...cell.border.left,
        ...cellStyle.border,
      };
      cell.border.right = {
        ...cell.border.right,
        ...cellStyle.border,
      };
      cell.border.bottom = {
        ...cell.border.bottom,
        ...cellStyle.border,
      };
      const leftSiblingCell = this.getCell(cell.rowIndex!, cell.colIndex! - 1);
      if (leftSiblingCell && !leftSiblingCell.fixed.x) {
        leftSiblingCell.border.right = {
          ...leftSiblingCell.border.right,
          ...cellStyle.border,
        };
      }
      const topSiblingCell = this.getCell(cell.rowIndex! - 1, cell.colIndex!);
      if (topSiblingCell && !topSiblingCell.fixed.y) {
        topSiblingCell.border.bottom = {
          ...topSiblingCell.border.bottom,
          ...cellStyle.border,
        };
      }
      const rightSiblingCell = this.getCell(cell.rowIndex!, cell.colIndex! + 1);
      if (rightSiblingCell && !rightSiblingCell.fixed.x) {
        rightSiblingCell.border.left = {
          ...rightSiblingCell.border.left,
          ...cellStyle.border,
        };
      }
      const bottomSiblingCell = this.getCell(
        cell.rowIndex! + 1,
        cell.colIndex!
      );
      if (bottomSiblingCell && !bottomSiblingCell.fixed.y) {
        bottomSiblingCell.border.top = {
          ...bottomSiblingCell.border.top,
          ...cellStyle.border,
        };
      }
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

  initEvents() {
    const globalEventListeners = {
      mousedown: (e: MouseEvent) => {
        this.fill.call(this, e);
        this.selectCellRange.call(this, e);
        this.resizeCell.call(this, e);
        this.selectFullCell.call(this, e);
      },
      mousemove: debounce(this.updateCursor.bind(this), 30),
      wheel: this.preventGlobalWheel.bind(this),
    };

    this.registerListenerFromOnProp(
      globalEventListeners,
      this.globalEventsObserver,
      this
    );
  }

  render() {
    this.initSheet();
    this.initScrollbar();
    this.initShadow();
    this.initCellResizer();
    this.initCellSelector();
    this.initCellMergence();
    this.initFillHandle();
    this.initFilling();
    this.initEvents();
    this.sheetEventsObserver.observe(this.$el!);
    this.globalEventsObserver.observe(window as any);
    this.draw();
  }

  initSheet() {
    this._ctx = this.$el!.getContext("2d")!;
    this.$el!.style.width = `${this.width}px`;
    this.$el!.style.height = `${this.height}px`;
    this.$el!.width = this.width * window.devicePixelRatio;
    this.$el!.height = this.height * window.devicePixelRatio;
    this._ctx!.translate(0.5, 0.5);
    this._ctx!.scale(window.devicePixelRatio, window.devicePixelRatio);
  }

  initLayout() {
    const bodyHeight = this.height - DEFAULT_CELL_HEIGHT;
    const verticalScrollbarShow = bodyHeight < this.realHeight;
    const horizontalScrollbarShow =
      this.realWidth > this.width - DEFAULT_INDEX_CELL_WIDTH;
    this.layout = {
      x: this.x,
      y: this.y,
      width:
        this.width - (verticalScrollbarShow ? DEFAULT_SCROLLBAR_TRACK_SIZE : 0),
      height:
        this.height -
        (horizontalScrollbarShow ? DEFAULT_SCROLLBAR_TRACK_SIZE : 0),
      headerHeight: DEFAULT_CELL_HEIGHT,
      fixedLeftWidth: DEFAULT_INDEX_CELL_WIDTH,
      bodyHeight,
      bodyRealWidth: this.realWidth,
      bodyRealHeight: this.realHeight,
      deviationCompareValue: DEVIATION_COMPARE_VALUE,
    };
  }

  initCells(cells: Excel.Cell.CellInstance[][] | undefined) {
    if (cells) {
      this.cells = cells!;
    } else {
      this.cells = [];
      this.fixedColWidth = 0;
      this.fixedRowHeight = 0;
      let x = 0;
      let y = 0;
      for (let i = 0; i < DEFAULT_CELL_ROW_COUNT + 1; i++) {
        let row: Excel.Cell.CellInstance[] = [];
        y = i * DEFAULT_CELL_HEIGHT;
        x = 0;
        let fixedColRows: Excel.Cell.CellInstance[] = [];
        let fixedRows: Excel.Cell.CellInstance[] = [];
        for (let j = 0; j < DEFAULT_CELL_COL_COUNT + 1; j++) {
          x =
            j === 0
              ? 0
              : (j - 1) * DEFAULT_CELL_WIDTH + DEFAULT_INDEX_CELL_WIDTH;
          const cell = new Cell(this.sheetEventsObserver);
          cell.x = x;
          cell.y = y;
          cell.width = j === 0 ? DEFAULT_INDEX_CELL_WIDTH : DEFAULT_CELL_WIDTH;
          cell.height = DEFAULT_CELL_HEIGHT;
          cell.rowIndex = i;
          cell.colIndex = j;
          cell.cellName = $10226(j - 1);
          cell.updatePosition();
          if (i === 0) {
            this.setCellMeta(
              cell,
              {
                type: "text",
                data: cell.cellName,
              },
              false
            );
          }
          if (j === 0) {
            this.setCellMeta(
              cell,
              {
                type: "text",
                data: i.toString(),
              },
              false
            );
          }
          if (i === 0 && j === 0) {
            cell.hidden = true;
          }
          if (i > 0 && j > 0) {
            // cell.value = i.toString() + "-" + j.toString();
            this.setCellMeta(
              cell,
              {
                type: "text",
                data: i.toString() + "-" + j.toString(),
              },
              false
            );
          }
          if (j < this.fixedColIndex) {
            fixedColRows.push(cell);
            if (i < this.fixedRowIndex) {
              fixedRows.push(cell);
            }
            if (i === 0) {
              this.fixedColWidth += cell.width!;
            }
          }
          if (i < this.fixedRowIndex || j < this.fixedColIndex) {
            if (i < this.fixedRowIndex) {
              cell.fixed.y = true;
            }
            if (j < this.fixedColIndex) {
              cell.fixed.x = true;
            }
            this.setCellStyle(cell, {
              border: {
                solid: true,
                color: DEFAULT_CELL_LINE_COLOR,
                bold: false,
              },
              text: {
                color: DEFAULT_FIXED_CELL_COLOR,
                backgroundColor: DEFAULT_FIXED_CELL_BACKGROUND_COLOR,
                fontSize: 13,
                align: "center",
              },
            });
          } else {
            this.setCellStyle(cell, {
              border: {
                solid: false,
                color: DEFAULT_CELL_LINE_COLOR,
                bold: false,
              },
              text: {
                align: "center",
              },
            });
          }
          row.push(cell);
        }
        this.fixedColCells.push(fixedColRows);
        if (i < this.fixedRowIndex) {
          this.fixedCells.push(fixedRows);
          this.fixedRowCells.push(row);
          this.fixedRowHeight += row[0].height!;
        }
        this.cells.push(row);
      }
    }
    if (this.cells.length > 0) {
      this.realWidth = this.cells[0].reduce((p, c) => p + c.width!, 0);
      this.realHeight = this.cells.reduce((p, c) => p + c[0].height!, 0);
    }
  }

  initCellResizer() {
    this.cellResizer = new CellResizer(this.layout!);
  }

  initCellSelector() {
    this.cellSelector = new CellSelector(
      this.layout!,
      this.cells,
      this.fixedColWidth,
      this.fixedRowHeight
    );
  }

  initCellMergence() {
    this.cellMergence = new CellMergence(
      this.layout!,
      this.cells,
      this.fixedColWidth,
      this.fixedRowHeight
    );
  }

  initScrollbar() {
    this.initLayout();
    this.horizontalScrollBar = new HorizontalScrollbar(
      this.layout!,
      this.sheetEventsObserver,
      this.globalEventsObserver
    );
    this.verticalScrollBar = new VerticalScrollbar(
      this.layout!,
      this.sheetEventsObserver,
      this.globalEventsObserver
    );
    this.horizontalScrollBar.addEvent("percent", this.redraw.bind(this));
    this.verticalScrollBar.addEvent("percent", this.redraw.bind(this));
  }

  initShadow() {
    this.horizontalScrollBarShadow = new Shadow(
      this.fixedColWidth,
      this.fixedRowHeight,
      this.width,
      DEFAULT_GRADIENT_OFFSET,
      [DEFAULT_GRADIENT_START_COLOR, DEFAULT_GRADIENT_STOP_COLOR],
      "vertical"
    );
    this.verticalScrollBarShadow = new Shadow(
      this.fixedColWidth,
      this.fixedRowHeight,
      DEFAULT_GRADIENT_OFFSET,
      this.height,
      [DEFAULT_GRADIENT_START_COLOR, DEFAULT_GRADIENT_STOP_COLOR],
      "horizontal"
    );
  }

  initFillHandle() {
    this.fillHandle = new FillHandle(
      this.sheetEventsObserver,
      this.layout!,
      this.cells,
      this.fixedColWidth,
      this.fixedRowHeight
    );
  }

  initFilling() {
    this.filling = new Filling(
      this.layout!,
      this.cells,
      this.fixedColWidth,
      this.fixedRowHeight
    );
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

  updateScroll(percent: number, type: Excel.Scrollbar.Type) {
    if (type === "horizontal") {
      this.scroll.x =
        Math.abs(percent) *
        (this.realWidth -
          this.width +
          (this.verticalScrollBar?.show ? DEFAULT_SCROLLBAR_TRACK_SIZE : 0));
      globalObj.SCROLL_X = this.scroll.x;
    } else {
      this.scroll.y =
        Math.abs(percent) *
        (this.realHeight -
          this.height +
          (this.horizontalScrollBar?.show ? DEFAULT_SCROLLBAR_TRACK_SIZE : 0));
      globalObj.SCROLL_Y = this.scroll.y;
    }
  }

  handleCellResize(resize: Excel.Cell.CellAction["resize"], isEnd = false) {
    if (resize.value) {
      this.resizeInfo = resize;
    }
    if (isEnd) {
      if (this.resizeInfo.x) {
        this.cells.forEach((row) => {
          row.forEach((cell, colIndex) => {
            if (colIndex === this.resizeInfo.colIndex!) {
              cell.width = cell.width! + this.resizeInfo.value!;
              cell.updatePosition();
            }
            if (colIndex > this.resizeInfo.colIndex!) {
              cell.x = cell.x! + this.resizeInfo.value!;
              cell.updatePosition();
            }
          });
        });
        this.layout!.bodyRealWidth += this.resizeInfo.value!;
        this.realWidth += this.resizeInfo.value!;
        this.horizontalScrollBar?.updateScrollbarInfo();
        this.horizontalScrollBar?.updatePosition();
      }
      if (this.resizeInfo.y) {
        this.cells.forEach((row, rowIndex) => {
          row.forEach((cell) => {
            if (rowIndex === this.resizeInfo.rowIndex!) {
              cell.height = cell.height! + this.resizeInfo.value!;
              cell.updatePosition();
            }
            if (rowIndex > this.resizeInfo.rowIndex!) {
              cell.y = cell.y! + this.resizeInfo.value!;
              cell.updatePosition();
            }
          });
        });
        this.layout!.bodyRealHeight += this.resizeInfo.value!;
        this.realHeight += this.resizeInfo.value!;
        this.verticalScrollBar?.updateScrollbarInfo();
        this.verticalScrollBar?.updatePosition();
      }
      this.resizeInfo = {
        x: false,
        y: false,
        rowIndex: null,
        colIndex: null,
        value: null,
      };
    }
    this.draw();
  }

  handleCellSelect(select: Excel.Cell.CellAction["select"], isEnd = false) {
    if (this.selectInfo.value === null) {
      if (select.x) {
        this._startCell = this.cells[this.fixedRowIndex][select.colIndex!];
        this.selectedCells = [
          this.fixedRowIndex,
          this.cells.length - 1,
          select.colIndex!,
          select.colIndex!,
        ];
      } else {
        this._startCell = this.cells[select.rowIndex!][this.fixedColIndex];
        this.selectedCells = [
          select.rowIndex!,
          select.rowIndex!,
          this.fixedColIndex,
          this.cells[0].length - 1,
        ];
      }
    } else {
      this.autoScroll(select.mouseX!, select.mouseY!);
      const { x, y } = this.getCellPointByMousePosition(
        select.mouseX!,
        select.mouseY!
      );
      const endCell = this.findCellByPoint(x, y);
      this.clearSelectCells();
      if (endCell) {
        if (select.x) {
          const endCellColIndex = endCell.colIndex!;
          const minColIndex = Math.min(select.colIndex!, endCellColIndex);
          const maxColIndex = Math.max(select.colIndex!, endCellColIndex);
          this.selectedCells = [
            this.fixedRowIndex,
            this.cells.length - 1,
            minColIndex,
            maxColIndex,
          ];
        }
        if (select.y) {
          const endCellRowIndex = endCell.rowIndex!;
          const minRowIndex = Math.min(select.rowIndex!, endCellRowIndex);
          const maxRowIndex = Math.max(select.rowIndex!, endCellRowIndex);
          this.selectedCells = [
            minRowIndex,
            maxRowIndex,
            this.fixedColIndex,
            this.cells[0].length - 1,
          ];
        }
      }
    }
    this.selectedCells = this.mergeIntersectMergedCells(
      this.mergedCells,
      this.selectedCells!
    );
    this.selectInfo = select;
    if (isEnd) {
      this.selectInfo = {
        x: false,
        y: false,
        rowIndex: null,
        colIndex: null,
        value: null,
      };
    }
    this.draw();
  }

  redraw(percent: number, type: Excel.Scrollbar.Type) {
    this.updateScroll(percent, type);
    this.draw();
  }

  draw() {
    this.drawSheetCells();
    this.drawMergedCells();
    this.drawShadow();
    this.drawScrollbar();
    this.drawCellSelector();
    this.drawCellResizer();
    this.drawFillHandle();
    this.drawFilling();
  }

  getRangeInView(
    cells: Excel.Cell.CellInstance[][],
    scrollX: number,
    scrollY: number,
    fixedInX: boolean,
    fixedInY: boolean
  ) {
    let minYIndex = this.binaryQuery(
      cells,
      0,
      (binaryIndex, value, arr) => {
        return arr[binaryIndex][0].position.leftBottom.y! - scrollY > value;
      },
      (binaryIndex, arr) => {
        return binaryIndex === 0;
      }
    );
    minYIndex = Math.max(minYIndex, 0);

    let maxYIndex = null;
    if (!fixedInY) {
      maxYIndex = this.binaryQuery(
        cells,
        this.height,
        (binaryIndex, value, arr) => {
          return arr[binaryIndex][0].position.leftTop.y! - scrollY > value;
        },
        (binaryIndex, arr) => {
          return binaryIndex === arr.length - 1;
        }
      );
      maxYIndex =
        this.verticalScrollBar?.percent === 1
          ? cells.length - 1
          : maxYIndex === -1
          ? cells.length - 1
          : maxYIndex;
    } else {
      maxYIndex = cells.length - 1;
    }
    let minXIndex = this.binaryQuery(
      cells[0],
      0,
      (binaryIndex, value, arr) => {
        return arr[binaryIndex].position.rightTop.x! - scrollX > value;
      },
      (binaryIndex, arr) => {
        return binaryIndex === 0;
      }
    );
    minXIndex = Math.max(minXIndex, 0);

    let maxXIndex = null;
    if (!fixedInX) {
      maxXIndex = this.binaryQuery(
        cells[0],
        this.width,
        (binaryIndex, value, arr) => {
          return arr[binaryIndex].position.leftTop.x! - scrollX > value;
        },
        (binaryIndex, arr) => {
          return binaryIndex === arr.length - 1;
        }
      );
      maxXIndex =
        this.horizontalScrollBar?.percent === 1
          ? cells[0].length - 1
          : maxXIndex === -1
          ? cells[0].length - 1
          : maxXIndex;
    } else {
      maxXIndex = cells[0].length - 1;
    }
    return [minXIndex, maxXIndex, minYIndex, maxYIndex];
  }

  drawSheetCells() {
    this.sheetEventsObserver.clearEventsWhenReRender();
    this.drawCells(
      this.cells,
      false,
      false,
      this.fixedColIndex,
      this.fixedRowIndex
    );
    this.drawCells(this.fixedRowCells, false, true, this.fixedColIndex, null);
    this.drawCells(this.fixedColCells, true, false, null, this.fixedRowIndex);
    this.drawCells(this.fixedCells, true, true, null, null);
  }

  drawCells(
    cells: Excel.Cell.CellInstance[][],
    fixedInX: boolean,
    fixedInY: boolean,
    ignoreXIndex: number | null,
    ignoreYIndex: number | null
  ) {
    this.clearCells(fixedInX, fixedInY);
    const scrollX = fixedInX ? 0 : this.scroll.x || 0;
    const scrollY = fixedInY ? 0 : this.scroll.y || 0;
    const [minXIndex, maxXIndex, minYIndex, maxYIndex] = this.getRangeInView(
      cells,
      scrollX,
      scrollY,
      fixedInX,
      fixedInY
    );
    for (let i = minYIndex; i <= maxYIndex; i++) {
      if (ignoreYIndex !== null && i < ignoreYIndex) continue;
      for (let j = minXIndex; j <= maxXIndex; j++) {
        if (ignoreXIndex !== null && j < ignoreXIndex) continue;
        const cell = cells[i][j];
        const {
          position: { leftTop, rightTop, rightBottom, leftBottom },
        } = cell;
        if (
          leftTop.x - scrollX > this.width ||
          leftBottom.y - scrollY > this.height
        ) {
          break;
        }
        if (!fixedInX && rightTop.x - scrollX < this.fixedColWidth) {
          continue;
        }
        if (!fixedInY && rightBottom.y - scrollY < this.fixedRowHeight) {
          continue;
        }
        cell.render(this._ctx!, scrollX, scrollY);
        let index = this.sheetEventsObserver.resize.findIndex(
          (e) => e === cell
        );
        if (!!~index) {
          this.sheetEventsObserver.resize.splice(index, 1);
        }
      }
    }
  }

  drawScrollbar() {
    if (this.verticalScrollBar!.show) {
      this.verticalScrollBar!.render(this._ctx!);
    }
    if (this.horizontalScrollBar!.show) {
      this.horizontalScrollBar!.render(this._ctx!);
    }
    if (this.verticalScrollBar!.show && this.horizontalScrollBar!.show) {
      this.verticalScrollBar!.fillCoincide(this._ctx!);
      this.horizontalScrollBar!.fillCoincide(this._ctx!);
    }
  }

  drawCellResizer() {
    if (this.resizeInfo.x || this.resizeInfo.y) {
      let cellInfo: Excel.Cell.CellInstance =
        this.cells[this.resizeInfo.rowIndex!][this.resizeInfo.colIndex!];
      this.cellResizer!.render(
        this._ctx!,
        cellInfo,
        this.resizeInfo,
        this.scroll
      );
    }
  }

  drawCellSelector() {
    this.cellSelector!.render(
      this._ctx!,
      this.selectedCells,
      this._startCell,
      this.scroll.x || 0,
      this.scroll.y || 0,
      this.mergedCells
    );
  }

  drawMergedCells() {
    this.cellMergence!.render(
      this._ctx!,
      this.mergedCells,
      this.scroll.x || 0,
      this.scroll.y || 0
    );
  }

  drawShadow() {
    if (
      this.fixedRowCells.length > 0 &&
      this.verticalScrollBar!.show &&
      this.verticalScrollBar!.percent > 0
    ) {
      this.horizontalScrollBarShadow!.render(this._ctx!);
    }
    if (
      this.fixedColCells.length > 0 &&
      this.horizontalScrollBar!.show &&
      this.horizontalScrollBar!.percent > 0
    ) {
      this.verticalScrollBarShadow!.render(this._ctx!);
    }
  }

  drawFillHandle() {
    this.fillHandle!.render(
      this._ctx!,
      this.selectedCells,
      this.scroll.x || 0,
      this.scroll.y || 0
    );
  }

  drawFilling() {
    this.filling!.render(
      this._ctx!,
      this.fillingCells,
      this.scroll.x || 0,
      this.scroll.y || 0
    );
  }
}

export default Sheet;
