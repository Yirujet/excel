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
  DEFAULT_CELL_PADDING,
  DEFAULT_CELL_ROW_COUNT,
  DEFAULT_CELL_TEXT_FONT_SIZE,
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
import CellInput from "./CellInput";
import getTextMetrics from "../utils/getTextMetrics";

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
  mode: Excel.Sheet.Mode = "edit";
  margin: Exclude<Excel.Sheet.Configuration["margin"], undefined> = {
    right: 0,
    bottom: 0,
  };
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
  cellInput: CellInput | null = null;
  sheetEventsObserver: Excel.Event.ObserverInstance = new EventObserver();
  globalEventsObserver: Excel.Event.ObserverInstance = new EventObserver();
  realWidth = 0;
  realHeight = 0;
  fixedRowIndex = 1;
  fixedColIndex = 1;
  rowCount = DEFAULT_CELL_ROW_COUNT;
  colCount = DEFAULT_CELL_COL_COUNT;
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
  editingCell: Excel.Cell.CellInstance | null = null;
  fixedWidths: Array<number | undefined> = [];

  constructor(name: string, config: Excel.Sheet.Configuration) {
    super("canvas");
    this.name = name;
    this.mode = config.mode || "edit";
    this.fixedRowIndex = config.fixedRowIndex;
    this.fixedColIndex = config.fixedColIndex;
    this.rowCount = config.rowCount;
    this.colCount = config.colCount;
    this.mergedCells = config.mergedCells || [];
    this.margin = config.margin || {
      right: DEFAULT_CELL_WIDTH,
      bottom: DEFAULT_CELL_HEIGHT,
    };
    this.initCells(config?.cells);
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
    const mergedCellsCopy = [...mergedCells];
    const processedIndices = new Set<number>();
    let i = 0;
    while (i < mergedCellsCopy.length) {
      if (processedIndices.has(i)) {
        i++;
        continue;
      }

      const mergedCell = mergedCellsCopy[i];
      const [
        minRowIndexMerged,
        maxRowIndexMerged,
        minColIndexMerged,
        maxColIndexMerged,
      ] = mergedCell;

      const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
        selectedCells;
      const isOverlap =
        minRowIndex <= maxRowIndexMerged &&
        maxRowIndex >= minRowIndexMerged &&
        minColIndex <= maxColIndexMerged &&
        maxColIndex >= minColIndexMerged;

      if (isOverlap) {
        selectedCells = [
          Math.min(minRowIndex, minRowIndexMerged),
          Math.max(maxRowIndex, maxRowIndexMerged),
          Math.min(minColIndex, minColIndexMerged),
          Math.max(maxColIndex, maxColIndexMerged),
        ];
        processedIndices.add(i);
        i = 0;
      } else {
        i++;
      }
    }

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
    const isInAbsFixedCell = this.pointInAbsFixedCell(e);
    const isInColResize = this.pointInColResize(e);
    const isInRowResize = this.pointInRowResize(e);
    if (isInAbsFixedCell) return;
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
    if (isInFixedXCell && isInFixedYCell) {
      this._startCell = this.cells[this.fixedRowIndex][this.fixedColIndex];
      this.selectedCells = [
        this.fixedRowIndex,
        this.cells.length - 1,
        this.fixedColIndex,
        this.cells[0].length - 1,
      ];
      this.draw();
      return;
    }
    this.handleCellAction(
      e,
      isInFixedXCell,
      isInFixedYCell,
      this.handleCellSelect
    );
  }

  private editCell(e: MouseEvent) {
    if (this.pointInFixedCell(e)) return;
    const x = e.x - this.layout!.x + (this.scroll.x || 0);
    const y = e.y - this.layout!.y + (this.scroll.y || 0);
    const cell = this.findCellByPoint(x, y);
    this.editingCell = cell;
    const containedMergedCell = this.mergedCells.find((e) => {
      const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] = e;
      return (
        minRowIndex <= cell.rowIndex! &&
        maxRowIndex >= cell.rowIndex! &&
        minColIndex <= cell.colIndex! &&
        maxColIndex >= cell.colIndex!
      );
    });
    if (containedMergedCell) {
      const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
        containedMergedCell;
      this.editingCell = {
        ...this.cells[minRowIndex][minColIndex],
        width:
          this.cells[minRowIndex][maxColIndex].position.rightTop.x -
          this.cells[minRowIndex][minColIndex].position.leftBottom.x,
        height:
          this.cells[maxRowIndex][minColIndex].position.leftBottom.y -
          this.cells[minRowIndex][minColIndex].position.leftTop.y,
      };
    }
    this.drawCellEditor();
  }

  private clearCellsMeta(e: KeyboardEvent) {
    if (e.key.toLocaleLowerCase() !== "delete") return;
    if (this.selectedCells === null) return;
    const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
      this.selectedCells!;
    for (let r = minRowIndex; r <= maxRowIndex; r++) {
      for (let c = minColIndex; c <= maxColIndex; c++) {
        const cell = this.cells[r][c];
        this.clearCellMeta(cell);
      }
    }
    this.draw();
  }

  private endEditingCell(e: MouseEvent) {
    const x = e.x - this.layout!.x + (this.scroll.x || 0);
    const y = e.y - this.layout!.y + (this.scroll.y || 0);
    const cell = this.findCellByPoint(x, y);
    if (cell === this.editingCell) return;
    this.editingCell = null;
  }

  private repeatByMergedCells(
    minIndex: number,
    maxIndex: number,
    mergedCells: Excel.Sheet.CellRange[],
    times: number,
    d: number
  ) {
    if (mergedCells.length > 0) {
      for (let i = 0; i < times; i++) {
        mergedCells.forEach((item) => {
          const [
            mergedMinRowIndex,
            mergedMaxRowIndex,
            mergedMinColIndex,
            mergedMaxColIndex,
          ] = item;
          const nextItem: Excel.Sheet.CellRange = [
            mergedMinRowIndex,
            mergedMaxRowIndex,
            mergedMinColIndex,
            mergedMaxColIndex,
          ];
          const offset =
            (i + 1) *
            d *
            (this.selectedCells![maxIndex] - this.selectedCells![minIndex] + 1);
          nextItem[minIndex] += offset;
          nextItem[maxIndex] += offset;
          this.mergedCells.push(nextItem);
        });
      }
    }
  }

  private repeatByCommonCells(
    minIndex: number,
    maxIndex: number,
    times: number,
    d: number
  ) {
    const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
      this.selectedCells!;
    for (let r = minRowIndex; r <= maxRowIndex; r++) {
      for (let c = minColIndex; c <= maxColIndex; c++) {
        const cell = this.cells[r][c];
        for (let i = 0; i < times; i++) {
          const offset =
            (i + 1) *
            d *
            (this.selectedCells![maxIndex] - this.selectedCells![minIndex] + 1);
          const nextRowIndex = minIndex === 0 ? r + offset : r;
          const nextColIndex = minIndex === 0 ? c : c + offset;
          const nextCell = this.cells[nextRowIndex][nextColIndex];
          nextCell.value = cell.value;
          nextCell.meta = cell.meta;
        }
      }
    }
  }

  private repeatCells(minIndex: number, maxIndex: number) {
    const mergedRangesInSelectedCells = this.getMergedRangesInSelectedCells();
    const times =
      (this.fillingCells![maxIndex] - this.fillingCells![minIndex] + 1) /
      (this.selectedCells![maxIndex] - this.selectedCells![minIndex] + 1);
    const d = Math.sign(
      this.fillingCells![minIndex] - this.selectedCells![minIndex]
    );
    this.repeatByCommonCells(minIndex, maxIndex, times, d);
    this.repeatByMergedCells(
      minIndex,
      maxIndex,
      mergedRangesInSelectedCells,
      times,
      d
    );
  }

  private getMergedRangesInSelectedCells() {
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

  private fillCells() {
    const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
      this.selectedCells!;
    const [
      minFillingRowIndex,
      maxFillingRowIndex,
      minFillingColIndex,
      maxFillingColIndex,
    ] = this.fillingCells!;
    if (
      minRowIndex === minFillingRowIndex &&
      maxRowIndex === maxFillingRowIndex
    ) {
      this.repeatCells(2, 3);
    } else if (
      minColIndex === minFillingColIndex &&
      maxColIndex === maxFillingColIndex
    ) {
      this.repeatCells(0, 1);
    }
  }

  private fill(e: MouseEvent) {
    if (!this.pointInFillHandle(e)) return;
    globalObj.EVENT_LOCKED = true;
    this.isFilling = true;
    const onFill = throttle(this.fillingCellRange.bind(this), 50);
    const onEndFill = () => {
      if (this.selectedCells) {
        this.fillCells();
        this.draw();
      }
      globalObj.EVENT_LOCKED = false;
      this.isFilling = false;
      this.clearFillingCells();
      window.removeEventListener("mousemove", onFill);
      window.removeEventListener("mouseup", onEndFill);
    };
    window.addEventListener("mousemove", onFill);
    window.addEventListener("mouseup", onEndFill);
  }

  private getFillingRangeByEndCell(
    endCell: Excel.Cell.CellInstance,
    key: "colIndex" | "rowIndex",
    compareMinIndex: (index: number) => boolean,
    compareMaxIndex: (index: number) => boolean
  ) {
    let minIndex, maxIndex;
    const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
      this.selectedCells!;
    if (key === "rowIndex") {
      minIndex = minRowIndex;
      maxIndex = maxRowIndex;
    } else {
      minIndex = minColIndex;
      maxIndex = maxColIndex;
    }
    const span = maxIndex - minIndex + 1;
    let fillingMinIndex, fillingMaxIndex;
    if (endCell[key]! < minIndex) {
      let cnt = Math.floor((endCell[key]! - minIndex) / span);
      fillingMinIndex = cnt * span + minIndex;
      if (compareMinIndex(fillingMinIndex)) {
        fillingMinIndex = (cnt + 1) * span + minIndex;
      }
      fillingMaxIndex = minIndex - 1;
    } else if (endCell[key]! > maxIndex) {
      let cnt = Math.ceil((endCell[key]! - maxIndex) / span);
      fillingMaxIndex = cnt * span + maxIndex;
      if (compareMaxIndex(fillingMaxIndex)) {
        fillingMaxIndex = (cnt - 1) * span + maxIndex;
      }
      fillingMinIndex = maxIndex + 1;
    }
    if (
      typeof fillingMinIndex === "undefined" ||
      typeof fillingMaxIndex === "undefined"
    ) {
      return null;
    }
    if (fillingMinIndex! > fillingMaxIndex!) {
      return null;
    }
    if (key === "rowIndex") {
      return [
        fillingMinIndex!,
        fillingMaxIndex!,
        minColIndex!,
        maxColIndex!,
      ] as Excel.Sheet.CellRange;
    } else {
      return [
        minRowIndex!,
        maxRowIndex!,
        fillingMinIndex!,
        fillingMaxIndex!,
      ] as Excel.Sheet.CellRange;
    }
  }

  private fillingCellRange(e: MouseEvent) {
    if (this.selectedCells) {
      this.autoScroll(e.x, e.y);
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const endCell = this.findCellByPoint(x, y);
      this.clearFillingCells();
      if (endCell) {
        const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
          this.selectedCells;
        if (
          endCell.rowIndex! >= minRowIndex &&
          endCell.rowIndex! <= maxRowIndex &&
          endCell.colIndex! >= minColIndex &&
          endCell.colIndex! <= maxColIndex
        ) {
          this.fillingCells = null;
        } else {
          const rightBottomCell = this.cells[maxRowIndex][maxColIndex];
          const tanVal =
            (rightBottomCell.position.rightBottom.y - y) /
            (x - rightBottomCell.position.rightBottom.x);
          const deg = Math.atan(tanVal) * (180 / Math.PI);
          if ((deg > 0 && deg > 45) || (deg < 0 && deg < -45)) {
            this.fillingCells = this.getFillingRangeByEndCell(
              endCell,
              "rowIndex",
              (index) => index < this.fixedRowIndex,
              (index) => index > this.cells.length - 1
            );
          } else {
            this.fillingCells = this.getFillingRangeByEndCell(
              endCell,
              "colIndex",
              (index) => index < this.fixedColIndex,
              (index) => index > this.cells[0].length - 1
            );
          }
        }
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
    colIndex = Math.max(colIndex, 0);
    rowIndex = Math.max(rowIndex, 0);
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

  private pointInAbsFixedCell(e: MouseEvent) {
    if (this.pointInCellRange(e)) {
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const cell = this.findCellByPoint(
        x - this.scroll.x,
        y - this.scroll.y,
        false,
        false
      );
      if (cell) {
        return cell.fixed.x && cell.fixed.y;
      }
    }
    return false;
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
    if (this.mode === "view") return;
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
      if (this.pointInAbsFixedCell(e)) {
        globalObj.SET_CURSOR("default");
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

  private checkCellInMergedCells(rowIndex: number, colIndex: number): boolean {
    return this.mergedCells.some((item) => {
      const [startRowIndex, endRowIndex, startColIndex, endColIndex] = item;
      return (
        rowIndex! >= startRowIndex &&
        rowIndex! <= endRowIndex &&
        colIndex! >= startColIndex &&
        colIndex! <= endColIndex
      );
    });
  }

  private truncateContent(
    content: string,
    width: number,
    fontSize: number
  ): string[] {
    const value = String(content || "");

    if (!value) return [];

    const availableWidth = width - DEFAULT_CELL_PADDING * 2;

    if (availableWidth <= 0) return [];

    let currentWidth = 0;
    let result: string[] = [];
    let currentLine = "";

    for (let i = 0; i < value.length; i++) {
      const char = value[i];
      const { width: charWidth } = getTextMetrics(char, fontSize);

      if (currentWidth + charWidth > availableWidth) {
        if (currentLine) {
          result.push(currentLine);
        }

        if (charWidth > availableWidth) {
          result.push(char);
          currentLine = "";
          currentWidth = 0;
        } else {
          currentLine = char;
          currentWidth = charWidth;
        }
      } else {
        currentLine += char;
        currentWidth += charWidth;
      }
    }

    if (currentLine) {
      result.push(currentLine);
    }

    return result;
  }

  private transformCells(
    cells: Excel.Cell.CellInstance[][]
  ): Excel.Cell.CellInstance[][] {
    const transformedCells = cells.map((row) =>
      row.map((cell) => JSON.parse(JSON.stringify(cell)))
    );
    const rowAdjust: Record<number, number> = {};
    for (let rowIndex = 0; rowIndex < transformedCells.length; rowIndex++) {
      const row = transformedCells[rowIndex];
      let currentX = 0;

      for (let colIndex = 0; colIndex < row.length; colIndex++) {
        const cell = row[colIndex];

        cell.x = currentX;
        cell.y =
          rowIndex === 0
            ? 0
            : transformedCells[rowIndex - 1][0].y +
              (transformedCells[rowIndex - 1][0].height || 0);
        const cellInMergedCells = this.checkCellInMergedCells(
          rowIndex,
          colIndex
        );
        if (
          cell.meta?.type === "text" &&
          cell.value &&
          cell.width &&
          !cellInMergedCells
        ) {
          const { width: textWidth } = getTextMetrics(
            cell.value,
            cell.textStyle.fontSize
          );
          if (textWidth > cell.width) {
            if (cell.wrap === "no-wrap") {
              const widthIncrease =
                textWidth - cell.width + DEFAULT_CELL_PADDING * 2;
              cell.width += widthIncrease;

              for (
                let adjustRowIndex = 0;
                adjustRowIndex < transformedCells.length;
                adjustRowIndex++
              ) {
                const adjustRow = transformedCells[adjustRowIndex];
                adjustRow[colIndex].width = cell.width;
                for (
                  let adjustColIndex = colIndex + 1;
                  adjustColIndex < adjustRow.length;
                  adjustColIndex++
                ) {
                  const adjustCell = adjustRow[adjustColIndex];
                  adjustCell.x = (adjustCell.x || 0) + widthIncrease;
                }
              }
            } else if (cell.wrap === "wrap") {
              const fontSize =
                cell.textStyle?.fontSize || DEFAULT_CELL_TEXT_FONT_SIZE;
              const valueSlices = cell.value
                .split("\n")
                .map((item: string) =>
                  this.truncateContent(item, cell.width, fontSize)
                );
              cell.valueSlices = valueSlices.flat();
              const heightIncrease = fontSize * cell.valueSlices.length;
              if (
                !rowAdjust[rowIndex] ||
                rowAdjust[rowIndex] < heightIncrease
              ) {
                let offset = 0;
                if (!rowAdjust[rowIndex]) {
                  offset = heightIncrease;
                } else {
                  offset = heightIncrease - rowAdjust[rowIndex];
                }
                rowAdjust[rowIndex] = heightIncrease;
                transformedCells[rowIndex].forEach((item) => {
                  item.height += offset;
                });
                for (
                  let adjustRowIndex = rowIndex + 1;
                  adjustRowIndex < transformedCells.length;
                  adjustRowIndex++
                ) {
                  const adjustRow = transformedCells[adjustRowIndex];
                  for (
                    let adjustColIndex = 0;
                    adjustColIndex < adjustRow.length;
                    adjustColIndex++
                  ) {
                    const adjustCell = adjustRow[adjustColIndex];
                    adjustCell.y = (adjustCell.y || 0) + offset;
                  }
                }
              }
            }
          }
        }
        currentX += cell.width || 0;
      }
    }
    return transformedCells;
  }

  private transformMergedCells() {
    const rowAdjust: Record<number, number> = {};
    return this.mergedCells.map((range) => {
      const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] = range;
      const leftTopCell = this.getCell(minRowIndex, minColIndex);
      const rightBottomCell = this.getCell(maxRowIndex, maxColIndex);
      const w =
        rightBottomCell.position.rightBottom.x! -
        leftTopCell.position.leftTop.x!;
      if (
        leftTopCell.meta?.type === "text" &&
        leftTopCell.value &&
        leftTopCell.width
      ) {
        const { width: textWidth } = getTextMetrics(
          leftTopCell.value,
          leftTopCell.textStyle.fontSize
        );
        if (textWidth > w) {
          if (leftTopCell.wrap === "no-wrap") {
            const widthIncrease = textWidth - w + DEFAULT_CELL_PADDING * 2;
            leftTopCell.width += widthIncrease;
            leftTopCell.updatePosition();

            for (
              let adjustRowIndex = 0;
              adjustRowIndex < this.cells.length;
              adjustRowIndex++
            ) {
              const adjustRow = this.cells[adjustRowIndex];
              adjustRow[minColIndex].width = leftTopCell.width;
              adjustRow[minColIndex].updatePosition();
              for (
                let adjustColIndex = minColIndex + 1;
                adjustColIndex < adjustRow.length;
                adjustColIndex++
              ) {
                const adjustCell = adjustRow[adjustColIndex];
                adjustCell.x = (adjustCell.x || 0) + widthIncrease;
                adjustCell.updatePosition();
              }
            }
          } else if (leftTopCell.wrap === "wrap") {
            const fontSize =
              leftTopCell.textStyle?.fontSize || DEFAULT_CELL_TEXT_FONT_SIZE;
            const valueSlices = leftTopCell.value
              .split("\n")
              .map((item: string) => this.truncateContent(item, w, fontSize));
            leftTopCell.valueSlices = valueSlices.flat();
            const heightIncrease = fontSize * leftTopCell.valueSlices!.length;
            if (
              !rowAdjust[minRowIndex] ||
              rowAdjust[minRowIndex] < heightIncrease
            ) {
              let offset = 0;
              if (!rowAdjust[minRowIndex]) {
                offset = heightIncrease;
              } else {
                offset = heightIncrease - rowAdjust[minRowIndex];
              }
              rowAdjust[minRowIndex] = heightIncrease;
              this.cells[minRowIndex].forEach((item) => {
                item.height! += offset;
                item.updatePosition();
              });
              for (
                let adjustRowIndex = minRowIndex + 1;
                adjustRowIndex < this.cells.length;
                adjustRowIndex++
              ) {
                const adjustRow = this.cells[adjustRowIndex];
                for (
                  let adjustColIndex = 0;
                  adjustColIndex < adjustRow.length;
                  adjustColIndex++
                ) {
                  const adjustCell = adjustRow[adjustColIndex];
                  adjustCell.y = (adjustCell.y || 0) + offset;
                  adjustCell.updatePosition();
                }
              }
            }
          }
        }
      }
    });
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

  clearCellMeta(cell: Excel.Cell.CellInstance) {
    cell.meta = null;
    cell.value = "";
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
        if (this.mode === "view") return;
        this.fill.call(this, e);
        this.selectCellRange.call(this, e);
        this.resizeCell.call(this, e);
        this.selectFullCell.call(this, e);
        this.endEditingCell.call(this, e);
      },
      dblclick: (e: MouseEvent) => {
        if (this.mode === "view") return;
        this.editCell.call(this, e);
      },
      mousemove: debounce(this.updateCursor.bind(this), 30),
      wheel: this.preventGlobalWheel.bind(this),
      keydown: (e: KeyboardEvent) => {
        if (this.mode === "view") return;
        this.clearCellsMeta.call(this, e);
      },
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
    this.initCellEditor();
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
      this.initConfigCells(cells);
    } else {
      this.initDefaultCells();
    }
    if (this.cells.length > 0) {
      this.transformMergedCells();
      this.realWidth = this.cells[0].reduce((p, c) => p + c.width!, 0);
      this.realHeight = this.cells.reduce((p, c) => p + c[0].height!, 0);
    }
  }

  initConfigCells(cells: Excel.Cell.CellInstance[][]) {
    cells = this.transformCells(cells);
    this.cells = [];
    this.fixedColWidth = 0;
    this.fixedRowHeight = 0;
    let cellXIndex = 0;
    let cellYIndex = 0;
    for (
      let i = 0;
      i < this.rowCount + 1 + (this.mode === "view" ? 1 : 0);
      i++
    ) {
      let row: Excel.Cell.CellInstance[] = [];
      let fixedColRows: Excel.Cell.CellInstance[] = [];
      let fixedRows: Excel.Cell.CellInstance[] = [];
      cellYIndex = i === 0 ? 0 : i - 1;
      if (i === this.rowCount + 1 && this.mode === "view") {
        cellYIndex--;
      }
      for (
        let j = 0;
        j < this.colCount + 1 + (this.mode === "view" ? 1 : 0);
        j++
      ) {
        cellXIndex = j === 0 ? 0 : j - 1;
        const cell = new Cell(this.sheetEventsObserver);
        if (this.mode === "edit") {
          cell.rowIndex = i;
          cell.colIndex = j;
        } else {
          cell.rowIndex = cellYIndex;
          cell.colIndex = cellXIndex;
        }
        if (i > 0 && j > 0) {
          if (j === this.colCount + 1 && this.mode === "view") {
            cell.x = cells[i - 1]?.[j - 2]!.x! + cells[i - 1]?.[j - 2]!.width!;
            cell.y = cells[i - 1]?.[j - 2]!.y!;
            cell.width = this.margin.right;
            cell.height = cells[i - 1]?.[j - 2]!.height!;
            cell.border = {
              top: null,
              bottom: null,
              left: null,
              right: null,
            };
            cell.textStyle.backgroundColor = "";
          } else if (i === this.rowCount + 1 && this.mode === "view") {
            cell.x = cells[i - 2]?.[j - 1]!.x!;
            cell.y = cells[i - 2]?.[j - 1]!.y! + cells[i - 2]?.[j - 1]!.height!;
            cell.width = cells[i - 2]?.[j - 1]!.width!;
            cell.height = this.margin.bottom;
            cell.border = {
              top: null,
              bottom: null,
              left: null,
              right: null,
            };
            cell.textStyle.backgroundColor = "";
          } else {
            Object.assign(cell, cells[cellYIndex]?.[cellXIndex]!);
            cell.x =
              cells[cellYIndex]?.[cellXIndex]!.x! +
              (this.mode === "edit" ? DEFAULT_INDEX_CELL_WIDTH : 0);
            cell.y =
              cells[cellYIndex]?.[cellXIndex]!.y! +
              (this.mode === "edit" ? DEFAULT_CELL_HEIGHT : 0);
          }
        } else {
          if (this.mode === "edit") {
            if (j === 0) {
              cell.x = 0;
              cell.y =
                i === 0
                  ? 0
                  : cells[cellYIndex]?.[cellXIndex]!.y! + DEFAULT_CELL_HEIGHT;
              cell.width = DEFAULT_INDEX_CELL_WIDTH;
              cell.height =
                i === 0
                  ? DEFAULT_CELL_HEIGHT
                  : cells[cellYIndex]?.[cellXIndex]!.height!;
            }
            if (i === 0) {
              cell.x =
                j === 0
                  ? 0
                  : cells[cellYIndex]?.[cellXIndex]!.x! +
                    DEFAULT_INDEX_CELL_WIDTH;
              cell.y = 0;
              cell.width =
                j === 0
                  ? DEFAULT_INDEX_CELL_WIDTH
                  : cells[cellYIndex]?.[cellXIndex]!.width!;
              cell.height = DEFAULT_CELL_HEIGHT;
            }
          }
        }
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
        }
        if (this.mode === "edit" || (i > 0 && j > 0 && this.mode === "view")) {
          row.push(cell);
        }
      }
      this.fixedColCells.push(fixedColRows);
      if (i < this.fixedRowIndex) {
        this.fixedCells.push(fixedRows);
        this.fixedRowCells.push(row);
        this.fixedRowHeight += row[0].height!;
      }
      if (row.length > 0) {
        this.cells.push(row);
      }
    }
  }

  initDefaultCells() {
    this.cells = [];
    this.fixedColWidth = 0;
    this.fixedRowHeight = 0;
    let x = 0;
    let y = 0;
    for (let i = 0; i < this.rowCount + 1; i++) {
      let row: Excel.Cell.CellInstance[] = [];
      y = i * DEFAULT_CELL_HEIGHT;
      x = 0;
      let fixedColRows: Excel.Cell.CellInstance[] = [];
      let fixedRows: Excel.Cell.CellInstance[] = [];
      for (let j = 0; j < this.colCount + 1; j++) {
        x =
          j === 0 ? 0 : (j - 1) * DEFAULT_CELL_WIDTH + DEFAULT_INDEX_CELL_WIDTH;
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

  initCellEditor() {
    this.cellInput = new CellInput(this.layout!);
    this.cellInput.addEvent(
      "input",
      (value: string, cell: Excel.Cell.CellInstance) => {
        this.setCellMeta(
          this.cells[cell.rowIndex!][cell.colIndex!],
          {
            type: "text",
            data: value,
          },
          true
        );
        this.cellInput!.hide();
      }
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
    if (cells.length === 0 || cells[0]?.length === 0) return;
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
        cell.render(this._ctx!, scrollX, scrollY, this.mergedCells);
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
    if (this.verticalScrollBar!.show && this.verticalScrollBar!.percent > 0) {
      this.horizontalScrollBarShadow!.render(this._ctx!);
    }
    if (
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

  drawCellEditor() {
    if (this.cellInput && this.editingCell) {
      this.cellInput.render(
        this.editingCell,
        this.scroll.x || 0,
        this.scroll.y || 0
      );
    }
  }
}

export default Sheet;
