import debounce from "../../utils/debounce";
import throttle from "../../utils/throttle";
import globalObj from "../globalObj";
import HorizontalScrollbar from "../Scrollbar/HorizontalScrollbar";
import VerticalScrollbar from "../Scrollbar/VerticalScrollbar";

export default abstract class SheetEvent {
  declare x: number;
  declare y: number;
  declare width: number;
  declare height: number;
  declare cells: Excel.Cell.CellInstance[][];
  declare mode: Excel.Sheet.Mode;
  declare isFilling: boolean;
  declare selectedCells: Excel.Sheet.CellRange | null;
  declare fillingCells: Excel.Sheet.CellRange | null;
  declare mergedCells: Excel.Sheet.CellRange[];
  declare layout: Excel.LayoutInfo | null;
  declare fixedColWidth: number;
  declare fixedRowHeight: number;
  declare scroll: Excel.PositionPoint;
  declare horizontalScrollBar: HorizontalScrollbar | null;
  declare verticalScrollBar: VerticalScrollbar | null;
  declare fixedColIndex: number;
  declare fixedRowIndex: number;
  declare resizeInfo: Excel.Cell.CellAction["resize"];
  declare selectInfo: Excel.Cell.CellAction["select"];
  declare realWidth: number;
  declare realHeight: number;
  declare editingCell: Excel.Cell.CellInstance | null;
  declare globalEventsObserver: Excel.Event.ObserverInstance;
  private declare _startCell: Excel.Cell.CellInstance | null;
  private declare pointInFillHandle: (e: MouseEvent) => boolean;
  declare getMergedRangesInSelectedCells: () => Excel.Sheet.CellRange[];
  declare draw: () => void;
  declare clearFillingCells: () => void;
  declare findCellByPoint: (
    x: number,
    y: number,
    ignoreFixedX?: boolean,
    ignoreFixedY?: boolean
  ) => Excel.Cell.CellInstance | null;
  declare clearSelectCells: () => void;
  declare getCellPointByMousePosition: (
    mouseX: number,
    mouseY: number
  ) => Excel.PositionPoint;
  declare updateScroll: (percent: number, type: Excel.Scrollbar.Type) => void;
  private declare pointInAbsFixedCell: (e: MouseEvent) => boolean;
  private declare pointInColResize: (e: MouseEvent) => boolean;
  private declare pointInRowResize: (e: MouseEvent) => boolean;
  private declare pointInFixedCell: (e: MouseEvent) => boolean;
  private declare pointInFixedXCell: (e: MouseEvent) => boolean;
  private declare pointInFixedYCell: (e: MouseEvent) => boolean;
  declare drawCellEditor: () => void;
  private declare pointInCellRange: (e: MouseEvent) => boolean;
  private declare pointInScrollbar: (e: MouseEvent) => boolean;
  private declare pointInNormalCell: (e: MouseEvent) => boolean;
  declare clearCellMeta: (cell: Excel.Cell.CellInstance) => void;
  declare registerListenerFromOnProp: (
    onObj: {
      [k in Excel.Event.Type]?: Excel.Event.FnType;
    },
    eventObserver: Excel.Event.ObserverInstance,
    obj: Excel.Event.ObserverTypes
  ) => void;

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
        rowIndex: cell!.rowIndex,
        colIndex: cell!.colIndex,
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
          rowIndex: cell!.rowIndex,
          colIndex: cell!.colIndex,
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

  private endEditingCell(e: MouseEvent) {
    const x = e.x - this.layout!.x + (this.scroll.x || 0);
    const y = e.y - this.layout!.y + (this.scroll.y || 0);
    const cell = this.findCellByPoint(x, y);
    if (cell === this.editingCell) return;
    this.editingCell = null;
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
        minRowIndex <= cell!.rowIndex! &&
        maxRowIndex >= cell!.rowIndex! &&
        minColIndex <= cell!.colIndex! &&
        maxColIndex >= cell!.colIndex!
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
      this as any
    );
  }
}
