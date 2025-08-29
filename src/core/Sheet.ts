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

class Sheet
  extends Element<HTMLCanvasElement>
  implements Excel.Sheet.SheetInstance
{
  static DEFAULT_CELL_ROW_COUNT = 500;
  static DEFAULT_CELL_COL_COUNT = 1000;
  static DEFAULT_CELL_MIN_WIDTH = 30;
  static DEFAULT_CELL_MIN_HEIGHT = 20;
  static DEFAULT_CELL_WIDTH = 100;
  static DEFAULT_CELL_HEIGHT = 25;
  static DEFAULT_INDEX_CELL_WIDTH = 50;
  static DEFAULT_CELL_FONT_FAMILY = "宋体";
  static DEFAULT_CELL_LINE_DASH = [3, 5];
  static DEFAULT_CELL_LINE_COLOR = "#ccc";
  static DEFAULT_FIXED_CELL_BACKGROUND_COLOR = "rgb(238, 238, 238)";
  static DEFAULT_FIXED_CELL_COLOR = "rgb(141, 87, 87)";
  static DEVIATION_COMPARE_VALUE = 10e-6;
  static DEFAULT_GRADIENT_OFFSET = 6;
  static DEFAULT_GRADIENT_START_COLOR = "rgba(0, 0, 0, 0.12)";
  static DEFAULT_GRADIENT_STOP_COLOR = "transparent";
  static SCROLL_X = 0;
  static SCROLL_Y = 0;
  static RESIZE_ROW_SIZE = 5;
  static RESIZE_COL_SIZE = 10;
  static DEFAULT_RESIZER_LINE_WIDTH = 2;
  static DEFAULT_RESIZER_LINE_DASH = [3, 5];
  static DEFAULT_RESIZER_LINE_COLOR = "#409EFF";
  static DEFAULT_CELL_SELECTED_COLOR = "#409EFF";
  static DEFAULT_CELL_SELECTED_BACKGROUND_COLOR = "rgba(64,158,255, 0.1)";
  private _ctx: CanvasRenderingContext2D | null = null;
  private _startCell: Excel.Cell.CellInstance | null = null;
  name = "";
  cells: Excel.Cell.CellInstance[][] = [];
  width = 0;
  height = 0;
  scroll: Excel.Sheet.ScrollInfo = { x: 0, y: 0 };
  horizontalScrollBar: HorizontalScrollbar | null = null;
  verticalScrollBar: VerticalScrollbar | null = null;
  cellResizer: CellResizer | null = null;
  cellSelector: CellSelector | null = null;
  cellMergence: CellMergence | null = null;
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
  resizeInfo: Excel.Cell.CellResize = {
    x: false,
    y: false,
    rowIndex: null,
    colIndex: null,
    value: null,
  };
  selectedCells: Excel.Sheet.CellRange | null = null;
  mergedCells: Excel.Sheet.CellRange[] = [];

  constructor(name: string, cells?: Excel.Cell.CellInstance[][]) {
    super("canvas");
    this.name = name;
    this.initCells(cells);
  }

  static SET_CURSOR(cursor: string) {
    document.body.style.cursor = cursor;
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
      exceedInfo.y.exceed = false;
      exceedInfo.y.value = 0;
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

  private selectCells(e: MouseEvent) {
    if (this._startCell) {
      const exceedInfo = this.exceed(e.x, e.y);
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
          this.verticalScrollBar!.value +
            this.verticalScrollBar!.track.height <=
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
      const x = Math.max(
        Math.min(e.x - this.layout!.x + (this.scroll.x || 0), this.realWidth),
        0
      );
      const y = Math.max(
        Math.min(e.y - this.layout!.y + (this.scroll.y || 0), this.realHeight),
        0
      );
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
        this.draw(false);
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

  private findCellByPoint(x: number, y: number) {
    let cell = null;
    y = Math.max(y, this.fixedRowHeight);
    x = Math.max(x, this.fixedColWidth);
    let rowIndex = this.cells.findIndex(
      (e) => e[0].position.leftTop.y <= y && e[0].position.leftBottom.y >= y
    );
    let colIndex = this.cells[0].findIndex(
      (e) => e.position.leftTop.x <= x && e.position.rightTop.x >= x
    );
    rowIndex = Math.max(rowIndex, this.fixedRowIndex);
    colIndex = Math.max(colIndex, this.fixedColIndex);
    cell = this.cells[rowIndex][colIndex];
    return cell;
  }

  clearSelectCells() {
    if (this.selectedCells) {
      this.selectedCells = null;
    }
  }

  merge([
    minRowIndex,
    maxRowIndex,
    minColIndex,
    maxColIndex,
  ]: Excel.Sheet.CellRange) {
    this.mergedCells.push([minRowIndex, maxRowIndex, minColIndex, maxColIndex]);
    this.draw(false);
  }

  getCell(rowIndex: number, colIndex: number) {
    return this.cells[rowIndex]?.[colIndex] || null;
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
            const leftSiblingCell = this.getCell(i, j - 1);
            if (leftSiblingCell) {
              leftSiblingCell.border.right = {
                ...leftSiblingCell.border.right,
                ...cellStyle.border,
              };
            }
            const topSiblingCell = this.getCell(i - 1, j);
            if (topSiblingCell) {
              topSiblingCell.border.bottom = {
                ...topSiblingCell.border.bottom,
                ...cellStyle.border,
              };
            }
            const rightSiblingCell = this.getCell(i, j + 1);
            if (rightSiblingCell) {
              rightSiblingCell.border.left = {
                ...rightSiblingCell.border.left,
                ...cellStyle.border,
              };
            }
            const bottomSiblingCell = this.getCell(i + 1, j);
            if (bottomSiblingCell) {
              bottomSiblingCell.border.top = {
                ...bottomSiblingCell.border.top,
                ...cellStyle.border,
              };
            }
          }
        }
      }
    }
    this.draw(true);
  }

  initEvents() {
    const onMousedown = (e: MouseEvent) => {
      if (
        this.verticalScrollBar?.dragging ||
        this.horizontalScrollBar?.dragging
      ) {
        return;
      }
      const exceedInfo = this.exceed(e.x, e.y);
      if (exceedInfo.x.exceed || exceedInfo.y.exceed) {
        return;
      }
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
          return (
            e[0] <= this._startCell!.rowIndex! &&
            e[1] >= this._startCell!.rowIndex! &&
            e[2] <= this._startCell!.colIndex! &&
            e[3] >= this._startCell!.colIndex!
          );
        });
        if (containedMergedCell) {
          this.selectedCells = [...containedMergedCell];
        }
      }
      const onSelectCells = throttle(this.selectCells.bind(this), 50);
      const onEndSelectCells = () => {
        if (this.selectedCells) {
          this.draw(true);
        }
        this._startCell = null;
        window.removeEventListener("mousemove", onSelectCells);
        window.removeEventListener("mouseup", onEndSelectCells);
      };
      window.addEventListener("mousemove", onSelectCells);
      window.addEventListener("mouseup", onEndSelectCells);
    };
    const globalEventListeners = {
      mousedown: onMousedown,
    };

    this.registerListenerFromOnProp(
      globalEventListeners,
      this.globalEventsObserver,
      this
    );
  }

  render() {
    this._ctx = this.$el!.getContext("2d")!;
    this.$el!.style.width = `${this.width}px`;
    this.$el!.style.height = `${this.height}px`;
    this.$el!.width = this.width * window.devicePixelRatio;
    this.$el!.height = this.height * window.devicePixelRatio;
    this._ctx!.translate(0.5, 0.5);
    this._ctx!.scale(window.devicePixelRatio, window.devicePixelRatio);
    this.initScrollbar();
    this.initCellResizer();
    this.initCellSelector();
    this.initCellMergence();
    this.initEvents();
    this.sheetEventsObserver.observe(this.$el!);
    this.globalEventsObserver.observe(window as any);
    this.draw(true);
  }

  initLayout() {
    const bodyHeight = this.height - Sheet.DEFAULT_CELL_HEIGHT;
    const verticalScrollbarShow = bodyHeight < this.realHeight;
    const horizontalScrollbarShow =
      this.realWidth > this.width - Sheet.DEFAULT_INDEX_CELL_WIDTH;
    this.layout = {
      x: this.x,
      y: this.y,
      width:
        this.width -
        (verticalScrollbarShow ? VerticalScrollbar.TRACK_WIDTH : 0),
      height:
        this.height -
        (horizontalScrollbarShow ? HorizontalScrollbar.TRACK_HEIGHT : 0),
      headerHeight: Sheet.DEFAULT_CELL_HEIGHT,
      fixedLeftWidth: Sheet.DEFAULT_INDEX_CELL_WIDTH,
      bodyHeight,
      bodyRealWidth: this.realWidth,
      bodyRealHeight: this.realHeight,
      deviationCompareValue: Sheet.DEVIATION_COMPARE_VALUE,
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
      for (let i = 0; i < Sheet.DEFAULT_CELL_ROW_COUNT + 1; i++) {
        let row: Excel.Cell.CellInstance[] = [];
        y = i * Sheet.DEFAULT_CELL_HEIGHT;
        x = 0;
        let fixedColRows: Excel.Cell.CellInstance[] = [];
        let fixedRows: Excel.Cell.CellInstance[] = [];
        for (let j = 0; j < Sheet.DEFAULT_CELL_COL_COUNT + 1; j++) {
          x =
            j === 0
              ? 0
              : (j - 1) * Sheet.DEFAULT_CELL_WIDTH +
                Sheet.DEFAULT_INDEX_CELL_WIDTH;
          const cell = new Cell(this.sheetEventsObserver);
          cell.x = x;
          cell.y = y;
          cell.width =
            j === 0 ? Sheet.DEFAULT_INDEX_CELL_WIDTH : Sheet.DEFAULT_CELL_WIDTH;
          cell.height = Sheet.DEFAULT_CELL_HEIGHT;
          cell.rowIndex = i;
          cell.colIndex = j;
          cell.cellName = $10226(j - 1);
          cell.updatePosition();
          if (i === 0) {
            cell.value = cell.cellName;
          }
          if (j === 0) {
            cell.value = i.toString();
          }
          if (i === 0 && j === 0) {
            cell.hidden = true;
          }
          if (i > 0 && j > 0) {
            cell.value = i.toString() + "-" + j.toString();
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
            cell.border = {
              top: {
                solid: true,
                color: Sheet.DEFAULT_CELL_LINE_COLOR,
                bold: false,
              },
              bottom: {
                solid: true,
                color: Sheet.DEFAULT_CELL_LINE_COLOR,
                bold: false,
              },
              left: {
                solid: true,
                color: Sheet.DEFAULT_CELL_LINE_COLOR,
                bold: false,
              },
              right: {
                solid: true,
                color: Sheet.DEFAULT_CELL_LINE_COLOR,
                bold: false,
              },
            };
            cell.textStyle.color = Sheet.DEFAULT_FIXED_CELL_COLOR;
            cell.textStyle.backgroundColor =
              Sheet.DEFAULT_FIXED_CELL_BACKGROUND_COLOR;
            cell.textStyle.fontSize = 13;
            cell.textStyle.align = "center";
          } else {
            cell.border = {
              top: {
                solid: false,
                color: Sheet.DEFAULT_CELL_LINE_COLOR,
                bold: false,
              },
              bottom: {
                solid: false,
                color: Sheet.DEFAULT_CELL_LINE_COLOR,
                bold: false,
              },
              left: {
                solid: false,
                color: Sheet.DEFAULT_CELL_LINE_COLOR,
                bold: false,
              },
              right: {
                solid: false,
                color: Sheet.DEFAULT_CELL_LINE_COLOR,
                bold: false,
              },
            };
            cell.textStyle.align = "center";
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
          (this.verticalScrollBar?.show ? VerticalScrollbar.TRACK_WIDTH : 0));
      Sheet.SCROLL_X = this.scroll.x;
    } else {
      this.scroll.y =
        Math.abs(percent) *
        (this.realHeight -
          this.height +
          (this.horizontalScrollBar?.show
            ? HorizontalScrollbar.TRACK_HEIGHT
            : 0));
      Sheet.SCROLL_Y = this.scroll.y;
    }
  }

  handleCellResize(resize: Excel.Cell.CellResize, isEnd = false) {
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
      this.draw(true);
    } else {
      this.draw(false);
    }
  }

  redraw(percent: number, type: Excel.Scrollbar.Type, isEnd: boolean) {
    this.updateScroll(percent, type);
    this.draw(isEnd);
  }

  draw(isEnd: boolean = false) {
    this.drawSheetCells(isEnd);
    this.drawMergedCells();
    this.drawFixedShadow();
    this.drawScrollbar();
    this.drawCellSelector();
    this.drawCellResizer();
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

  drawSheetCells(isEnd: boolean = false) {
    this.sheetEventsObserver.clearEventsWhenReRender();
    this.drawCells(
      this.cells,
      false,
      false,
      this.fixedColIndex,
      this.fixedRowIndex,
      isEnd
    );
    this.drawCells(
      this.fixedRowCells,
      false,
      true,
      this.fixedColIndex,
      null,
      isEnd
    );
    this.drawCells(
      this.fixedColCells,
      true,
      false,
      null,
      this.fixedRowIndex,
      isEnd
    );
    this.drawCells(this.fixedCells, true, true, null, null, isEnd);
  }

  drawCells(
    cells: Excel.Cell.CellInstance[][],
    fixedInX: boolean,
    fixedInY: boolean,
    ignoreXIndex: number | null,
    ignoreYIndex: number | null,
    isEnd: boolean
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
        if (leftTop.x - scrollX < 0 || rightBottom.y - scrollY < 0) {
          continue;
        }
        if (!isEnd) {
          cell.clearEvents!(this.sheetEventsObserver, cell);
        }
        cell.render(this._ctx!, scrollX, scrollY, isEnd);
        if (cell.events["resize"]) {
          cell.events["resize"] = [];
        }
        let index = this.sheetEventsObserver.resize.findIndex(
          (e) => e === cell
        );
        if (!!~index) {
          this.sheetEventsObserver.resize.splice(index, 1);
        }
        cell.addEvent!("resize", this.handleCellResize.bind(this));
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
      this.drawScrollbarCoincide();
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
      this.scroll.x || 0,
      this.scroll.y || 0
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

  drawScrollbarCoincide() {
    this._ctx!.save();
    this._ctx!.strokeStyle = this.verticalScrollBar!.track.borderColor;
    this._ctx!.strokeRect(
      this.verticalScrollBar!.x,
      this.horizontalScrollBar!.y,
      this.verticalScrollBar!.track.width,
      this.horizontalScrollBar!.track.height
    );
    this._ctx!.fillStyle = this.verticalScrollBar!.track.backgroundColor;
    this._ctx!.fillRect(
      this.verticalScrollBar!.x,
      this.horizontalScrollBar!.y,
      this.verticalScrollBar!.track.width,
      this.horizontalScrollBar!.track.height
    );
    this._ctx!.restore();
  }

  drawFixedRowCellsShadow() {
    const gradient = this._ctx!.createLinearGradient(
      this.width / 2,
      this.fixedRowHeight,
      this.width / 2,
      this.fixedRowHeight + Sheet.DEFAULT_GRADIENT_OFFSET
    );
    gradient.addColorStop(0, Sheet.DEFAULT_GRADIENT_START_COLOR);
    gradient.addColorStop(1, Sheet.DEFAULT_GRADIENT_STOP_COLOR);
    this._ctx!.save();
    this._ctx!.fillStyle = gradient;
    this._ctx!.fillRect(
      this.fixedColWidth,
      this.fixedRowHeight,
      this.width,
      Sheet.DEFAULT_GRADIENT_OFFSET
    );
    this._ctx!.restore();
  }

  drawFixedColCellsShadow() {
    const gradient = this._ctx!.createLinearGradient(
      this.fixedColWidth,
      this.height / 2,
      this.fixedColWidth + Sheet.DEFAULT_GRADIENT_OFFSET,
      this.height / 2
    );
    gradient.addColorStop(0, Sheet.DEFAULT_GRADIENT_START_COLOR);
    gradient.addColorStop(1, Sheet.DEFAULT_GRADIENT_STOP_COLOR);
    this._ctx!.save();
    this._ctx!.fillStyle = gradient;
    this._ctx!.fillRect(
      this.fixedColWidth,
      this.fixedRowHeight,
      Sheet.DEFAULT_GRADIENT_OFFSET,
      this.height
    );
    this._ctx!.restore();
  }

  drawFixedShadow() {
    if (
      this.fixedRowCells.length > 0 &&
      this.verticalScrollBar!.show &&
      this.verticalScrollBar!.percent > 0
    ) {
      this.drawFixedRowCellsShadow();
    }
    if (
      this.fixedColCells.length > 0 &&
      this.horizontalScrollBar!.show &&
      this.horizontalScrollBar!.percent > 0
    ) {
      this.drawFixedColCellsShadow();
    }
  }
}

export default Sheet;
