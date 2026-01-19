import {
  DEFAULT_CELL_HEIGHT,
  DEFAULT_CELL_LINE_COLOR,
  DEFAULT_CELL_WIDTH,
  DEFAULT_FIXED_CELL_BACKGROUND_COLOR,
  DEFAULT_FIXED_CELL_COLOR,
  DEFAULT_GRADIENT_OFFSET,
  DEFAULT_GRADIENT_START_COLOR,
  DEFAULT_GRADIENT_STOP_COLOR,
  DEFAULT_INDEX_CELL_WIDTH,
  DEFAULT_SCROLLBAR_TRACK_SIZE,
  DEVIATION_COMPARE_VALUE,
} from "../../config";
import Cell from "../Cell";
import CellInput from "../CellInput";
import CellMergence from "../CellMergence";
import CellResizer from "../../plugins/CellResizer";
import CellSelector from "../CellSelector";
import FillHandle from "../FillHandle";
import Filling from "../Filling";
import globalObj from "../globalObj";
import HorizontalScrollbar from "../Scrollbar/HorizontalScrollbar";
import VerticalScrollbar from "../Scrollbar/VerticalScrollbar";
import Shadow from "../Shadow";
import $10226 from "../../utils/10226";

export default abstract class SheetRender {
  declare private _ctx: CanvasRenderingContext2D | null;
  declare sheetEventsObserver: Excel.Event.ObserverInstance;
  declare globalEventsObserver: Excel.Event.ObserverInstance;
  declare $el: HTMLCanvasElement | null;
  declare mode: Excel.Sheet.Mode;
  declare rowCount: number;
  declare colCount: number;
  declare width: number;
  declare height: number;
  declare realWidth: number;
  declare realHeight: number;
  declare layout: Excel.LayoutInfo | null;
  declare x: number;
  declare y: number;
  declare horizontalScrollBar: HorizontalScrollbar | null;
  declare verticalScrollBar: VerticalScrollbar | null;
  declare horizontalScrollBarShadow: Shadow | null;
  declare verticalScrollBarShadow: Shadow | null;
  declare fixedColWidth: number;
  declare fixedRowHeight: number;
  declare cellResizer: CellResizer | null;
  declare cellSelector: CellSelector | null;
  declare cells: Excel.Cell.CellInstance[][];
  declare cellMergence: CellMergence | null;
  declare fillHandle: FillHandle | null;
  declare filling: Filling | null;
  declare cellInput: CellInput | null;
  declare scroll: Excel.PositionPoint;
  declare fixedColIndex: number;
  declare fixedRowIndex: number;
  declare fixedRowCells: Excel.Cell.CellInstance[][];
  declare fixedColCells: Excel.Cell.CellInstance[][];
  declare fixedCells: Excel.Cell.CellInstance[][];
  declare mergedCells: Excel.Sheet.CellRange[];
  declare selectedCells: Excel.Sheet.CellRange | null;
  declare private _startCell: Excel.Cell.CellInstance | null;
  declare fillingCells: Excel.Sheet.CellRange | null;
  declare editingCell: Excel.Cell.CellInstance | null;
  declare margin: Exclude<Excel.Sheet.Configuration["margin"], undefined>;
  declare plugins: Excel.Sheet.PluginType[];
  declare private _animationFrameId: number | null;
  declare private _redrawTimer: number | null;
  declare initEvents: () => void;
  declare setCellMeta: (
    cell: Excel.Cell.CellInstance,
    cellMeta: Excel.Cell.Meta,
    needDraw: boolean,
  ) => void;
  declare clearCells: (fixedInX: boolean, fixedInY: boolean) => void;
  declare setCellStyle: (
    cell: Excel.Cell.CellInstance,
    cellStyle: Excel.Cell.Style,
  ) => void;
  declare private transformMergedCells: () => void;
  declare private transformCells: (
    cells: Excel.Cell.CellInstance[][],
  ) => Excel.Cell.CellInstance[][];
  declare initPlugins: (plugins: Excel.Sheet.PluginType[]) => void;

  render(autoRegisteEvents: boolean = true) {
    this.initSheet();
    this.initScrollbar();
    this.initShadow();
    this.initCellSelector();
    this.initCellMergence();
    this.initFillHandle();
    this.initFilling();
    this.initCellEditor();
    if (autoRegisteEvents) {
      this.initPlugin();
      this.initEvents();
      this.sheetEventsObserver.observe(this.$el!);
      this.globalEventsObserver.observe(window as any);
    }
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

  initScrollbar() {
    this.initLayout();
    this.horizontalScrollBar = new HorizontalScrollbar(
      this.layout!,
      this.sheetEventsObserver,
      this.globalEventsObserver,
    );
    this.verticalScrollBar = new VerticalScrollbar(
      this.layout!,
      this.sheetEventsObserver,
      this.globalEventsObserver,
    );
    this.horizontalScrollBar.addEvent("percent", this.redraw.bind(this));
    this.verticalScrollBar.addEvent("percent", this.redraw.bind(this));
    this.autoScrollToCurrentView();
  }

  autoScrollToCurrentView() {
    if (this.scroll) {
      const horizontablePercent =
        this.scroll.x /
        (this.realWidth -
          this.width +
          (this.verticalScrollBar?.show ? DEFAULT_SCROLLBAR_TRACK_SIZE : 0));
      const verticalPercent =
        this.scroll.y /
        (this.realHeight -
          this.height +
          (this.horizontalScrollBar?.show ? DEFAULT_SCROLLBAR_TRACK_SIZE : 0));
      this.updateScroll(horizontablePercent, "horizontal");
      this.updateScroll(verticalPercent, "vertical");
      this.horizontalScrollBar?.scrollTo(horizontablePercent);
      this.verticalScrollBar?.scrollTo(verticalPercent);
    }
  }

  initShadow() {
    this.horizontalScrollBarShadow = new Shadow(
      this.fixedColWidth,
      this.fixedRowHeight,
      this.width,
      DEFAULT_GRADIENT_OFFSET,
      [DEFAULT_GRADIENT_START_COLOR, DEFAULT_GRADIENT_STOP_COLOR],
      "vertical",
    );
    this.verticalScrollBarShadow = new Shadow(
      this.fixedColWidth,
      this.fixedRowHeight,
      DEFAULT_GRADIENT_OFFSET,
      this.height,
      [DEFAULT_GRADIENT_START_COLOR, DEFAULT_GRADIENT_STOP_COLOR],
      "horizontal",
    );
  }

  initCellSelector() {
    this.cellSelector = new CellSelector(
      this.layout!,
      this.cells,
      this.fixedColWidth,
      this.fixedRowHeight,
    );
  }

  initCellMergence() {
    this.cellMergence = new CellMergence(
      this.layout!,
      this.cells,
      this.fixedColWidth,
      this.fixedRowHeight,
    );
  }

  initFillHandle() {
    this.fillHandle = new FillHandle(
      this.sheetEventsObserver,
      this.layout!,
      this.cells,
      this.fixedColWidth,
      this.fixedRowHeight,
    );
  }

  initFilling() {
    this.filling = new Filling(
      this.layout!,
      this.cells,
      this.fixedColWidth,
      this.fixedRowHeight,
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
          true,
        );
        this.cellInput!.hide();
      },
    );
  }

  initPlugin() {
    this.initPlugins(this.plugins);
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
      cellYIndex = i > 0 ? i - 1 : 0;
      if (i === this.rowCount + 1 && this.mode === "view") {
        cellYIndex--;
      }
      for (
        let j = 0;
        j < this.colCount + 1 + (this.mode === "view" ? 1 : 0);
        j++
      ) {
        cellXIndex = j > 0 ? j - 1 : 0;
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
            false,
          );
        }
        if (j === 0) {
          this.setCellMeta(
            cell,
            {
              type: "text",
              data: i.toString(),
            },
            false,
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
            false,
          );
        }
        if (j === 0) {
          this.setCellMeta(
            cell,
            {
              type: "text",
              data: i.toString(),
            },
            false,
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

  private binaryQuery<T>(
    arr: T[],
    value: number,
    compare: (_binaryIndex: number, value: number, arr: T[]) => boolean,
    complete: (_binaryIndex: number, arr: T[]) => boolean,
    judge?: (_binaryIndex: number, value: number, arr: T[]) => boolean,
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

  getRangeInView(
    cells: Excel.Cell.CellInstance[][],
    scrollX: number,
    scrollY: number,
    fixedInX: boolean,
    fixedInY: boolean,
  ) {
    let minYIndex = this.binaryQuery(
      cells,
      0,
      (binaryIndex, value, arr) => {
        return arr[binaryIndex][0].position.leftBottom.y! - scrollY > value;
      },
      (binaryIndex, arr) => {
        return binaryIndex === 0;
      },
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
        },
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
      },
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
        },
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

  redraw(percent: number, type: Excel.Scrollbar.Type) {
    this.updateScroll(percent, type);

    if (this._redrawTimer) {
      clearTimeout(this._redrawTimer);
    }

    this._redrawTimer = window.setTimeout(() => {
      this.draw();
      this._redrawTimer = null;
    }, 16);
  }

  draw() {
    if (this._animationFrameId) {
      cancelAnimationFrame(this._animationFrameId);
    }
    this._animationFrameId = requestAnimationFrame(() => {
      this.drawSheetCells();
      this.drawMergedCells();
      this.drawShadow();
      this.drawScrollbar();
      this.drawCellSelector();
      this.drawCellResizer();
      this.drawFillHandle();
      this.drawFilling();
    });
  }

  drawSheetCells() {
    this.sheetEventsObserver.clearEventsWhenReRender();
    this.drawCells(
      this.cells,
      false,
      false,
      this.fixedColIndex,
      this.fixedRowIndex,
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
    ignoreYIndex: number | null,
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
      fixedInY,
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
          leftTop.y - scrollY > this.height ||
          rightBottom.x - scrollX < 0 ||
          rightBottom.y - scrollY < 0
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
          (e) => e === cell,
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
    if (!this.cellResizer) return;
    if (this.cellResizer.resizeInfo.x || this.cellResizer.resizeInfo.y) {
      let cellInfo: Excel.Cell.CellInstance =
        this.cells[this.cellResizer.resizeInfo.rowIndex!][
          this.cellResizer.resizeInfo.colIndex!
        ];
      this.cellResizer!.render(this._ctx!, cellInfo, this.scroll);
    }
  }

  drawCellSelector() {
    this.cellSelector!.render(
      this._ctx!,
      this.selectedCells,
      this._startCell,
      this.scroll.x || 0,
      this.scroll.y || 0,
      this.mergedCells,
    );
  }

  drawMergedCells() {
    this.cellMergence!.render(
      this._ctx!,
      this.mergedCells,
      this.scroll.x || 0,
      this.scroll.y || 0,
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
      this.scroll.y || 0,
    );
  }

  drawFilling() {
    this.filling!.render(
      this._ctx!,
      this.fillingCells,
      this.scroll.x || 0,
      this.scroll.y || 0,
    );
  }

  drawCellEditor() {
    if (this.cellInput && this.editingCell) {
      this.cellInput.render(
        this.editingCell,
        this.scroll.x || 0,
        this.scroll.y || 0,
      );
    }
  }
}
