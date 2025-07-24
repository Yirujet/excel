import Element from "../components/Element";
import $10226 from "../utils/10226";
import EventObserver from "../utils/EventObserver";
import Cell from "./Cell";
import HorizontalScrollbar from "./Scrollbar/HorizontalScrollbar";
import VerticalScrollbar from "./Scrollbar/VerticalScrollbar";

class Sheet extends Element implements Excel.Sheet.SheetInstance {
  static TOOLS_CONFIG: Excel.Sheet.toolsConfig = {
    cellFontFamily: true,
    cellFontSize: true,
    cellBold: true,
    cellItalic: true,
    cellUnderline: true,
    cellBorder: true,
    cellColor: true,
    cellBackgroundColor: true,
    cellAlign: true,
    cellMerge: true,
    cellSplit: true,
    cellFunction: true,
    cellInsert: true,
    cellDiagonal: true,
    cellFreeze: true,
  };
  static DEFAULT_CELL_WIDTH = 100;
  static DEFAULT_CELL_HEIGHT = 25;
  static DEFAULT_INDEX_CELL_WIDTH = 50;
  static DEFAULT_CELL_FONT_FAMILY = "宋体";
  static DEFAULT_CELL_ROW_COUNT = 50;
  static DEFAULT_CELL_COL_COUNT = 50;
  static DEVIATION_COMPARE_VALUE = 10e-6;
  static DEFAULT_GRADIENT_OFFSET = 6;
  static DEFAULT_GRADIENT_START_COLOR = "rgba(0, 0, 0, 0.12)";
  static DEFAULT_GRADIENT_STOP_COLOR = "transparent";
  private ctx: CanvasRenderingContext2D | null = null;
  name = "";
  cells: Excel.Cell.CellInstance[][] = [];
  _tools: Excel.Tools.ToolInstance[] = [];
  toolsConfig: Partial<Excel.Sheet.toolsConfig> = {};
  width = 0;
  height = 0;
  scroll: { x: number; y: number } = { x: 0, y: 0 };
  horizontalScrollBar: HorizontalScrollbar | null = null;
  verticalScrollBar: VerticalScrollbar | null = null;
  sheetEventsObserver: EventObserver = new EventObserver();
  globalEventsObserver: EventObserver = new EventObserver();
  realWidth = 0;
  realHeight = 0;
  fixedRowIndex = 1;
  fixedColIndex = 1;
  fixedRowCells: Excel.Cell.CellInstance[][] = [];
  fixedColCells: Excel.Cell.CellInstance[][] = [];
  fixedCells: Excel.Cell.CellInstance[][] = [];
  fixedRowHeight = 0;
  fixedColWidth = 0;

  constructor(
    name: string,
    toolsConfig?: Partial<Excel.Sheet.toolsConfig>,
    cells?: Excel.Cell.CellInstance[][]
  ) {
    super("canvas");
    this.name = name;
    this.initToolConfig(toolsConfig);
    this.initCells(cells);
  }

  initToolConfig(toolsConfig?: Partial<Excel.Sheet.toolsConfig>) {
    if (toolsConfig) {
      this.toolsConfig = toolsConfig;
    } else {
      this.toolsConfig = Sheet.TOOLS_CONFIG;
    }
  }

  render() {
    this.ctx = (this.$el as HTMLCanvasElement).getContext("2d")!;
    (this.$el as HTMLCanvasElement).style.width = `${this.width}px`;
    (this.$el as HTMLCanvasElement).style.height = `${this.height}px`;
    (this.$el as HTMLCanvasElement).width =
      this.width * window.devicePixelRatio;
    (this.$el as HTMLCanvasElement).height =
      this.height * window.devicePixelRatio;
    this.ctx!.translate(0.5, 0.5);
    this.ctx!.scale(window.devicePixelRatio, window.devicePixelRatio);
    this.initScrollbar();
    this.sheetEventsObserver.observe(this.$el as HTMLCanvasElement);
    this.globalEventsObserver.observe(window as any);
    this.draw();
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
          const cell = new Cell();
          cell.x = x;
          cell.y = y;
          cell.width =
            j === 0 ? Sheet.DEFAULT_INDEX_CELL_WIDTH : Sheet.DEFAULT_CELL_WIDTH;
          cell.height = Sheet.DEFAULT_CELL_HEIGHT;
          cell.rowIndex = i;
          cell.colIndex = j;
          cell.cellName = $10226(j - 1);
          cell.position = {
            leftTop: {
              x,
              y,
            },
            rightTop: {
              x: x + cell.width,
              y,
            },
            rightBottom: {
              x: x + cell.width,
              y: y + cell.height,
            },
            leftBottom: {
              x,
              y: y + cell.height,
            },
          };
          if (i === 0) {
            cell.value = cell.cellName;
          }
          if (j === 0) {
            cell.value = i.toString();
          }
          if (i > 0 && j > 0) {
            cell.value = i.toString() + "-" + j.toString();
          }
          if (i < this.fixedRowIndex) {
            cell.fixed = true;
          }
          if (j < this.fixedColIndex) {
            cell.fixed = true;
            fixedColRows.push(cell);
            if (i < this.fixedRowIndex) {
              fixedRows.push(cell);
            }
            if (i === 0) {
              this.fixedColWidth += cell.width!;
            }
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

  initScrollbar() {
    const layout = {
      x: this.x,
      y: this.y,
      width: this.width - VerticalScrollbar.TRACK_WIDTH,
      height: this.height - HorizontalScrollbar.TRACK_HEIGHT,
      headerHeight: Sheet.DEFAULT_CELL_HEIGHT,
      fixedLeftWidth: Sheet.DEFAULT_INDEX_CELL_WIDTH,
      bodyHeight: this.height - Sheet.DEFAULT_CELL_HEIGHT,
      bodyRealWidth: this.realWidth,
      bodyRealHeight: this.realHeight,
      deviationCompareValue: Sheet.DEVIATION_COMPARE_VALUE,
    };
    this.horizontalScrollBar = new HorizontalScrollbar(
      layout,
      this.sheetEventsObserver,
      this.globalEventsObserver,
      this.redraw.bind(this)
    );
    this.verticalScrollBar = new VerticalScrollbar(
      layout,
      this.sheetEventsObserver,
      this.globalEventsObserver,
      this.redraw.bind(this)
    );
  }

  clear() {
    this.ctx!.clearRect(0, 0, this.width, this.height);
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
    this.ctx!.clearRect(0, 0, w, h);
  }

  updateScroll(percent: number, type: Excel.Scrollbar.Type) {
    if (type === "horizontal") {
      this.scroll.x =
        Math.abs(percent) *
        (this.realWidth -
          this.width +
          (this.verticalScrollBar?.show ? VerticalScrollbar.TRACK_WIDTH : 0));
    } else {
      this.scroll.y =
        Math.abs(percent) *
        (this.realHeight -
          this.height +
          (this.horizontalScrollBar?.show
            ? HorizontalScrollbar.TRACK_HEIGHT
            : 0));
    }
  }

  redraw(percent: number, type: Excel.Scrollbar.Type) {
    this.updateScroll(percent, type);
    this.draw();
  }

  draw() {
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
    this.drawFixedShadow();
    this.drawScrollbar();
  }
  getRangeInView(
    cells: Excel.Cell.CellInstance[][],
    scrollX: number,
    scrollY: number
  ) {
    let minYIndex = cells.findIndex(
      (e) => e[0].position.leftBottom.y! - scrollY > 0
    );
    minYIndex = Math.max(minYIndex, 0);
    let maxYIndex = cells.findIndex(
      (e) => e[0].position.leftTop.y! - scrollY > this.height
    );
    maxYIndex =
      this.verticalScrollBar?.percent === 1
        ? cells.length - 1
        : maxYIndex === -1
        ? cells.length - 1
        : maxYIndex;
    let minXIndex = cells[0].findIndex(
      (e) => e.position.rightTop.x! - scrollX > 0
    );
    minXIndex = Math.max(minXIndex, 0);
    let maxXIndex = cells[0].findIndex(
      (e) => e.position.leftTop.x! - scrollX > this.width
    );
    maxXIndex =
      this.horizontalScrollBar?.percent === 1
        ? cells[0].length - 1
        : maxXIndex === -1
        ? cells[0].length - 1
        : maxXIndex;
    return [minXIndex, maxXIndex, minYIndex, maxYIndex];
  }
  drawCells(
    cells: Excel.Cell.CellInstance[][],
    fixedInX: boolean,
    fixedInY: boolean,
    ignoreXIndex: number | null,
    ignoreYIndex: number | null
  ) {
    this.clearCells(fixedInX, fixedInY);
    const scrollX = fixedInX ? 0 : this.scroll.x;
    const scrollY = fixedInY ? 0 : this.scroll.y;
    const [minXIndex, maxXIndex, minYIndex, maxYIndex] = this.getRangeInView(
      cells,
      scrollX,
      scrollY
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
        if (rightTop.x - scrollX < 0 || rightBottom.y - scrollY < 0) {
          continue;
        }
        cell.render(this.ctx!, scrollX, scrollY);
      }
    }
  }
  drawScrollbar() {
    if (this.verticalScrollBar!.show) {
      this.verticalScrollBar!.render(this.ctx!);
    }
    if (this.horizontalScrollBar!.show) {
      this.horizontalScrollBar!.render(this.ctx!);
    }
    if (this.verticalScrollBar!.show && this.horizontalScrollBar!.show) {
      this.drawScrollbarCoincide();
    }
  }
  drawScrollbarCoincide() {
    this.ctx!.save();
    this.ctx!.strokeStyle = this.verticalScrollBar!.track.borderColor;
    this.ctx!.strokeRect(
      this.verticalScrollBar!.x,
      this.horizontalScrollBar!.y,
      this.verticalScrollBar!.track.width,
      this.horizontalScrollBar!.track.height
    );
    this.ctx!.fillStyle = this.verticalScrollBar!.track.backgroundColor;
    this.ctx!.fillRect(
      this.verticalScrollBar!.x,
      this.horizontalScrollBar!.y,
      this.verticalScrollBar!.track.width,
      this.horizontalScrollBar!.track.height
    );
    this.ctx!.restore();
  }
  drawFixedRowCellsShadow() {
    const gradient = this.ctx!.createLinearGradient(
      this.width / 2,
      this.fixedRowHeight,
      this.width / 2,
      this.fixedRowHeight + Sheet.DEFAULT_GRADIENT_OFFSET
    );
    gradient.addColorStop(0, Sheet.DEFAULT_GRADIENT_START_COLOR);
    gradient.addColorStop(1, Sheet.DEFAULT_GRADIENT_STOP_COLOR);
    this.ctx!.save();
    this.ctx!.fillStyle = gradient;
    this.ctx!.fillRect(
      0,
      this.fixedRowHeight,
      this.width,
      Sheet.DEFAULT_GRADIENT_OFFSET
    );
    this.ctx!.restore();
  }
  drawFixedColCellsShadow() {
    const gradient = this.ctx!.createLinearGradient(
      this.fixedColWidth,
      this.height / 2,
      this.fixedColWidth + Sheet.DEFAULT_GRADIENT_OFFSET,
      this.height / 2
    );
    gradient.addColorStop(0, Sheet.DEFAULT_GRADIENT_START_COLOR);
    gradient.addColorStop(1, Sheet.DEFAULT_GRADIENT_STOP_COLOR);
    this.ctx!.save();
    this.ctx!.fillStyle = gradient;
    this.ctx!.fillRect(
      this.fixedColWidth,
      0,
      Sheet.DEFAULT_GRADIENT_OFFSET,
      this.height
    );
    this.ctx!.restore();
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
