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
  static DEFAULT_CELL_ROW_COUNT = 300;
  static DEFAULT_CELL_COL_COUNT = 300;
  static DEVIATION_COMPARE_VALUE = 10e-6;
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
      let x = 0;
      let y = 0;
      for (let i = 0; i < Sheet.DEFAULT_CELL_ROW_COUNT + 1; i++) {
        let row: Excel.Cell.CellInstance[] = [];
        y = i * Sheet.DEFAULT_CELL_HEIGHT;
        x = 0;
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
          row.push(cell);
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

  redraw(percent: number, type: Excel.Scrollbar.Type) {
    // console.log(type, percent);
    if (type === "horizontal") {
      this.scroll.x = Math.abs(percent) * (this.realWidth - this.width);
    } else {
      this.scroll.y = Math.abs(percent) * (this.realHeight - this.height);
    }
    this.ctx!.clearRect(0, 0, this.width, this.height);
    this.draw();
  }

  draw() {
    this.drawCells();
    this.drawScrollbar();
  }

  drawCells() {
    let minYIndex = this.cells.findIndex(
      (e) => e[0].position.leftBottom.y! - this.scroll.y > 0
    );
    minYIndex = Math.max(minYIndex, 0);
    let maxYIndex = this.cells.findIndex(
      (e) => e[0].position.leftTop.y! - this.scroll.y > this.height
    );
    maxYIndex =
      this.verticalScrollBar?.percent === 1
        ? this.cells.length - 1
        : maxYIndex === -1
        ? this.cells.length - 1
        : maxYIndex;
    let minXIndex = this.cells[0].findIndex(
      (e) => e.position.rightTop.x! - this.scroll.x > 0
    );
    minXIndex = Math.max(minXIndex, 0);
    let maxXIndex = this.cells[0].findIndex(
      (e) => e.position.leftTop.x! - this.scroll.x > this.width
    );
    maxXIndex =
      this.horizontalScrollBar?.percent === 1
        ? this.cells[0].length - 1
        : maxXIndex === -1
        ? this.cells[0].length - 1
        : maxXIndex;
    // console.log(minXIndex, maxXIndex, minYIndex, maxYIndex);
    for (let i = minYIndex; i <= maxYIndex; i++) {
      for (let j = minXIndex; j <= maxXIndex; j++) {
        const cell = this.cells[i][j];
        const {
          position: { leftTop, rightTop, rightBottom, leftBottom },
        } = cell;
        if (
          leftTop.x - this.scroll.x > this.width ||
          leftBottom.y - this.scroll.y > this.height
        ) {
          break;
        }
        if (
          rightTop.x - this.scroll.x < 0 ||
          rightBottom.y - this.scroll.y < 0
        ) {
          continue;
        }
        this.ctx!.fillStyle = "#000";
        this.ctx!.strokeStyle = "#ccc";
        this.ctx!.textBaseline = "middle";
        this.ctx!.textAlign = "center";
        if (i === 0) {
          this.ctx!.strokeRect(
            cell.x! - this.scroll.x,
            cell.y! - this.scroll.y,
            cell.width!,
            cell.height!
          );
          if (j > 0) {
            this.ctx!.fillText(
              cell.cellName,
              cell.x! + cell.width! / 2 - this.scroll.x,
              cell.y! + cell.height! / 2 - this.scroll.y
            );
          }
        } else if (j === 0) {
          this.ctx!.strokeRect(
            cell.x! - this.scroll.x,
            cell.y! - this.scroll.y,
            cell.width!,
            cell.height!
          );
          this.ctx!.fillText(
            i.toString(),
            cell.x! + cell.width! / 2 - this.scroll.x,
            cell.y! + cell.height! / 2 - this.scroll.y
          );
        } else {
          this.ctx!.save();
          this.ctx!.setLineDash([2, 4]);
          this.ctx!.fillText(
            i.toString() + "-" + j.toString(),
            cell.x! + cell.width! / 2 - this.scroll.x,
            cell.y! + cell.height! / 2 - this.scroll.y
          );
          this.ctx!.strokeRect(
            cell.x! - this.scroll.x,
            cell.y! - this.scroll.y,
            cell.width!,
            cell.height!
          );
          this.ctx!.restore();
        }
      }
    }
  }

  drawScrollbar() {
    if (this.verticalScrollBar!.show) {
      this.drawVerticalScrollbar();
    }
    if (this.horizontalScrollBar!.show) {
      this.drawHorizontalScrollbar();
    }
    if (this.verticalScrollBar!.show && this.horizontalScrollBar!.show) {
      this.drawScrollbarCoincide();
    }
  }
  drawVerticalScrollbar() {
    this.ctx!.save();
    this.ctx!.strokeStyle = this.verticalScrollBar!.track.borderColor;
    this.ctx!.strokeRect(
      this.verticalScrollBar!.x,
      this.verticalScrollBar!.y,
      this.verticalScrollBar!.track.width,
      this.verticalScrollBar!.track.height
    );
    this.ctx!.fillStyle = this.verticalScrollBar!.track.backgroundColor;
    this.ctx!.fillRect(
      this.verticalScrollBar!.x,
      this.verticalScrollBar!.y,
      this.verticalScrollBar!.track.width,
      this.verticalScrollBar!.track.height
    );
    this.ctx!.fillStyle = this.verticalScrollBar!.dragging
      ? this.verticalScrollBar!.thumb.draggingColor
      : this.verticalScrollBar!.thumb.backgroundColor;
    this.ctx!.fillRect(
      this.verticalScrollBar!.x + this.verticalScrollBar!.thumb.padding,
      this.verticalScrollBar!.y - this.verticalScrollBar!.value,
      this.verticalScrollBar!.thumb.width,
      this.verticalScrollBar!.thumb.height
    );
    this.ctx!.restore();
  }
  drawHorizontalScrollbar() {
    this.ctx!.save();
    this.ctx!.strokeStyle = this.horizontalScrollBar!.track.borderColor;
    this.ctx!.strokeRect(
      this.horizontalScrollBar!.x,
      this.horizontalScrollBar!.y,
      this.horizontalScrollBar!.track.width,
      this.horizontalScrollBar!.track.height
    );
    this.ctx!.fillStyle = this.horizontalScrollBar!.track.backgroundColor;
    this.ctx!.fillRect(
      this.horizontalScrollBar!.x,
      this.horizontalScrollBar!.y,
      this.horizontalScrollBar!.track.width,
      this.horizontalScrollBar!.track.height
    );
    this.ctx!.fillStyle = this.horizontalScrollBar!.dragging
      ? this.horizontalScrollBar!.thumb.draggingColor
      : this.horizontalScrollBar!.thumb.backgroundColor;
    this.ctx!.fillRect(
      this.horizontalScrollBar!.x - this.horizontalScrollBar!.value,
      this.horizontalScrollBar!.y + this.horizontalScrollBar!.thumb.padding,
      this.horizontalScrollBar!.thumb.width,
      this.horizontalScrollBar!.thumb.height
    );
    this.ctx!.restore();
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
}

export default Sheet;
