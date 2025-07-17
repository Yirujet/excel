import Element from "../components/Element";
import cellIndex2CellWord from "../utils/cellIndex2CellWord";

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
  static DEFAULT_CELL_ROW_COUNT = 100;
  static DEFAULT_CELL_COL_COUNT = 100;
  name = "";
  cells: Excel.Cell.CellInstance[] = [];
  _tools: Excel.Tools.ToolInstance[] = [];
  toolsConfig: Partial<Excel.Sheet.toolsConfig> = {};
  width = 0;
  height = 0;
  scroll: { x: number; y: number } = { x: 0, y: 0 };

  constructor(name: string, toolsConfig?: Partial<Excel.Sheet.toolsConfig>) {
    super("canvas");
    this.name = name;
    this.initToolConfig(toolsConfig);
  }

  initToolConfig(toolsConfig?: Partial<Excel.Sheet.toolsConfig>) {
    if (toolsConfig) {
      this.toolsConfig = toolsConfig;
    } else {
      this.toolsConfig = Sheet.TOOLS_CONFIG;
    }
  }

  render() {
    const ctx = (this.$el as HTMLCanvasElement).getContext("2d")!;
    (this.$el as HTMLCanvasElement).style.width = `${this.width}px`;
    (this.$el as HTMLCanvasElement).style.height = `${this.height}px`;
    (this.$el as HTMLCanvasElement).width =
      this.width * window.devicePixelRatio;
    (this.$el as HTMLCanvasElement).height =
      this.height * window.devicePixelRatio;
    ctx.translate(0.5, 0.5);
    ctx.scale(window.devicePixelRatio, window.devicePixelRatio);
    this.createCells(ctx);
  }

  createCells(ctx: CanvasRenderingContext2D) {
    for (let i = 0; i < Sheet.DEFAULT_CELL_ROW_COUNT + 1; i++) {
      for (let j = 0; j < Sheet.DEFAULT_CELL_COL_COUNT + 1; j++) {
        ctx.fillStyle = "#000";
        ctx.strokeStyle = "#ccc";
        ctx.textBaseline = "middle";
        ctx.textAlign = "center";
        if (i === 0) {
          if (j > 0) {
            ctx.strokeRect(
              Sheet.DEFAULT_INDEX_CELL_WIDTH +
                (j - 1) * Sheet.DEFAULT_CELL_WIDTH,
              i * Sheet.DEFAULT_CELL_HEIGHT,
              Sheet.DEFAULT_CELL_WIDTH,
              Sheet.DEFAULT_CELL_HEIGHT
            );
          } else {
            ctx.strokeRect(
              j * Sheet.DEFAULT_INDEX_CELL_WIDTH,
              i * Sheet.DEFAULT_CELL_HEIGHT,
              Sheet.DEFAULT_INDEX_CELL_WIDTH,
              Sheet.DEFAULT_CELL_HEIGHT
            );
          }
          if (j > 0) {
            const text = cellIndex2CellWord(j);
            ctx.fillText(
              text,
              Sheet.DEFAULT_INDEX_CELL_WIDTH +
                (j - 1) * Sheet.DEFAULT_CELL_WIDTH +
                Sheet.DEFAULT_CELL_WIDTH / 2,
              i * Sheet.DEFAULT_CELL_HEIGHT + Sheet.DEFAULT_CELL_HEIGHT / 2
            );
          }
        } else if (j === 0) {
          ctx.strokeRect(
            j * Sheet.DEFAULT_INDEX_CELL_WIDTH,
            i * Sheet.DEFAULT_CELL_HEIGHT,
            Sheet.DEFAULT_INDEX_CELL_WIDTH,
            Sheet.DEFAULT_CELL_HEIGHT
          );
          ctx.fillText(
            i.toString(),
            j * Sheet.DEFAULT_INDEX_CELL_WIDTH +
              Sheet.DEFAULT_INDEX_CELL_WIDTH / 2,
            i * Sheet.DEFAULT_CELL_HEIGHT + Sheet.DEFAULT_CELL_HEIGHT / 2
          );
        } else {
          ctx.save();
          ctx.setLineDash([2, 4]);
          ctx.fillText(
            i.toString() + "-" + j.toString(),
            Sheet.DEFAULT_INDEX_CELL_WIDTH +
              (j - 1) * Sheet.DEFAULT_CELL_WIDTH +
              Sheet.DEFAULT_CELL_WIDTH / 2,
            i * Sheet.DEFAULT_CELL_HEIGHT + Sheet.DEFAULT_CELL_HEIGHT / 2
          );
          ctx.strokeRect(
            Sheet.DEFAULT_INDEX_CELL_WIDTH + (j - 1) * Sheet.DEFAULT_CELL_WIDTH,
            i * Sheet.DEFAULT_CELL_HEIGHT,
            Sheet.DEFAULT_CELL_WIDTH,
            Sheet.DEFAULT_CELL_HEIGHT
          );
          ctx.restore();
        }
      }
    }
  }
}

export default Sheet;
