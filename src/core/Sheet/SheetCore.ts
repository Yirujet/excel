import Element from "../../components/Element";
import EventObserver from "../../utils/EventObserver";
import CellResizer from "../CellResizer";
import HorizontalScrollbar from "../Scrollbar/HorizontalScrollbar";
import VerticalScrollbar from "../Scrollbar/VerticalScrollbar";
import CellSelector from "../CellSelector";
import CellMergence from "../CellMergence";
import Shadow from "../Shadow";
import {
  DEFAULT_CELL_COL_COUNT,
  DEFAULT_CELL_HEIGHT,
  DEFAULT_CELL_ROW_COUNT,
  DEFAULT_CELL_WIDTH,
} from "../../config/index";
import FillHandle from "../FillHandle";
import Filling from "../Filling";
import CellInput from "../CellInput";
import mixin from "../../utils/mixin";
import SheetApi from "./SheetApi";
import SheetHandler from "./SheetHandler";
import SheetRender from "./SheetRender";
import SheetEvent from "./SheetEvent";

class Sheet
  extends Element<HTMLCanvasElement>
  implements Excel.Sheet.SheetInstance
{
  declare render: (autoRegisteEvents: boolean) => void;
  declare initCells: (cells: Excel.Cell.CellInstance[][] | undefined) => void;

  private _ctx: CanvasRenderingContext2D | null = null;
  private _startCell: Excel.Cell.CellInstance | null = null;
  private _animationFrameId: number | null = null;
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
}

mixin(Sheet, [
  { ctor: SheetRender },
  { ctor: SheetEvent },
  { ctor: SheetApi },
  { ctor: SheetHandler, private: true },
]);

export default Sheet;
