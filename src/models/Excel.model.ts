/// <reference path="./Sheet.model.ts" />
/// <reference path="./Tools.model.ts" />
/// <reference path="./Cell.model.ts" />

namespace Excel {
  export interface ExcelInstance {
    name: string;
    sheets: Array<Sheet.SheetInstance>;
    tools: Array<Tools.ToolInstance>;
    cells: Array<Cell.CellInstance>;
  }
}
