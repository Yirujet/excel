/// <reference path="../models/Excel.model.ts" />

class Excel implements Excel.ExcelInstance {
  name = "";
  sheets: Array<Excel.Sheet.SheetInstance> = [];
  tools: Array<Excel.Tools.ToolInstance> = [];
  cells: Array<Excel.Cell.CellInstance> = [];
  constructor() {}
}

export default Excel;
