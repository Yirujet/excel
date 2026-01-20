import Element from "../components/Element";
import { Coordinate, Parser as FormulaParser } from "hot-formula-parser";
import globalObj from "../core/globalObj";

export class Formula extends Element<null> {
  parser: FormulaParser | null = null;
  cells: Excel.Cell.CellInstance[][] = [];
  constructor(cells: Excel.Cell.CellInstance[][]) {
    super("");
    this.cells = cells;
    this.init();
  }

  init() {
    this.parser = new FormulaParser();
    globalObj.FORMULA_PARSER = this.parser;
    this.addVariableEvent();
  }

  addVariableEvent() {
    this.parser!.on(
      "callCellValue",
      (
        cellCoord: { row: Coordinate; label: string; column: Coordinate },
        done: (value: any) => void,
      ) => {
        done(
          this.cells[cellCoord.row.index + 1][cellCoord.column.index + 1].value,
        );
      },
    );
  }

  parse(expression: string) {
    if (!this.parser) {
      console.error("Parser not initialized");
      return null;
    }

    try {
      const result = this.parser.parse(expression);
      return result;
    } catch (error) {
      console.error("Parse error:", error);
      return null;
    }
  }
}
