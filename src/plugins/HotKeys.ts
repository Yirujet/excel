import hotkeys from "hotkeys-js";
import Element from "../components/Element";

export class HotKeys extends Element<null> {
  declare cells: Excel.Cell.CellInstance[][];
  declare selectedCells: Excel.Sheet.CellRange | null;
  declare clearCellMeta: (cell: Excel.Cell.CellInstance) => void;
  declare setCellMeta: (
    cell: Excel.Cell.CellInstance,
    cellMeta: Excel.Cell.Meta,
    needDraw: boolean
  ) => void;
  declare adjust: () => void;
  constructor() {
    super("");
    this.init();
  }

  init() {
    hotkeys("ctrl+c", () => {
      this.triggerEvent("ctrl+c");
    });

    hotkeys("ctrl+v", () => {
      this.triggerEvent("ctrl+v");
    });
  }

  copy() {
    if (this.selectedCells) {
      const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
        this.selectedCells;
      const selectedCells = this.cells
        .slice(minRowIndex, maxRowIndex + 1)
        .map((row) => row.slice(minColIndex, maxColIndex + 1));
      const copyText = selectedCells
        .map((row) =>
          row.map((cell) => cell.value.replaceAll("\n", "")).join("\t")
        )
        .join("\n");
      navigator.clipboard.writeText(copyText);
    }
  }

  paste() {
    if (this.selectedCells) {
      const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
        this.selectedCells;
      navigator.clipboard.readText().then((text) => {
        const lines = text.split("\n");
        for (let i = 0; i < lines.length; i++) {
          const cells = lines[i].split("\t");
          for (let j = 0; j < cells.length; j++) {
            const cell = this.cells[minRowIndex + i][minColIndex + j];
            this.clearCellMeta(cell);
            this.setCellMeta(
              cell,
              {
                type: "text",
                data: cells[j],
              },
              false
            );
            cell.wrap = "wrap";
          }
        }
        this.adjust();
      });
    }
  }
}
