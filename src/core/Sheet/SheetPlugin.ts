import { HotKeys } from "../../plugins/HotKeys";

export default abstract class SheetPlugin {
  declare cells: Excel.Cell.CellInstance[][];
  declare hotKeys: HotKeys | null;
  declare selectedCells: Excel.Sheet.CellRange | null;
  declare clearCellMeta: (cell: Excel.Cell.CellInstance) => void;
  declare setCellMeta: (
    cell: Excel.Cell.CellInstance,
    cellMeta: Excel.Cell.Meta,
    needDraw: boolean
  ) => void;
  declare adjust: () => void;

  initPlugins(plugins: Excel.Sheet.PluginType[]) {
    if (plugins.includes("hotkeys")) {
      this.initHotKeysPlugin();
    }
  }

  initHotKeysPlugin() {
    this.hotKeys = new HotKeys();

    this.hotKeys.addEvent("ctrl+c", this.copy.bind(this));
    this.hotKeys.addEvent("ctrl+v", this.paste.bind(this));
  }

  private copy() {
    if (this.selectedCells) {
      const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] =
        this.selectedCells;
      const selectedCells = this.cells
        .slice(minRowIndex, maxRowIndex + 1)
        .map((row) => row.slice(minColIndex, maxColIndex + 1));
      const copyText = selectedCells
        .map((row) => row.map((cell) => cell.value).join("\t"))
        .join("\n");
      navigator.clipboard.writeText(copyText);
    }
  }

  private paste() {
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
