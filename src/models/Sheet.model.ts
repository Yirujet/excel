namespace Excel {
  export namespace Sheet {
    export interface Configuration {
      fixedRowIndex: number;
      fixedColIndex: number;
      rowCount: number;
      colCount: number;
      cells?: Cell.CellInstance[][];
      mergedCells?: CellRange[];
    }

    export type CellRange = [number, number, number, number];

    export interface SheetInstance extends Required<Configuration> {
      $el: HTMLCanvasElement | null;
      name: string;
      width: number;
      height: number;
      scroll: Excel.PositionPoint;
      fixedRowCells: Cell.CellInstance[][];
      fixedColCells: Cell.CellInstance[][];
      fixedCells: Cell.CellInstance[][];
      selectedCells: CellRange | null;
      render: () => void;
    }
  }
}
