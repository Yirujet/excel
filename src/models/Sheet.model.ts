namespace Excel {
  export namespace Sheet {
    export type CellRange = [number, number, number, number];

    export interface SheetInstance {
      $el: HTMLCanvasElement | null;
      name: string;
      width: number;
      height: number;
      cells: Cell.CellInstance[][];
      scroll: Excel.PositionPoint;
      fixedRowIndex: number;
      fixedColIndex: number;
      fixedRowCells: Cell.CellInstance[][];
      fixedColCells: Cell.CellInstance[][];
      fixedCells: Cell.CellInstance[][];
      selectedCells: CellRange | null;
      mergedCells: CellRange[];
      render: () => void;
    }
  }
}
