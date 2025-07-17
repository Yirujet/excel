namespace Excel {
  export namespace Cell {
    export interface CellInstance {
      width: number | null;
      height: number | null;
      rowIndex: number | null;
      colIndex: number | null;
      selected?: boolean;
      x: number | null;
      y: number | null;
      textStyle: {
        fontFamily: string;
        fontSize: number;
        bold: boolean;
        italic: boolean;
        underline: boolean;
        backgroundColor: string;
        color: string;
        align: string;
      };
      borderStyle: {
        solid: boolean;
        color: string;
        bold: boolean;
      };
      meta: any | null;
      value: string;
      fn: string | null;
    }
  }
}
