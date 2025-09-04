namespace Excel {
  export namespace Cell {
    export interface TextStyle {
      fontFamily: string;
      fontSize: number;
      bold: boolean;
      italic: boolean;
      underline: boolean;
      backgroundColor: string;
      color: string;
      align: string;
    }

    export interface BorderStyle {
      solid: boolean;
      color: string;
      bold: boolean;
    }

    export interface Style {
      text?: Partial<TextStyle>;
      border?: Partial<BorderStyle>;
    }

    export interface CellFixed {
      x: boolean;
      y: boolean;
    }

    export interface CellSelect {
      x: boolean;
      y: boolean;
      rowIndex: number | null;
      colIndex: number | null;
      value?: number | null;
    }

    export interface CellResize extends CellSelect {}

    export type BorderSide = "top" | "bottom" | "left" | "right";

    export type Border = Record<BorderSide, BorderStyle>;

    export interface CellInstance extends Excel.Event.EventInstance {
      width: number | null;
      height: number | null;
      rowIndex: number | null;
      colIndex: number | null;
      selected?: boolean;
      cellName: string;
      x: number | null;
      y: number | null;
      position: Excel.Position;
      textStyle: TextStyle;
      border: Border;
      meta: any | null;
      value: string;
      fn: string | null;
      fixed: CellFixed;
      hidden?: boolean;
      mouseEntered?: boolean;
      scrollX: number;
      scrollY: number;
      events: Record<string, Array<Excel.Event.FnType>>;
      render(
        ctx: CanvasRenderingContext2D,
        scrollX: number,
        scrollY: number,
        isEnd: boolean
      ): void;
      updatePosition: Excel.Event.FnType;
      getTextAlignOffsetX(w: number): number;
    }
  }
}
