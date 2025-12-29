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

    export type BorderSide = "top" | "bottom" | "left" | "right";

    export type Border = Record<BorderSide, BorderStyle | null>;

    export type Action = "select" | "resize";

    export type CellAction = Record<
      Action,
      {
        x: boolean;
        y: boolean;
        rowIndex: number | null;
        colIndex: number | null;
        value?: number | null;
        mouseX?: number | null;
        mouseY?: number | null;
      }
    >;

    export type CellTextMetaData = string;

    export type CellTextMeta = CellMeta<"text", CellTextMetaData>;

    export type CellImageMetaFill = "fill" | "contain" | "cover" | "none";

    export type CellWrap = "normal" | "wrap" | "no-wrap";

    export type CellImageMetaData = {
      img: CanvasImageSource;
      width: number;
      height: number;
      fit: CellImageMetaFill;
    };

    export type CellImageMeta = CellMeta<"image", CellImageMetaData>;

    export type CellDiagonalMetaData = {
      direction: "top-left" | "top-right" | "bottom-left" | "bottom-right";
      value: string[];
    };

    export type CellDiagonalMeta = CellMeta<"diagonal", CellDiagonalMetaData>;

    export interface CellMeta<T, D> {
      type: T;
      data: D;
      [key: string]: any;
    }

    export type Meta = CellTextMeta | CellImageMeta | CellDiagonalMeta | null;

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
      meta: Meta;
      value: any;
      valueSlices?: string[];
      fn: string | null;
      fixed: CellFixed;
      hidden?: boolean;
      mouseEntered?: boolean;
      wrap?: CellWrap;
      scrollX: number;
      scrollY: number;
      events: Record<string, Array<Excel.Event.FnType>>;
      render(
        ctx: CanvasRenderingContext2D,
        scrollX: number,
        scrollY: number,
        mergedCells: Excel.Sheet.CellRange[]
      ): void;
      updatePosition: Excel.Event.FnType;
      getTextAlignOffsetX(w: number): number;
      setTextStyle(ctx: CanvasRenderingContext2D): void;
    }
  }
}
