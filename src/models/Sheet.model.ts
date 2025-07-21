namespace Excel {
  export namespace Sheet {
    type toolTypeUnion = `${Tools.ToolType}`;
    export type toolsConfig = {
      [k in toolTypeUnion as `cell${Capitalize<k>}`]: boolean;
    };

    export interface SheetInstance {
      $el: HTMLElement | HTMLCanvasElement | null;
      name: string;
      width: number;
      height: number;
      toolsConfig: Partial<toolsConfig>;
      _tools: Tools.ToolInstance[];
      cells: Cell.CellInstance[][];
      scroll: Excel.PositionPoint;
      render: () => void;
    }
  }
}
