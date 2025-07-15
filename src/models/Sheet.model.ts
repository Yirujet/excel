namespace Excel {
  export namespace Sheet {
    type toolTypeUnion = `${Tools.ToolType}`;
    export type toolsConfig = {
      [k in toolTypeUnion as `cell${Capitalize<k>}`]: boolean;
    };

    export interface SheetInstance {
      $el: HTMLCanvasElement | null;
      name: string;
      toolsConfig: Partial<toolsConfig>;
      _tools: Tools.ToolInstance[];
      cells: Cell.CellInstance[];
    }
  }
}
