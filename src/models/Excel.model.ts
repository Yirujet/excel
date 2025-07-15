/// <reference path="./Sheet.model.ts" />
/// <reference path="./Tools.model.ts" />
/// <reference path="./Cell.model.ts" />

namespace Excel {
  export interface ExcelConfiguration {
    name: string;
    cssPrefix?: string;
    sheets?: Sheet.SheetInstance[];
  }

  export interface ExcelInstance {
    $el: HTMLElement | null;
    name: string;
    cssPrefix: string;
    sheets: Sheet.SheetInstance[];
    configuration: ExcelConfiguration;
    sheetIndex: number;
  }
}
