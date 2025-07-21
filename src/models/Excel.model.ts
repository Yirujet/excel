/// <reference path="./Sheet.model.ts" />
/// <reference path="./Tools.model.ts" />
/// <reference path="./Cell.model.ts" />
/// <reference path="./Event.model.ts" />

namespace Excel {
  export interface ExcelConfiguration {
    name: string;
    cssPrefix?: string;
    sheets?: Sheet.SheetInstance[];
  }

  export interface ExcelInstance {
    $el: HTMLElement | null;
    $target: HTMLElement | null;
    name: string;
    sheets: Sheet.SheetInstance[];
    configuration: ExcelConfiguration;
    sheetIndex: number;
  }

  export interface PositionPoint {
    x: number;
    y: number;
  }

  export interface LayoutInfo extends PositionPoint {
    width: number;
    height: number;
    headerHeight: number;
    fixedLeftWidth: number;
    bodyHeight: number;
    bodyRealWidth: number;
    bodyRealHeight: number;
    deviationCompareValue: number;
  }

  export interface Position {
    leftTop: PositionPoint;
    rightTop: PositionPoint;
    rightBottom: PositionPoint;
    leftBottom: PositionPoint;
  }
}
