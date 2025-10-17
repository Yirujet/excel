/// <reference path="../models/Excel.model.ts" />

import Sheet from "./Sheet";
import Element from "../components/Element";
import "../styles/index.less";
import {
  DEFAULT_CELL_COL_COUNT,
  DEFAULT_CELL_ROW_COUNT,
} from "../config/index";

class Excel extends Element<HTMLDivElement> implements Excel.ExcelInstance {
  static CSS_PREFIX = "excel";
  private _sequence = 0;
  $target: HTMLElement | null = null;
  name = "";
  sheets: Excel.Sheet.SheetInstance[] = [];
  configuration!: Excel.Configuration;
  sheetIndex = 0;

  constructor(config: Excel.Configuration) {
    super("div");
    this.configuration = config;
    this.name = config.name || `Excel-${Date.now()}`;
    this.addClass(`${Excel.CSS_PREFIX}-wrapper`);
    if (config.cssPrefix) {
      Excel.CSS_PREFIX = config.cssPrefix;
    }
  }

  mount(target: HTMLElement) {
    this.$target = target;
    this.sheets = this.configuration.sheets || [];
    this.initSheets();
    this.initSequence();
    const sheetRender = this.createSheetRender();
    this.add(sheetRender.$el!);
    this.$target.appendChild(this.$el!);
  }

  private initSheets() {
    const { width, height, x, y } = this.$target!.getBoundingClientRect();
    if (this.sheets.length === 0) {
      this.addSheet();
    } else {
      this.sheets.forEach((item) => {
        item.width = width;
        item.height = height;
      });
    }
  }

  private initSequence() {
    const sheetNum =
      this.sheets.length > 0
        ? Math.max(
            ...this.sheets.map(
              (e) => parseInt(e.name.replace("Sheet-", "")) || 0
            )
          )
        : 0;
    this._sequence = sheetNum;
  }

  private getNextSheetName() {
    const baseName = "Sheet";
    this._sequence++;
    return `${baseName}-${this._sequence}`;
  }

  addSheet(name?: string, config?: Excel.Sheet.Configuration) {
    const { width, height, x, y } = this.$target!.getBoundingClientRect();
    const sheesName = name || this.getNextSheetName();
    const sheetConfig = config || {
      fixedRowIndex: 1,
      fixedColIndex: 1,
      rowCount: DEFAULT_CELL_ROW_COUNT,
      colCount: DEFAULT_CELL_COL_COUNT,
    };
    const sheet = new Sheet(sheesName, sheetConfig);
    sheet.x = x;
    sheet.y = y;
    sheet.width = width;
    sheet.height = height;
    this.sheets.push(sheet);
  }

  selectSheet(index: number) {
    this.sheetIndex = index;
  }

  createSheetRender() {
    const sheetRender = new Element<HTMLDivElement>("div");
    sheetRender.addClass(`${Excel.CSS_PREFIX}-sheet-render`);
    const sheet = this.sheets[this.sheetIndex];
    sheet.render();
    sheetRender.add(sheet.$el!);
    return sheetRender;
  }
}

export default Excel;
