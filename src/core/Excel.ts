/// <reference path="../models/Excel.model.ts" />

import Sheet from "./Sheet/SheetCore";
import Element from "../components/Element";
import {
  DEFAULT_CELL_COL_COUNT,
  DEFAULT_CELL_ROW_COUNT,
} from "../config/index";

class Excel extends Element<HTMLDivElement> implements Excel.ExcelInstance {
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
    this.$el!.style.display = "flex";
    this.$el!.style.flexDirection = "column";
    this.$el!.style.height = "100%";
  }

  mount(target: HTMLElement) {
    this.$target = target;
    this.initSheets();
    this.initSequence();
    const sheetRender = this.createSheetRender();
    this.add(sheetRender.$el!);
    this.$target.appendChild(this.$el!);
  }

  private initSheets() {
    const { width, height, x, y } = this.$target!.getBoundingClientRect();
    if (!this.configuration.sheets?.length) {
      this.addSheet();
    } else {
      this.configuration.sheets.forEach((item) => {
        const sheet = new Sheet(item.name, item);
        sheet.$el!.style.overflow = "hidden";
        sheet.x = x;
        sheet.y = y;
        sheet.width = width;
        sheet.height = height;
        this.sheets.push(sheet);
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
    sheet.$el!.style.overflow = "hidden";
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
    sheetRender.$el!.style.flex = "1";
    const sheet = this.sheets[this.sheetIndex];
    sheet.render(true);
    sheetRender.add(sheet.$el!);
    return sheetRender;
  }
}

export default Excel;
