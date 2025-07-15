/// <reference path="../models/Excel.model.ts" />

import Sheet from "./Sheet";
import Element from "../components/Element";
import "../styles/index.less";
import "../assets/fonts/iconfont.css";

class Excel implements Excel.ExcelInstance {
  $el: HTMLElement | null = null;
  name = "";
  cssPrefix = "excel";
  sheets: Excel.Sheet.SheetInstance[] = [];
  configuration!: Excel.ExcelConfiguration;
  sheetIndex = 0;
  private sequence = 0;

  constructor(config: Excel.ExcelConfiguration) {
    this.configuration = config;
    this.name = config.name || `Excel-${Date.now()}`;
    if (config.cssPrefix) {
      this.cssPrefix = config.cssPrefix;
    }
    this.sheets = config.sheets || [];
    this.initSheets();
    this.initSequence();
    this.render();
  }

  render() {
    const excel = new Element("div");
    excel.addClass(`${this.cssPrefix}-wrapper`);
    const sheetManageRender = this.createSheetManageRender();
    const sheetRender = this.createSheetRender();
    const toolsRender = this.createToolsRender();
    excel.add(toolsRender.$el!);
    excel.add(sheetRender.$el!);
    excel.add(sheetManageRender.$el!);
    this.$el = excel.$el!;
  }

  private initSheets() {
    if (this.sheets.length === 0) {
      this.addSheet();
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
    this.sequence = sheetNum;
  }

  private getNextSheetName() {
    const baseName = "Sheet";
    this.sequence++;
    return `${baseName}-${this.sequence}`;
  }

  addSheet() {
    const sheesName = this.getNextSheetName();
    const sheet = new Sheet(sheesName);
    this.sheets.push(sheet);
  }

  selectSheet(index: number) {
    this.sheetIndex = index;
  }

  createSheetManageRender() {
    const sheetManageRender = new Element("div");
    sheetManageRender.addClass(`${this.cssPrefix}-sheet-manage`);
    this.sheets.forEach((e) => {
      const sheet = new Element("div");
      sheet.addClass(`${this.cssPrefix}-sheet`);
      sheet.$el!.innerHTML = e.name;
      sheetManageRender.add(sheet.$el!);
    });
    const addBtn = new Element("div");
    addBtn.addClass(`${this.cssPrefix}-add-btn`);
    addBtn.$el!.innerHTML = "+";
    sheetManageRender.add(addBtn.$el!);
    return sheetManageRender;
  }

  createSheetRender() {
    const sheetRender = new Element("div");
    sheetRender.addClass(`${this.cssPrefix}-sheet-render`);
    const sheet = this.sheets[this.sheetIndex];
    sheetRender.add(sheet.$el!);
    return sheetRender;
  }

  createToolsRender() {
    const toolsRender = new Element("div");
    toolsRender.addClass(`${this.cssPrefix}-tools-render`);
    return toolsRender;
  }
}

export default Excel;
