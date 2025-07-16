/// <reference path="../models/Excel.model.ts" />

import Sheet from "./Sheet";
import Element from "../components/Element";
import "../styles/index.less";
import "../assets/fonts/iconfont.css";
import Tool from "./Tool";

class Excel implements Excel.ExcelInstance {
  static CSS_PREFIX = "excel";
  private sequence = 0;

  $el: HTMLElement | null = null;
  name = "";
  sheets: Excel.Sheet.SheetInstance[] = [];
  configuration!: Excel.ExcelConfiguration;
  sheetIndex = 0;

  constructor(config: Excel.ExcelConfiguration) {
    this.configuration = config;
    this.name = config.name || `Excel-${Date.now()}`;
    if (config.cssPrefix) {
      Excel.CSS_PREFIX = config.cssPrefix;
    }
    this.sheets = config.sheets || [];
    this.initSheets();
    this.initSequence();
    this.render();
  }

  render() {
    const excel = new Element("div");
    excel.addClass(`${Excel.CSS_PREFIX}-wrapper`);
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
    sheetManageRender.addClass(`${Excel.CSS_PREFIX}-sheet-manage`);
    this.sheets.forEach((e) => {
      const sheet = new Element("div");
      sheet.addClass(`${Excel.CSS_PREFIX}-sheet`);
      sheet.$el!.innerHTML = e.name;
      sheetManageRender.add(sheet.$el!);
    });
    const addBtn = new Element("div");
    addBtn.addClass(`${Excel.CSS_PREFIX}-add-btn`);
    addBtn.$el!.innerHTML = "+";
    sheetManageRender.add(addBtn.$el!);
    return sheetManageRender;
  }

  createSheetRender() {
    const sheetRender = new Element("div");
    sheetRender.addClass(`${Excel.CSS_PREFIX}-sheet-render`);
    const sheet = this.sheets[this.sheetIndex];
    sheetRender.add(sheet.$el!);
    return sheetRender;
  }

  createToolsRender() {
    const toolsRender = new Element("div");
    toolsRender.addClass(`${Excel.CSS_PREFIX}-tools-render`);
    const sheet = this.sheets[this.sheetIndex];
    if (sheet.toolsConfig) {
      if (sheet.toolsConfig.cellFontFamily) {
        const tool = new Tool("fontFamily" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellFontSize) {
        const tool = new Tool("fontSize" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellBold) {
        const tool = new Tool("bold" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellItalic) {
        const tool = new Tool("italic" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellUnderline) {
        const tool = new Tool("underline" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellBorder) {
        const tool = new Tool("border" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
    }
    return toolsRender;
  }
}

export default Excel;
