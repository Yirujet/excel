/// <reference path="../models/Excel.model.ts" />

import Sheet from "./Sheet";
import Element from "../components/Element";
import "../styles/index.less";
import "../assets/fonts/iconfont.css";
import Tool from "./Tool";

class Excel extends Element implements Excel.ExcelInstance {
  static CSS_PREFIX = "excel";
  private _sequence = 0;
  $target: HTMLElement | null = null;
  name = "";
  sheets: Excel.Sheet.SheetInstance[] = [];
  configuration!: Excel.ExcelConfiguration;
  sheetIndex = 0;

  constructor(config: Excel.ExcelConfiguration) {
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
    const sheetManageRender = this.createSheetManageRender();
    const sheetRender = this.createSheetRender();
    // const toolsRender = this.createToolsRender();
    // this.add(toolsRender.$el!);
    this.add(sheetRender.$el!);
    this.add(sheetManageRender.$el!);
    this.$target.appendChild(this.$el!);
  }

  private initSheets() {
    const { width, height, x, y } = this.$target!.getBoundingClientRect();
    if (this.sheets.length === 0) {
      this.addSheet(width, height, x, y);
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

  addSheet(width: number, height: number, x: number, y: number) {
    const sheesName = this.getNextSheetName();
    const sheet = new Sheet(sheesName);
    sheet.x = x;
    sheet.y = y;
    sheet.width = width;
    sheet.height = height;
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
    sheet.render();
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
        tool.addEvent("cell-text1-bold", () => {
          console.log("****excel -> tool-text-bold triggered");
        });
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
      if (sheet.toolsConfig.cellBackgroundColor) {
        const tool = new Tool("backgroundColor" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellColor) {
        const tool = new Tool("color" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellAlign) {
        const tool = new Tool("align" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellMerge) {
        const tool = new Tool("merge" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellSplit) {
        const tool = new Tool("split" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellFunction) {
        const tool = new Tool("function" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellInsert) {
        const tool = new Tool("insert" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellDiagonal) {
        const tool = new Tool("diagonal" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
      if (sheet.toolsConfig.cellFreeze) {
        const tool = new Tool("freeze" as Excel.Tools.ToolType);
        tool.addClass(`${Excel.CSS_PREFIX}-tool`);
        toolsRender.add(tool.$el!);
      }
    }
    return toolsRender;
  }
}

export default Excel;
