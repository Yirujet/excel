import Element from "../components/Element";
import CellBackgroundColor from "./tools/CellBackgroundColor";
import CellBorder from "./tools/CellBorder";
import CellDiagonal from "./tools/CellDiagonal";
import CellFontFamily from "./tools/CellFontFamily";
import CellFontSize from "./tools/CellFontSize";
import CellFreeze from "./tools/CellFreeze";
import CellFunction from "./tools/CellFunction";
import CellInsert from "./tools/CellInsert";
import CellMerge from "./tools/CellMerge";
import CellSplit from "./tools/CellSplit";
import CellTextAlignCenter from "./tools/CellTextAlignCenter";
import CellTextAlignLeft from "./tools/CellTextAlignLeft";
import CellTextAlignRight from "./tools/CellTextAlignRight";
import CellTextBold from "./tools/CellTextBold";
import CellTextColor from "./tools/CellTextColor";
import CellTextItalic from "./tools/CellTextItalic";
import CellTextUnderline from "./tools/CellTextUnderline";

class Tool extends Element implements Excel.Tools.ToolInstance {
  type!: Excel.Tools.ToolType;
  disabled = false;

  constructor(type: Excel.Tools.ToolType) {
    super("div");
    this.type = type;
    this.render();
  }

  render() {
    switch (this.type) {
      case "fontFamily":
        const cellFontFamilyTool = new CellFontFamily();
        cellFontFamilyTool.addEvent("click", () => {
          console.log("****fontFamily triggered");
        });
        this.add(cellFontFamilyTool.$el!);
        break;
      case "fontSize":
        this.add(new CellFontSize().$el!);
        break;
      case "bold":
        const cellTextBoldTool = new CellTextBold();
        cellTextBoldTool.addEvent("click", () => {
          console.log("****tool:text-bold triggered");
          this.triggerEvent("cell-text-bold");
        });
        this.add(cellTextBoldTool.$el!);
        break;
      case "italic":
        this.add(new CellTextItalic().$el!);
        break;
      case "underline":
        this.add(new CellTextUnderline().$el!);
        break;
      case "border":
        this.add(new CellBorder().$el!);
        break;
      case "backgroundColor":
        this.add(new CellBackgroundColor().$el!);
        break;
      case "color":
        this.add(new CellTextColor().$el!);
        break;
      case "align":
        this.add(new CellTextAlignLeft().$el!);
        this.add(new CellTextAlignCenter().$el!);
        this.add(new CellTextAlignRight().$el!);
        break;
      case "merge":
        this.add(new CellMerge().$el!);
        break;
      case "split":
        this.add(new CellSplit().$el!);
        break;
      case "function":
        this.add(new CellFunction().$el!);
        break;
      case "insert":
        this.add(new CellInsert().$el!);
        break;
      case "diagonal":
        this.add(new CellDiagonal().$el!);
        break;
      case "freeze":
        this.add(new CellFreeze().$el!);
        break;
    }
  }
}

export default Tool;
