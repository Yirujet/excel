import Element from "../components/Element";
import CellBorder from "./tools/CellBorder";
import CellFontFamily from "./tools/CellFontFamily";
import CellFontSize from "./tools/CellFontSize";
import CellTextBold from "./tools/CellTextBold";
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
        this.add(new CellFontFamily().$el!);
        break;
      case "fontSize":
        this.add(new CellFontSize().$el!);
        break;
      case "bold":
        this.add(new CellTextBold().$el!);
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
    }
  }
}

export default Tool;
