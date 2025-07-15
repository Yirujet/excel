import Element from "../components/Element";

class Sheet implements Excel.Sheet.SheetInstance {
  $el: HTMLCanvasElement | null = null;
  name = "";
  cells: Excel.Cell.CellInstance[] = [];
  _tools: Excel.Tools.ToolInstance[] = [];
  toolsConfig: Partial<Excel.Sheet.toolsConfig> = {};
  static TOOLS_CONFIG: Excel.Sheet.toolsConfig = {
    cellFontFamily: true,
    cellFontSize: true,
    cellBold: true,
    cellItalic: true,
    cellUnderline: true,
    cellBorder: true,
    cellColor: true,
    cellBackgroundColor: true,
    cellAlign: true,
    cellMerge: true,
    cellSplit: true,
    cellFunction: true,
    cellInsert: true,
    cellDiagonal: true,
    cellFreeze: true,
  };

  constructor(name: string) {
    this.name = name;
    this.render();
  }

  render() {
    const sheet = new Element("canvas");
    sheet.addClass("sheet");
    this.$el = sheet.$el! as HTMLCanvasElement;
  }
}

export default Sheet;
