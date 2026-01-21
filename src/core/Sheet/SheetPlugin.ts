import { HotKeys } from "../../plugins/HotKeys";
import CellResizer from "../../plugins/CellResizer";
import { Formula } from "../../plugins/Formula";

export default abstract class SheetPlugin {
  declare cells: Excel.Cell.CellInstance[][];
  declare layout: Excel.LayoutInfo | null;
  declare hotKeys: HotKeys | null;
  declare cellResizer: CellResizer | null;
  declare formula: Formula | null;

  initPlugins(plugins: Excel.Sheet.PluginType[]) {
    if (plugins.includes("hotkeys")) {
      this.initHotKeysPlugin();
    }
    if (plugins.includes("resize")) {
      this.initCellResizer();
    }
    if (plugins.includes("formula")) {
      this.initFormulaPlugin();
    }
  }

  /**
   * 初始化热键插件
   */
  initHotKeysPlugin() {
    this.hotKeys = new HotKeys();

    this.hotKeys.addEvent("ctrl+c", this.hotKeys.copy.bind(this));
    this.hotKeys.addEvent("ctrl+v", this.hotKeys.paste.bind(this));
    this.hotKeys.addEvent("ctrl+x", this.hotKeys.cut.bind(this));
  }

  /**
   * 初始化单元格调整大小插件
   */
  initCellResizer() {
    this.cellResizer = new CellResizer(this.layout!);
  }

  /**
   * 初始化公式插件
   */
  initFormulaPlugin() {
    this.formula = new Formula(this.cells);
  }
}
