import { HotKeys } from "../../plugins/HotKeys";
import CellResizer from "../../plugins/CellResizer";

export default abstract class SheetPlugin {
  declare layout: Excel.LayoutInfo | null;
  declare hotKeys: HotKeys | null;
  declare cellResizer: CellResizer | null;

  initPlugins(plugins: Excel.Sheet.PluginType[]) {
    if (plugins.includes("hotkeys")) {
      this.initHotKeysPlugin();
    }
    if (plugins.includes("resize")) {
      this.initCellResizer();
    }
  }

  initHotKeysPlugin() {
    this.hotKeys = new HotKeys();

    this.hotKeys.addEvent("ctrl+c", this.hotKeys.copy.bind(this));
    this.hotKeys.addEvent("ctrl+v", this.hotKeys.paste.bind(this));
    this.hotKeys.addEvent("ctrl+x", this.hotKeys.cut.bind(this));
  }

  initCellResizer() {
    this.cellResizer = new CellResizer(this.layout!);
  }
}
