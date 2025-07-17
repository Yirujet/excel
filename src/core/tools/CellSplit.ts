import Element from "../../components/Element";
import Excel from "../../core/Excel";

export default class CellSplit extends Element {
  constructor() {
    super("div");
    this.render();
  }

  render() {
    this.addClass(`${Excel.CSS_PREFIX}-icon-button`);
    const icon = new Element("i");
    icon.addClass("iconfont");
    icon.addClass("icon-split");
    this.add(icon.$el!);
  }
}
