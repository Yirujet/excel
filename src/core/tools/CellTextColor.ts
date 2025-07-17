import Element from "../../components/Element";
import Excel from "../../core/Excel";

export default class CellTextColor extends Element {
  constructor() {
    super("div");
    this.render();
  }

  render() {
    this.addClass(`${Excel.CSS_PREFIX}-cell-border`);
    this.addClass(`${Excel.CSS_PREFIX}-icon-button`);
    const icon = new Element("i");
    icon.addClass("iconfont");
    icon.addClass("icon-color");
    this.add(icon.$el!);
    const arrowDownIcon = new Element("i");
    arrowDownIcon.addClass("iconfont");
    arrowDownIcon.addClass("icon-arrow-down");
    this.add(arrowDownIcon.$el!);
  }
}
