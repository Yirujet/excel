import Element from "../../components/Element";
import Excel from "../../core/Excel";

export default class CellTextAlignLeft extends Element {
  constructor() {
    super("div");
    this.render();
  }

  render() {
    this.addClass(`${Excel.CSS_PREFIX}-icon-button`);
    const icon = new Element("i");
    icon.addClass("iconfont");
    icon.addClass("icon-align-left");
    this.add(icon.$el!);
  }
}
