import Element from "../../components/Element";
import Excel from "../../core/Excel";

export default class CellTextBold extends Element {
  constructor() {
    super("div");
    this.render();
    this.initEvents();
  }

  render() {
    this.addClass(`${Excel.CSS_PREFIX}-icon-button`);
    const icon = new Element("i");
    icon.addClass("iconfont");
    icon.addClass("icon-bold");
    this.add(icon.$el!);
  }

  initEvents() {
    this.addListener("click", () => {
      this.triggerEvent("click");
    });
  }
}
