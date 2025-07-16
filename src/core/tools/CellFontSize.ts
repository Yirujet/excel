import Element from "../../components/Element";
import Select from "../../components/Select";
import Excel from "../../core/Excel";

export default class CellFontSize extends Element {
  constructor() {
    super("div");
    this.render();
  }

  render() {
    this.addClass(`${Excel.CSS_PREFIX}-cell-font-size`);
    const select = new Select(
      Array.from({ length: 99 }).map((_, i) => ({
        label: (i + 1).toString(),
        value: (i + 1).toString(),
      }))
    );
    this.add(select.$el!);
  }
}
