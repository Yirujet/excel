import Element from "../../components/Element";
import Select from "../../components/Select";
import Excel from "../../core/Excel";

export default class CellFontFamily extends Element {
  constructor() {
    super("div");
    this.render();
  }

  render() {
    this.addClass(`${Excel.CSS_PREFIX}-cell-font-family`);
    const select = new Select([
      {
        label: "sans-serif",
        value: "sans-serif",
      },
      {
        label: "宋体",
        value: "SimSun",
      },
      {
        label: "黑体",
        value: "SimHei",
      },
      {
        label: "仿宋",
        value: "FangSong",
      },
    ]);
    this.add(select.$el!);
  }
}
