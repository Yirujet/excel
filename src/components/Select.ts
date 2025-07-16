import Element from "./Element";
import Excel from "../core/Excel";

interface SelectOption {
  value: string;
  label: string;
}

export default class Select extends Element {
  options: SelectOption[] = [];
  constructor(options: SelectOption[]) {
    super("select");
    this.options = options;
    this.render();
  }
  render() {
    this.addClass(`${Excel.CSS_PREFIX}-select`);
    this.$el?.setAttribute("is", "ui-select");
    this.options.forEach((e) => {
      const option = new Element("option");
      (option.$el as HTMLOptionElement)!.value = e.value;
      (option.$el as HTMLOptionElement)!.innerHTML = e.label;
      this.$el!.appendChild(option.$el!);
    });
  }
}
