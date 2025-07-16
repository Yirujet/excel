import ExcelEvent from "../utils/ExcelEvent";

export default class Element extends ExcelEvent {
  $el: HTMLElement | null = null;
  constructor(tagName: string) {
    super();
    this.$el = document.createElement(tagName);
  }

  add(el: HTMLElement) {
    this.$el!.appendChild(el);
  }

  addClass(className: string) {
    this.$el!.classList.add(className);
  }
}
