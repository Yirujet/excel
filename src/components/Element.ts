import ExcelEvent from "../utils/ExcelEvent";

export default class Element<T extends HTMLElement | null> extends ExcelEvent {
  $el: T | null = null;
  x!: number;
  y!: number;
  mouseEntered = false;
  constructor(tagName: string, clearEventsWhenReRender = false) {
    super(clearEventsWhenReRender);
    if (tagName) {
      this.$el = document.createElement(tagName) as T;
    }
  }

  add(el: HTMLElement) {
    this.$el!.appendChild(el);
  }

  addClass(className: string) {
    this.$el!.classList.add(className);
  }

  addListener(name: string, callback: Excel.Event.FnType) {
    this.$el!.addEventListener(name, callback);
  }

  removeListener(name: string, callback: Excel.Event.FnType) {
    this.$el!.removeEventListener(name, callback);
  }
}
