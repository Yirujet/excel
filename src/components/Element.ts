import EventObserver from "../utils/EventObserver";
import ExcelEvent from "../utils/ExcelEvent";

export default class Element extends ExcelEvent {
  $el: HTMLElement | null = null;
  eventObserver: EventObserver = new EventObserver();
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

  addListener(name: string, callback: (...params: any[]) => void) {
    this.$el!.addEventListener(name, callback);
  }

  removeListener(name: string, callback: (...params: any[]) => void) {
    this.$el!.removeEventListener(name, callback);
  }
}
