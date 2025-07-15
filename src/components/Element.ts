export default class Element {
  $el: HTMLElement | null = null;
  constructor(tagName: string) {
    this.$el = document.createElement(tagName);
  }

  add(el: HTMLElement) {
    this.$el!.appendChild(el);
  }

  addClass(className: string) {
    this.$el!.classList.add(className);
  }
}
