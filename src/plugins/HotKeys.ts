import hotkeys from "hotkeys-js";
import Element from "../components/Element";

export class HotKeys extends Element<null> {
  constructor() {
    super("");
    this.init();
  }

  init() {
    hotkeys("ctrl+c", () => {
      this.triggerEvent("ctrl+c");
    });

    hotkeys("ctrl+v", () => {
      this.triggerEvent("ctrl+v");
    });
  }
}
