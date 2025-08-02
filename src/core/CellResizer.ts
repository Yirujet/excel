import Element from "../components/Element";

class CellResizer extends Element {
  eventObserver: Excel.Event.ObserverInstance;
  constructor(eventObserver: Excel.Event.ObserverInstance) {
    super("");
    this.eventObserver = eventObserver;
  }

  init() {
    this.initEvents();
  }

  initEvents() {
    const onMouseDown = (e: MouseEvent) => {
      e.preventDefault();
      console.log("***", e);
    };
    const defaultEventListeners = {
      mousedown: onMouseDown,
    };
    this.registerListenerFromOnProp(
      defaultEventListeners,
      this.eventObserver,
      this as any
    );
  }
}

export default CellResizer;
