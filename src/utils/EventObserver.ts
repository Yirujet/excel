const GlobalEvents = {
  mouseleave: {
    dispatchEvents: {
      mouseleave: {
        triggerName: "triggerEvent",
      },
    },
  },
  mousemove: {
    dispatchEvents: {
      mousemove: {
        triggerName: "triggerEvent",
      },
      mouseenter: {
        triggerName: "triggerEvent",
      },
      mouseleave: {
        triggerName: "triggerEvent",
      },
    },
  },
  mousedown: {
    dispatchEvents: {
      mousedown: {
        triggerName: "triggerEvent",
      },
    },
  },
  mouseup: {
    dispatchEvents: {
      mouseup: {
        triggerName: "triggerEvent",
      },
    },
  },
  click: {
    dispatchEvents: {
      click: {
        triggerName: "triggerClickEvent",
      },
      clickoutside: {
        triggerName: "triggerEvent",
      },
    },
  },
  wheel: {
    dispatchEvents: {
      wheel: {
        triggerName: "triggerEvent",
      },
    },
  },
};

export default class EventObserver {
  mouseenter: any[] = [];
  mouseleave: any[] = [];
  mousemove: any[] = [];
  mousedown: any[] = [];
  mouseup: any[] = [];
  click: any[] = [];
  clickoutside: any[] = [];
  wheel: any[] = [];
  observe(target: HTMLCanvasElement) {
    if (target && "addEventListener" in target) {
      Object.entries(GlobalEvents).forEach(
        ([targetEvent, { dispatchEvents }]) => {
          const listener = (e: any) => {
            Object.entries(dispatchEvents).forEach(
              ([dispatchEvent, { triggerName }]) => {
                this[
                  dispatchEvent as
                    | "mouseenter"
                    | "mouseleave"
                    | "mousemove"
                    | "mousedown"
                    | "mouseup"
                    | "click"
                    | "clickoutside"
                    | "wheel"
                ].forEach((element) => {
                  element[triggerName].call(element, dispatchEvent, e);
                });
                const curActiveElement = this[
                  dispatchEvent as
                    | "mouseenter"
                    | "mouseleave"
                    | "mousemove"
                    | "mousedown"
                    | "mouseup"
                    | "click"
                    | "clickoutside"
                    | "wheel"
                ].find((element) => element.mouseEntered);
                if (curActiveElement) {
                  target.style.cursor = curActiveElement.cursor;
                } else {
                  target.style.cursor = "default";
                }
              }
            );
          };
          target.addEventListener(targetEvent, listener);
        }
      );
    }
  }
  clear(gcList: any[]) {
    for (let name in this) {
      if (Array.isArray(this[name])) {
        gcList.forEach((item) => {
          let i = (this[name] as any[]).findIndex((e) => e === item);
          if (!!~i) {
            (this[name] as any[]).splice(i, 1);
          }
          if (item && item.destroy) {
            item.destroy();
          }
        });
      }
    }
  }
  clearAll() {
    this.mouseenter = [];
    this.mouseleave = [];
    this.mousemove = [];
    this.mousedown = [];
    this.mouseup = [];
    this.click = [];
    this.clickoutside = [];
    this.wheel = [];
  }
}
