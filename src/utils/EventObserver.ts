const GlobalEvents: Excel.Event.GlobalEvent = {
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
  keydown: {
    dispatchEvents: {
      keydown: {
        triggerName: "triggerEvent",
      },
    },
  },
  keyup: {
    dispatchEvents: {
      keyup: {
        triggerName: "triggerEvent",
      },
    },
  },
  resize: {
    dispatchEvents: {
      resize: {
        triggerName: "triggerEvent",
      },
    },
  },
};

export default class EventObserver implements Excel.Event.ObserverInstance {
  mouseenter: Excel.Event.ObserverTypes[] = [];
  mouseleave: Excel.Event.ObserverTypes[] = [];
  mousemove: Excel.Event.ObserverTypes[] = [];
  mousedown: Excel.Event.ObserverTypes[] = [];
  mouseup: Excel.Event.ObserverTypes[] = [];
  click: Excel.Event.ObserverTypes[] = [];
  clickoutside: Excel.Event.ObserverTypes[] = [];
  wheel: Excel.Event.ObserverTypes[] = [];
  keydown: Excel.Event.ObserverTypes[] = [];
  keyup: Excel.Event.ObserverTypes[] = [];
  resize: Excel.Event.ObserverTypes[] = [];
  select: Excel.Event.ObserverTypes[] = [];
  observe(target: HTMLCanvasElement) {
    if (target && "addEventListener" in target) {
      Object.entries(GlobalEvents).forEach(
        ([targetEvent, { dispatchEvents }]) => {
          const listener: Excel.Event.FnType = (...args) => {
            Object.entries(dispatchEvents).forEach(
              ([dispatchEvent, { triggerName }]) => {
                this[dispatchEvent as Excel.Event.Type].forEach((element) => {
                  element[triggerName as "triggerEvent"]?.call(
                    element,
                    dispatchEvent as Excel.Event.Type,
                    ...args
                  );
                });
              }
            );
          };
          target.addEventListener(targetEvent, listener);
        }
      );
    }
  }
  clear(gcList: Excel.Event.ObserverTypes[]) {
    for (let name in this) {
      if (Array.isArray(this[name])) {
        gcList.forEach((item) => {
          let i = (this[name] as Excel.Event.ObserverTypes[]).findIndex(
            (e) => e === item
          );
          if (!!~i) {
            (this[name] as Excel.Event.ObserverTypes[]).splice(i, 1);
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
  clearEventsWhenReRender() {
    Object.entries(GlobalEvents).forEach(
      ([targetEvent, { dispatchEvents }]) => {
        Object.entries(dispatchEvents).forEach(
          ([dispatchEvent, { triggerName }]) => {
            this[dispatchEvent as Excel.Event.Type].forEach((element) => {
              if (element.clearEventsWhenReRender) {
                let i = (
                  this[
                    dispatchEvent as Excel.Event.Type
                  ] as Excel.Event.ObserverTypes[]
                ).findIndex((e) => e === element);
                if (!!~i) {
                  (
                    this[
                      dispatchEvent as Excel.Event.Type
                    ] as Excel.Event.ObserverTypes[]
                  ).splice(i, 1);
                }
              }
            });
          }
        );
      }
    );
  }
}
