namespace Excel {
  export namespace Event {
    export type FnType = (...args: any[]) => void;

    export type Type =
      | "mouseenter"
      | "mouseleave"
      | "mousemove"
      | "mousedown"
      | "mouseup"
      | "click"
      | "clickoutside"
      | "wheel"
      | "keydown"
      | "keyup";

    export type GlobalEvent = {
      [key in Exclude<
        Type,
        "mouseenter" | "mousemove" | "click" | "clickoutside"
      >]: {
        dispatchEvents: {
          [p in key]: {
            triggerName: "triggerEvent";
          };
        };
      };
    } & {
      mousemove: {
        dispatchEvents: {
          [p in "mousemove" | "mouseenter" | "mouseleave"]: {
            triggerName: "triggerEvent";
          };
        };
      };
    } & {
      click: {
        dispatchEvents: {
          click: {
            triggerName: "triggerClickEvent";
          };
          clickoutside: {
            triggerName: "triggerEvent";
          };
        };
      };
    };

    export type EventInstance = {
      addEvent(eventName: Type, callback: FnType): void;
      removeEvent(eventName: Type, callback: FnType): void;
      triggerEvent(eventName: Type, ...args: any[]): void;
      destroy?: Excel.Event.FnType;
    };

    export type ObserverTypes = Excel.Scrollbar.ScrollbarInstance &
      EventInstance;

    export type ObserverInstance = {
      [key in Type]: ObserverTypes[];
    };
  }
}
