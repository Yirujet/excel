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
      | "keyup"
      | "resize"
      | "select";

    export type GlobalEvent = {
      [key in Exclude<
        Type,
        | "mouseenter"
        | "mousemove"
        | "click"
        | "clickoutside"
        | "resize"
        | "select"
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
    } & {
      resize?: {
        dispatchEvents: {
          resize: {
            triggerName: "triggerEvent";
          };
        };
      };
    } & {
      select?: {
        dispatchEvents: {
          select: {
            triggerName: "triggerEvent";
          };
        };
      };
    };

    export type EventInstance = {
      addEvent?: (eventName: Type, callback: FnType) => void;
      removeEvent?: (eventName: Type, callback: FnType) => void;
      triggerEvent?: (eventName: Type, ...args: any[]) => void;
      clearEvents?: (
        eventObserver: Excel.Event.ObserverInstance,
        obj: Excel.Event.ObserverTypes
      ) => void;
      destroy?: Excel.Event.FnType;
      clearEventsWhenReRender: boolean;
    };

    export type ObserverTypes = (
      | Excel.Sheet.SheetInstance
      | Excel.Scrollbar.ScrollbarInstance
      | Excel.Cell.CellInstance
      | Excel.FillHandle.FillHandleInstance
    ) &
      EventInstance;

    export type ObserverInstance = {
      [key in Type]: ObserverTypes[];
    } & {
      observe: (target: HTMLCanvasElement) => void;
      clear: (observers: ObserverTypes[]) => void;
      clearAll: () => void;
      clearEventsWhenReRender: Excel.Event.FnType;
    };
  }
}
