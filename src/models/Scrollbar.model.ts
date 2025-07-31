namespace Excel {
  export namespace Scrollbar {
    interface ScrollbarTrack {
      width: number;
      height: number;
      borderColor: string;
      backgroundColor: string;
    }

    interface ScrollbarThumb {
      width: number;
      height: number;
      padding: number;
      min: number;
      backgroundColor: string;
      draggingColor: string;
    }

    export type Type = "vertical" | "horizontal";

    export interface ScrollbarInstance {
      track: ScrollbarTrack;
      thumb: ScrollbarThumb;
      value: number;
      percent: number;
      start: number;
      end: number;
      lastVal: number | null;
      position: Excel.Position | null;
      show: boolean;
      dragging: boolean;
      moveEvent: Excel.Event.FnType | null;
      offsetPercent: number;
      isLast: boolean;
      layout: Excel.LayoutInfo | null;
      isHorizontalScrolling: boolean;
      callback: Excel.Event.FnType;
      eventObserver: Excel.Event.ObserverInstance;
      globalEventsObserver: Excel.Event.ObserverInstance;
      type: Type;
      mouseEntered: boolean;
    }
  }
}
