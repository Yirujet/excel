import {
  DEFAULT_SCROLLBAR_THUMB_SIZE,
  DEFAULT_SCROLLBAR_TRACK_SIZE,
} from "../../config/index";
import throttle from "../../utils/throttle";
import globalObj from "../globalObj";
import Scrollbar from "./Scrollbar";

export default class HorizontalScrollbar extends Scrollbar {
  constructor(
    layout: Excel.LayoutInfo,
    eventObserver: Excel.Event.ObserverInstance,
    globalEventsObserver: Excel.Event.ObserverInstance
  ) {
    super(layout, eventObserver, globalEventsObserver, "horizontal");
    this.track.height = DEFAULT_SCROLLBAR_TRACK_SIZE;
    this.thumb.height = DEFAULT_SCROLLBAR_THUMB_SIZE;
    this.init();
  }
  init() {
    this.updateScrollbarInfo();
    this.updatePosition();
    this.initEvents();
  }
  updateScrollbarInfo() {
    if (
      this.layout!.bodyRealWidth >
      this.layout!.width - this.layout!.fixedLeftWidth
    ) {
      this.show = true;
    }
    if (this.show) {
      this.x = 0;
      this.y = this.layout!.height;
      this.track.width = this.layout!.width;
      this.thumb.width =
        this.track.width * (this.track.width / this.layout!.bodyRealWidth);
      if (this.thumb.width < this.thumb.min) {
        this.thumb.width = this.thumb.min;
      }
      this.thumb.padding = (this.track.height - this.thumb.height) / 2;
    }
  }

  initEvents() {
    const onStartScroll = (e: MouseEvent) => {
      if (!this.show) return;
      this.updatePosition();
      const { x } = e;
      this.checkHit(e);
      if (!this.mouseEntered) return;
      this.dragging = true;
      globalObj.EVENT_LOCKED = true;
      this.scrollMove(
        x - this.layout!.x,
        "x",
        this.track.width - this.thumb.width
      );
      const onEndScroll = () => {
        this.lastVal = null;
        this.dragging = false;
        globalObj.EVENT_LOCKED = false;
        this.triggerEvent("percent", this.percent, this.type, true);
        window.removeEventListener("mousemove", this.moveEvent!);
        this.moveEvent = null;
        window.removeEventListener("mouseup", onEndScroll);
      };
      if (this.dragging) {
        window.addEventListener("mousemove", this.moveEvent!);
        window.addEventListener("mouseup", onEndScroll);
      }
    };
    this.isLast = false;
    const onWheel = throttle((e: WheelEvent) => {
      if (!this.show) return;
      e.stopPropagation();
      e.preventDefault();
      if (!this.isHorizontalScrolling) return;
      const { offsetX, offsetY } = e;
      if (
        offsetX >= this.layout!.x &&
        offsetX <= this.layout!.x + this.layout!.width + this.track.width &&
        offsetY >= this.layout!.y &&
        offsetY <= this.layout!.y + this.layout!.height
      ) {
        this.value -=
          e.deltaY * (this.track.width / this.layout!.bodyRealWidth);
        this.isLast = false;
        if (e.deltaY > 0) {
          const deviation = Math.abs(
            this.value + this.track.width - this.thumb.width
          );
          if (
            deviation < this.layout!.deviationCompareValue ||
            this.value + this.track.width <= this.thumb.width
          ) {
            this.value = this.thumb.width - this.track.width;
            this.isLast = true;
          }
        } else {
          if (this.value > 0) {
            this.value = 0;
          }
        }
        this.percent = this.value / (this.thumb.width - this.track.width);
        this.triggerEvent("percent", this.percent, this.type, true);
      }
    }, 50);
    const onKeydown = (e: KeyboardEvent) => {
      if (e.key === "Shift") {
        this.isHorizontalScrolling = true;
      }
    };
    const onKeyup = (e: KeyboardEvent) => {
      if (e.key === "Shift") {
        this.isHorizontalScrolling = false;
      }
    };
    const defaultEventListeners = {
      mousedown: onStartScroll,
      wheel: onWheel,
    };
    const globalEventListeners = {
      keydown: onKeydown,
      keyup: onKeyup,
    };
    this.registerListenerFromOnProp(
      defaultEventListeners,
      this.eventObserver,
      this
    );
    this.registerListenerFromOnProp(
      globalEventListeners,
      this.globalEventsObserver,
      this
    );
  }
  updatePosition() {
    const horizontalThumbX = this.x! - this.value;
    const horizontalThumbY = this.y! + this.thumb.padding;
    const horizontalThumbWidth = this.thumb.width;
    const horizontalThumbHeight = this.thumb.height;
    this.position = {
      leftTop: {
        x: horizontalThumbX,
        y: horizontalThumbY,
      },
      rightTop: {
        x: horizontalThumbX + horizontalThumbWidth,
        y: horizontalThumbY,
      },
      rightBottom: {
        x: horizontalThumbX + horizontalThumbWidth,
        y: horizontalThumbY + horizontalThumbHeight,
      },
      leftBottom: {
        x: horizontalThumbX,
        y: horizontalThumbY + horizontalThumbHeight,
      },
    };
  }
}
