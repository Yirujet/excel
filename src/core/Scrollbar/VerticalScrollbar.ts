import throttle from "../../utils/throttle";
import Sheet from "../Sheet";
import Scrollbar from "./Scrollbar";

export default class VerticalScrollbar extends Scrollbar {
  constructor(
    layout: Excel.LayoutInfo,
    eventObserver: Excel.Event.ObserverInstance,
    globalEventsObserver: Excel.Event.ObserverInstance
  ) {
    super(layout, eventObserver, globalEventsObserver, "vertical");
    this.track.width = Sheet.DEFAULT_SCROLLBAR_TRACK_SIZE;
    this.thumb.width = Sheet.DEFAULT_SCROLLBAR_THUMB_SIZE;
    this.init();
  }
  init() {
    this.updateScrollbarInfo();
    this.updatePosition();
    this.initEvents();
  }
  updateScrollbarInfo() {
    if (this.layout!.bodyHeight < this.layout!.bodyRealHeight) {
      this.show = true;
    }
    if (this.show) {
      this.x = this.layout!.width;
      this.y = 0;
      this.track.height = this.layout!.height;
      this.thumb.height =
        this.track.height * (this.track.height / this.layout!.bodyRealHeight);
      if (this.thumb.height < this.thumb.min) {
        this.thumb.height = this.thumb.min;
      }
      this.thumb.padding = (this.track.width - this.thumb.width) / 2;
    }
  }
  initEvents() {
    const onStartScroll = (e: MouseEvent) => {
      if (!this.show) return;
      this.updatePosition();
      const { y } = e;
      this.checkHit(e);
      if (!this.mouseEntered) return;
      this.dragging = true;
      this.scrollMove(
        y - this.layout!.y,
        "y",
        this.track.height - this.thumb.height
      );
      const onEndScroll = () => {
        this.lastVal = null;
        this.dragging = false;
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
      if (this.isHorizontalScrolling) return;
      const { offsetX, offsetY } = e;
      if (
        offsetX >= this.layout!.x &&
        offsetX <= this.layout!.x + this.layout!.width + this.track.width &&
        offsetY >= this.layout!.y &&
        offsetY <= this.layout!.y + this.layout!.height
      ) {
        this.value -=
          e.deltaY * (this.track.height / this.layout!.bodyRealHeight);
        this.isLast = false;
        if (e.deltaY > 0) {
          const deviation = Math.abs(
            this.value + this.track.height - this.thumb.height
          );
          if (
            deviation < this.layout!.deviationCompareValue ||
            this.value + this.track.height <= this.thumb.height
          ) {
            this.value = this.thumb.height - this.track.height;
            this.isLast = true;
          }
        } else {
          if (this.value > 0) {
            this.value = 0;
          }
        }
        this.percent = this.value / (this.thumb.height - this.track.height);
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
    const onMouseMove = throttle(this.checkIn.bind(this), 50);
    const defaultEventListeners = {
      wheel: onWheel,
      mousedown: onStartScroll,
      mousemove: onMouseMove,
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
    const verticalThumbX = this.x! + this.thumb.padding;
    const verticalThumbY = this.y! - this.value;
    const verticalThumbWidth = this.thumb.width;
    const verticalThumbHeight = this.thumb.height;
    this.position = {
      leftTop: {
        x: verticalThumbX,
        y: verticalThumbY,
      },
      rightTop: {
        x: verticalThumbX + verticalThumbWidth,
        y: verticalThumbY,
      },
      rightBottom: {
        x: verticalThumbX + verticalThumbWidth,
        y: verticalThumbY + verticalThumbHeight,
      },
      leftBottom: {
        x: verticalThumbX,
        y: verticalThumbY + verticalThumbHeight,
      },
    };
  }
}
