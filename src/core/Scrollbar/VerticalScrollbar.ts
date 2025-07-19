import throttle from "../../utils/throttle";
import Scrollbar from "./Scrollbar";

export default class VerticalScrollbar extends Scrollbar {
  constructor(
    layout: any,
    eventObserver: any,
    callback: (...params: any[]) => void
  ) {
    super(layout, eventObserver, callback);
    this.track.width = 16;
    this.thumb.width = 16;
    this.init();
  }
  init() {
    if (this.layout.bodyHeight < this.layout.bodyRealHeight) {
      this.show = true;
    }
    if (this.show) {
      this.x = this.layout.width;
      this.y = this.layout.headerHeight;
      this.track.height = this.layout.height - this.layout.headerHeight;
      this.thumb.height =
        this.track.height * (this.track.height / this.layout.bodyRealHeight);
      if (this.thumb.height < this.thumb.min) {
        this.thumb.height = this.thumb.min;
      }
      this.thumb.padding = (this.track.width - this.thumb.width) / 2;
    }
    this.updatePosition();
    this.initEvents();
  }
  initEvents() {
    const onStartScroll = (e: any) => {
      this.updatePosition();
      const { y } = e;
      this.checkHit(e);
      this.scrollMove(
        y - this.layout.y,
        "y",
        this.track.height - this.thumb.height,
        this.callback,
        "vertical"
      );
      const onEndScroll = () => {
        this.lastVal = null;
        this.dragging = false;
        this.callback(this.percent, "vertical");
        window.removeEventListener("mousemove", this.moveEvent);
        window.removeEventListener("mouseup", onEndScroll);
      };
      if (this.dragging) {
        window.addEventListener("mousemove", this.moveEvent);
        window.addEventListener("mouseup", onEndScroll);
      }
    };
    this.isLast = false;
    const onWheel = throttle((e: any) => {
      e.stopPropagation();
      e.preventDefault();
      const { offsetX, offsetY } = e;
      if (
        this.show &&
        offsetX >= this.layout.x &&
        offsetX <= this.layout.x + this.layout.width + this.track.width &&
        offsetY >= this.layout.y &&
        offsetY <= this.layout.y + this.layout.height
      ) {
        this.value -=
          e.deltaY * (this.track.height / this.layout.bodyRealHeight);
        this.isLast = false;
        if (e.deltaY > 0) {
          const deviation = Math.abs(
            this.value + this.track.height - this.thumb.height
          );
          if (
            deviation < this.layout.deviationCompareValue ||
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
        this.callback(this.percent, "vertical");
      }
    }, 50);
    const defaultEventListeners = {
      wheel: onWheel,
      mousedown: onStartScroll,
    };
    this.registerListenerFromOnProp(defaultEventListeners, this.eventObserver);
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
