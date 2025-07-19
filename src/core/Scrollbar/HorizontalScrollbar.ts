import Scrollbar from "./Scrollbar";

export default class HorizontalScrollbar extends Scrollbar {
  constructor(
    layout: any,
    eventObserver: any,
    callback: (...params: any[]) => void
  ) {
    super(layout, eventObserver, callback);
    this.track.height = 16;
    this.thumb.height = 16;
    this.init();
  }
  init() {
    if (
      this.layout.bodyRealWidth >
      this.layout.width - this.layout.fixedLeftWidth
    ) {
      this.show = true;
    }
    if (this.show) {
      this.x = 0;
      this.y = this.layout.height;
      this.track.width = this.layout.width;
      this.thumb.width =
        this.track.width * (this.track.width / this.layout.bodyRealWidth);
      if (this.thumb.width < this.thumb.min) {
        this.thumb.width = this.thumb.min;
      }
      this.thumb.padding = (this.track.height - this.thumb.height) / 2;
    }
    this.updatePosition();
    this.initEvents();
  }
  initEvents() {
    const onStartScroll = (e: any) => {
      this.updatePosition();
      const { x } = e;
      this.checkHit(e);
      this.scrollMove(
        x - this.layout.x,
        "x",
        this.track.width - this.thumb.width,
        this.callback,
        "horizontal"
      );
      const onEndScroll = () => {
        this.lastVal = null;
        this.dragging = false;
        this.callback(this.percent, "horizontal");
        window.removeEventListener("mousemove", this.moveEvent);
        this.moveEvent = null;
        window.removeEventListener("mouseup", onEndScroll);
      };
      if (this.dragging) {
        window.addEventListener("mousemove", this.moveEvent);
        window.addEventListener("mouseup", onEndScroll);
      }
    };
    const defaultEventListeners = {
      mousedown: onStartScroll,
    };
    this.registerListenerFromOnProp(defaultEventListeners, this.eventObserver);
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
