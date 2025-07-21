import Element from "../../components/Element";
import EventObserver from "../../utils/EventObserver";
import throttle from "../../utils/throttle";

export default class Scrollbar
  extends Element
  implements Excel.Scrollbar.ScrollbarInstance
{
  track = {
    width: 0,
    height: 0,
    borderColor: "#ebeef5",
    backgroundColor: "#f1f1f1",
  };
  thumb = {
    width: 0,
    height: 0,
    padding: 0,
    min: 20,
    backgroundColor: "#c1c1c1",
    draggingColor: "#787878",
  };
  value = 0;
  percent = 0;
  start = 0;
  end = 0;
  lastVal: number | null = null;
  position: Excel.Position | null = null;
  show = false;
  dragging = false;
  moveEvent: Excel.Event.FnType | null = null;
  offsetPercent = 0;
  isLast = false;
  layout: Excel.LayoutInfo | null = null;
  isHorizontalScrolling = false;
  callback: Excel.Event.FnType = () => {};
  eventObserver: EventObserver;
  globalEventsObserver: EventObserver;
  type: Excel.Scrollbar.Type = "vertical";
  constructor(
    layout: Excel.LayoutInfo,
    eventObserver: EventObserver,
    globalEventsObserver: EventObserver,
    callback: Excel.Event.FnType,
    type: Excel.Scrollbar.Type
  ) {
    super("");
    this.layout = layout;
    this.eventObserver = eventObserver;
    this.globalEventsObserver = globalEventsObserver;
    this.callback = callback;
    this.type = type;
  }
  checkHit(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    if (
      !(
        offsetX < this.position!.leftTop.x ||
        offsetX > this.position!.rightTop.x ||
        offsetY < this.position!.leftTop.y ||
        offsetY > this.position!.leftBottom.y
      )
    ) {
      this.mouseEntered = true;
      this.cursor = "default";
      this.dragging = true;
    }
  }
  scrollMove(
    offset: number,
    offsetProp: "x" | "y",
    maxScrollDistance: number,
    callback: Excel.Event.FnType
  ) {
    if (this.dragging) {
      const curScrollbarVal = -this.value;
      const minMoveVal = offset - curScrollbarVal;
      const maxMoveVal = minMoveVal + maxScrollDistance;
      this.isLast = false;
      callback(this.percent, this.type);
      this.moveEvent = throttle((e: MouseEvent) => {
        if (!this.dragging) return;
        this.isLast = false;
        let moveVal = e[offsetProp] - this.layout![offsetProp];
        if (moveVal > maxMoveVal) {
          moveVal = maxMoveVal;
        }
        if (moveVal < minMoveVal) {
          moveVal = minMoveVal;
        }
        let d = moveVal - offset;
        let direction;
        if (this.lastVal === null) {
          direction = d >= 0 ? true : false;
        } else {
          direction = moveVal - this.lastVal >= 0 ? true : false;
        }
        this.value = -d - curScrollbarVal;
        if (direction) {
          const deviation = Math.abs(this.value + maxScrollDistance);
          if (
            deviation < this.layout!.deviationCompareValue ||
            this.value <= -maxScrollDistance
          ) {
            this.value = -maxScrollDistance;
            this.isLast = true;
          }
        } else {
          if (
            this.value > 0 ||
            Math.abs(this.value) < this.layout!.deviationCompareValue
          ) {
            this.value = 0;
          }
        }
        this.percent = this.value / -maxScrollDistance;
        this.lastVal = offset;
        callback(this.percent, this.type);
      }, 50);
    }
  }
  render(ctx: CanvasRenderingContext2D) {
    ctx.save();
    ctx.strokeStyle = this.track.borderColor;
    ctx.strokeRect(this.x, this.y, this.track.width, this.track.height);
    ctx.fillStyle = this.track.backgroundColor;
    ctx.fillRect(this.x, this.y, this.track.width, this.track.height);
    ctx.fillStyle = this.dragging
      ? this.thumb.draggingColor
      : this.thumb.backgroundColor;
    if (this.type === "horizontal") {
      ctx.fillRect(
        this.x - this.value,
        this.y + this.thumb.padding,
        this.thumb.width,
        this.thumb.height
      );
    } else {
      ctx.fillRect(
        this.x + this.thumb.padding,
        this.y - this.value,
        this.thumb.width,
        this.thumb.height
      );
    }
    ctx.restore();
  }
}
