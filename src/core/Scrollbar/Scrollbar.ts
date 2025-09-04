import Element from "../../components/Element";
import {
  DEFAULT_SCROLLBAR_THUMB_BACKGROUND_COLOR,
  DEFAULT_SCROLLBAR_THUMB_DRAGGING_COLOR,
  DEFAULT_SCROLLBAR_THUMB_MIN_SIZE,
  DEFAULT_SCROLLBAR_TRACK_BACKGROUND_COLOR,
  DEFAULT_SCROLLBAR_TRACK_BORDER_COLOR,
  DEFAULT_SCROLLBAR_TRACK_SIZE,
} from "../../config/index";
import throttle from "../../utils/throttle";
import globalObj from "../globalObj";

export default class Scrollbar
  extends Element<null>
  implements Excel.Scrollbar.ScrollbarInstance
{
  track = {
    width: 0,
    height: 0,
    borderColor: DEFAULT_SCROLLBAR_TRACK_BORDER_COLOR,
    backgroundColor: DEFAULT_SCROLLBAR_TRACK_BACKGROUND_COLOR,
  };
  thumb = {
    width: 0,
    height: 0,
    padding: 0,
    min: DEFAULT_SCROLLBAR_THUMB_MIN_SIZE,
    backgroundColor: DEFAULT_SCROLLBAR_THUMB_BACKGROUND_COLOR,
    draggingColor: DEFAULT_SCROLLBAR_THUMB_DRAGGING_COLOR,
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
  eventObserver: Excel.Event.ObserverInstance;
  globalEventsObserver: Excel.Event.ObserverInstance;
  type: Excel.Scrollbar.Type = "vertical";
  constructor(
    layout: Excel.LayoutInfo,
    eventObserver: Excel.Event.ObserverInstance,
    globalEventsObserver: Excel.Event.ObserverInstance,
    type: Excel.Scrollbar.Type
  ) {
    super("");
    this.layout = layout;
    this.eventObserver = eventObserver;
    this.globalEventsObserver = globalEventsObserver;
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
    } else {
      this.mouseEntered = false;
    }
  }
  checkIn(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    if (
      !(
        offsetX < this.x ||
        offsetX > this.x + this.track.width ||
        offsetY < this.y ||
        offsetY > this.y + this.track.height
      )
    ) {
      globalObj.SET_CURSOR("default");
    }
  }
  scrollMove(offset: number, offsetProp: "x" | "y", maxScrollDistance: number) {
    if (this.dragging) {
      const curScrollbarVal = -this.value;
      const minMoveVal = offset - curScrollbarVal;
      const maxMoveVal = minMoveVal + maxScrollDistance;
      this.isLast = false;
      this.triggerEvent("percent", this.percent, this.type);
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
        this.triggerEvent("percent", this.percent, this.type);
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
  fillCoincide(ctx: CanvasRenderingContext2D) {
    ctx.save();
    ctx.strokeStyle = this.track.borderColor;
    ctx.beginPath();
    if (this.type === "horizontal") {
      ctx.moveTo(this.track.width, this.y);
      ctx.lineTo(this.track.width, this.y + this.track.height);
      ctx.lineTo(
        this.track.width + DEFAULT_SCROLLBAR_TRACK_SIZE,
        this.y + this.track.height
      );
    } else {
      ctx.moveTo(this.x, this.track.height);
      ctx.lineTo(this.x + this.track.width, this.track.height);
      ctx.lineTo(
        this.x + this.track.width,
        this.track.height + DEFAULT_SCROLLBAR_TRACK_SIZE
      );
    }
    ctx.closePath();
    ctx.stroke();
    ctx.fillStyle = this!.track.backgroundColor;
    ctx.fill();
    ctx.restore();
  }
}
