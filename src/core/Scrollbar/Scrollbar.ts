import Element from "../../components/Element";
import throttle from "../../utils/throttle";

export default class Scrollbar extends Element {
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
  position: any = null;
  show = false;
  dragging = false;
  moveEvent: any = null;
  offsetPercent = 0;
  isLast = false;
  layout: any = {
    x: 0,
    y: 0,
    width: 0,
    height: 0,
    headerHeight: 0,
    fixedLeftWidth: 0,
    bodyHeight: 0,
    bodyRealWidth: 0,
    bodyRealHeight: 0,
    target: null,
    restHeight: 0,
    restWidth: 0,
  };
  callback: (...params: any[]) => void = () => {};
  eventObserver: any;
  constructor(
    layout: any,
    eventObserver: any,
    callback: (...params: any[]) => void
  ) {
    super("");
    this.layout = layout;
    this.eventObserver = eventObserver;
    this.callback = callback;
  }
  checkHit(e: any) {
    const { offsetX, offsetY } = e;
    if (
      !(
        offsetX < this.position.leftTop.x ||
        offsetX > this.position.rightTop.x ||
        offsetY < this.position.leftTop.y ||
        offsetY > this.position.leftBottom.y
      )
    ) {
      this.mouseEntered = true;
      this.cursor = "default";
      this.dragging = true;
    }
  }
  scrollMove(
    offset: number,
    offsetProp: string,
    maxScrollDistance: number,
    callback: (...params: any[]) => void,
    type: "vertical" | "horizontal"
  ) {
    if (this.dragging) {
      const curScrollbarVal = -this.value;
      const minMoveVal = offset - curScrollbarVal;
      const maxMoveVal = minMoveVal + maxScrollDistance;
      this.isLast = false;
      callback(this.percent, type);
      this.moveEvent = throttle((e: any) => {
        if (!this.dragging) return;
        this.isLast = false;
        let moveVal = e[offsetProp] - this.layout[offsetProp];
        if (moveVal > maxMoveVal) {
          moveVal = maxMoveVal;
        }
        if (moveVal < minMoveVal) {
          moveVal = minMoveVal;
        }
        // console.log("***", e[offsetProp], e, moveVal, maxMoveVal);
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
            deviation < this.layout.deviationCompareValue ||
            this.value <= -maxScrollDistance
          ) {
            this.value = -maxScrollDistance;
            this.isLast = true;
          }
        } else {
          if (
            this.value > 0 ||
            Math.abs(this.value) < this.layout.deviationCompareValue
          ) {
            this.value = 0;
          }
        }
        // console.log("***", this.isLast, this.value);
        this.percent = this.value / -maxScrollDistance;
        this.lastVal = offset;
        callback(this.percent, type);
      }, 50);
    }
  }
}
