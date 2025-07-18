export default abstract class ExcelEvent {
  events: Record<string, Array<(...params: any[]) => void>> = {};
  addEvent(eventName: string, callback: (...params: any[]) => void) {
    if (!this.events[eventName]) {
      this.events[eventName] = [callback];
    } else {
      this.events[eventName].push(callback);
    }
  }
  removeEvent(eventName: string, callback: (...params: any[]) => void) {
    if (this.events[eventName]) {
      this.events[eventName] = this.events[eventName].filter(
        (e) => e !== callback
      );
    }
  }
  triggerEvent(eventName: string, ...params: any[]) {
    const eventList = this.events[eventName];
    if (eventList) {
      eventList.forEach((event) => {
        event.call(this, ...params);
      });
    }
  }
  registerListenerFromOnProp(onObj: any, eventObserver: any) {
    if (onObj) {
      Object.entries(onObj).forEach(([eventname, callback]: any) => {
        this.addEvent(eventname, callback);
        if (eventname in eventObserver) {
          const i = eventObserver[eventname].findIndex((e: any) => e === this);
          if (!!~i) {
            eventObserver[eventname].splice(i, 1, this);
          } else {
            eventObserver[eventname].push(this);
          }
        }
      });
    }
  }
}
