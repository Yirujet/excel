export default abstract class ExcelEvent implements Excel.Event.EventInstance {
  events: Record<string, Array<Excel.Event.FnType>> = {};
  clearEventsWhenReRender: boolean;
  constructor(clearEventsWhenReRender = false) {
    this.clearEventsWhenReRender = clearEventsWhenReRender;
  }
  addEvent(eventName: string, callback: Excel.Event.FnType) {
    if (!this.events[eventName]) {
      this.events[eventName] = [callback];
    } else {
      this.events[eventName].push(callback);
    }
  }
  removeEvent(eventName: string, callback: Excel.Event.FnType) {
    if (this.events[eventName]) {
      this.events[eventName] = this.events[eventName].filter(
        (e) => e !== callback
      );
    }
  }
  triggerEvent(eventName: string, ...args: any[]) {
    const eventList = this.events[eventName];
    if (eventList) {
      eventList.forEach((event) => {
        event.call(this, ...args);
      });
    }
  }
  clearEvents(
    eventObserver: Excel.Event.ObserverInstance,
    obj: Excel.Event.ObserverTypes
  ) {
    this.events = {};
    eventObserver.clear([obj]);
  }
  registerListenerFromOnProp(
    onObj: {
      [k in Excel.Event.Type]?: Excel.Event.FnType;
    },
    eventObserver: Excel.Event.ObserverInstance,
    obj: Excel.Event.ObserverTypes
  ) {
    if (onObj) {
      Object.entries(onObj).forEach(([eventname, callback]) => {
        this.addEvent(eventname, callback);
        if (eventname in eventObserver) {
          const i = eventObserver[eventname as Excel.Event.Type].findIndex(
            (e: any) => e === this
          );
          if (!!~i) {
            eventObserver[eventname as Excel.Event.Type].splice(i, 1, obj);
          } else {
            eventObserver[eventname as Excel.Event.Type].push(obj);
          }
        }
      });
    }
  }
}
