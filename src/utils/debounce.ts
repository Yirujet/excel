export default (func: Excel.Event.FnType, interval: number) => {
  let timeout: number;
  return function (...args: any[]) {
    clearTimeout(timeout);
    timeout = window.setTimeout(() => {
      func.apply(null, args);
    }, interval);
  };
};
