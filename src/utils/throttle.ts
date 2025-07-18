export default (func: any, interval: number) => {
  let lastTime = 0;
  return function (...args: any[]) {
    const now = Date.now();
    if (now - lastTime >= interval) {
      lastTime = now;
      func.apply(null, args);
    }
  };
};
