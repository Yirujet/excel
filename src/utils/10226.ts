export default (index: number) => {
  const val = [].slice
    .call(parseInt(index.toString()).toString(26))
    .map((e: string) => {
      if (/^\d$/.test(e)) {
        return String.fromCharCode(64 + ~~e);
      } else {
        return String.fromCharCode(e.codePointAt(0)! - 23);
      }
    })
    .join("");
  return val;
};
