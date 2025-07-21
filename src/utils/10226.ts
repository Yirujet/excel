export default (index: number) => {
  let result = "";
  while (index >= 0) {
    const charCode = (index % 26) + 65;
    result = String.fromCharCode(charCode) + result;
    index = Math.floor(index / 26) - 1;
  }
  return result;
};
