export default (text: string, fontSize: number) => {
  const width = [].slice
    .call(text || "")
    .map((e) => (String(e).charCodeAt(0) > 255 ? fontSize : fontSize / 2))
    .reduce((p, c) => p + c, 0);
  const height = fontSize;
  // 为了更好的性能，计算字符串宽度不使用canvas的measureText
  return { width, height };
};
