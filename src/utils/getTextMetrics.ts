export default (
  text: string,
  fontSize: number,
  ctx?: CanvasRenderingContext2D
): { width: number; height: number } => {
  if (!ctx) {
    text = text.toString();
    // 处理空文本情况
    if (!text || text.length === 0) {
      return { width: 0, height: fontSize };
    }

    let width = 0;
    const textLength = text.length;

    // 直接循环字符串，避免数组方法的额外开销
    for (let i = 0; i < textLength; i++) {
      // 根据字符编码判断是中文字符还是英文字符
      // 中文字符编码范围：\u4e00-\u9fa5
      const charCode = text.charCodeAt(i);
      // 中文字符宽度为fontSize，英文字符为fontSize/2
      width +=
        charCode >= 0x4e00 && charCode <= 0x9fa5 ? fontSize : fontSize / 2;
    }

    // 高度直接使用fontSize
    const height = fontSize;
    return { width, height };
  } else {
    const metrics = ctx.measureText(text);
    return { width: metrics.width, height: fontSize };
  }
};
