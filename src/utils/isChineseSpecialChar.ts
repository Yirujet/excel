export default (char: string): boolean => {
  if (char.length !== 1) {
    return false;
  }

  // 中文特殊字符正则表达式
  // 包括：中文标点符号、全角字符、中文特殊符号等
  // 排除基本汉字范围 \u4e00-\u9fa5（这些已在getTextMetrics中处理）
  const chineseSpecialCharRegex =
    /[\u3000-\u303f\uff00-\uffef\u2000-\u206f\u3400-\u4dbf\uf900-\ufaff]/;

  // 检查是否是中文特殊字符
  return chineseSpecialCharRegex.test(char);
};
