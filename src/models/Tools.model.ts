namespace Excel {
  export namespace Tools {
    /**
     * 定义工具类型的枚举，用于表示 Excel 中可用的各种工具。
     * 这些工具类型可用于标识不同的操作，例如设置字体、调整单元格样式等。
     */
    export enum ToolType {
      /** 单元格字体家族选择工具，用于设置单元格内文本的字体 */
      CellFontFamily = "fontFamily",
      /** 单元格字体大小设置工具，用于调整单元格内文本的字号 */
      CellFontSize = "fontSize",
      /** 单元格文本加粗工具，用于将单元格内文本设置为加粗样式 */
      CellTextBold = "bold",
      /** 单元格文本倾斜工具，用于将单元格内文本设置为倾斜样式 */
      CellTextItalic = "italic",
      /** 单元格文本下划线工具，用于为单元格内文本添加下划线 */
      CellTextUnderline = "underline",
      /** 单元格边框设置工具，用于设置单元格的边框样式 */
      CellBorder = "border",
      /** 单元格文本颜色设置工具，用于改变单元格内文本的颜色 */
      CellTextColor = "color",
      /** 单元格背景颜色设置工具，用于设置单元格的背景颜色 */
      CellBackgroundColor = "backgroundColor",
      /** 单元格文本对齐方式设置工具，用于调整单元格内文本的对齐方式 */
      CellTextAlign = "align",
      /** 单元格合并工具，用于将多个相邻单元格合并为一个大单元格 */
      CellMerge = "merge",
      /** 单元格拆分工具，用于将合并后的单元格拆分为多个小单元格 */
      CellSplit = "split",
      /** 单元格函数工具，用于在单元格中插入和使用各种函数 */
      CellFunction = "function",
      /** 单元格插入工具，用于在工作表中插入新的单元格 */
      CellInsert = "insert",
      /** 单元格对角线设置工具，用于为单元格添加对角线 */
      CellDiagonal = "diagonal",
      /** 单元格冻结工具，用于冻结工作表中的行或列，方便查看数据 */
      CellFreeze = "freeze",
    }

    export interface ToolInstance {
      $el: HTMLElement | null;
      type: ToolType;
      disabled: boolean;
      render(): void;
    }
  }
}
