import {
  DEFAULT_CELL_PADDING,
  DEFAULT_CELL_TEXT_FONT_SIZE,
  RESIZE_COL_SIZE,
  RESIZE_ROW_SIZE,
} from "../../config";
import getTextMetrics from "../../utils/getTextMetrics";
import FillHandle from "../FillHandle";
import HorizontalScrollbar from "../Scrollbar/HorizontalScrollbar";
import Scrollbar from "../Scrollbar/Scrollbar";
import VerticalScrollbar from "../Scrollbar/VerticalScrollbar";

export default abstract class SheetHandler {
  declare cells: Excel.Cell.CellInstance[][];
  declare verticalScrollBar: VerticalScrollbar | null;
  declare horizontalScrollBar: HorizontalScrollbar | null;
  declare layout: Excel.LayoutInfo | null;
  declare x: number;
  declare y: number;
  declare scroll: Excel.PositionPoint;
  declare fillHandle: FillHandle | null;
  declare mergedCells: Excel.Sheet.CellRange[];
  declare mode: Excel.Sheet.Mode;
  declare getCellPointByMousePosition: (
    mouseX: number,
    mouseY: number,
  ) => Excel.PositionPoint;
  declare findCellByPoint: (
    x: number,
    y: number,
    ignoreFixedX?: boolean,
    ignoreFixedY?: boolean,
  ) => Excel.Cell.CellInstance | null;
  declare getCell: (
    rowIndex: number,
    colIndex: number,
  ) => Excel.Cell.CellInstance;

  /**
   * 判断鼠标事件是否在滚动条范围内
   * @param e 鼠标事件
   * @returns 是否在滚动条范围内
   */
  private pointInScrollbar(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    if (this.verticalScrollBar || this.horizontalScrollBar) {
      const checkInScrollbar = (
        offsetX: number,
        offsetY: number,
        scrollbar: Scrollbar | null,
      ) => {
        if (scrollbar) {
          return !(
            offsetX < scrollbar.x ||
            offsetX > scrollbar.x + scrollbar.track.width ||
            offsetY < scrollbar.y ||
            offsetY > scrollbar.y + scrollbar.track.height
          );
        } else {
          return false;
        }
      };
      return (
        checkInScrollbar(offsetX, offsetY, this.verticalScrollBar) ||
        checkInScrollbar(offsetX, offsetY, this.horizontalScrollBar)
      );
    } else {
      return false;
    }
  }

  /**
   * 判断鼠标事件是否在画布范围内
   * @param e 鼠标事件
   * @returns 是否在画布范围内
   */
  private pointInCellRange(e: MouseEvent) {
    const { x, y, offsetX, offsetY } = e;
    return (
      x >= this.layout!.x &&
      y >= this.layout!.y &&
      x <= this.layout!.x + this.layout!.width &&
      y <= this.layout!.y + this.layout!.height &&
      offsetX >= this.x - this.layout!.x &&
      offsetX <= this.layout!.width &&
      offsetY >= this.y - this.layout!.y &&
      offsetY <= this.layout!.height
    );
  }

  /**
   * 判断鼠标事件是否在绝对固定单元格范围内(冻结行列交叉区域)
   * @param e 鼠标事件
   * @returns 是否在绝对固定单元格范围内
   */
  private pointInAbsFixedCell(e: MouseEvent) {
    if (this.pointInCellRange(e)) {
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const cell = this.findCellByPoint(
        x - this.scroll.x,
        y - this.scroll.y,
        false,
        false,
      );
      if (cell) {
        return cell.fixed.x && cell.fixed.y;
      }
    }
    return false;
  }

  /**
   * 判断鼠标事件是否在冻结单元格范围内(冻结行或列)
   * @param e 鼠标事件
   * @returns 是否在冻结单元格范围内
   */
  private pointInFixedCell(e: MouseEvent) {
    if (this.pointInCellRange(e)) {
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const cell = this.findCellByPoint(
        x - this.scroll.x,
        y - this.scroll.y,
        false,
        false,
      );
      if (cell) {
        return cell.fixed.x || cell.fixed.y;
      }
    }
    return false;
  }

  /**
   * 判断鼠标事件是否在冻结行单元格范围内
   * @param e 鼠标事件
   * @returns 是否在冻结行单元格范围内
   */
  private pointInFixedXCell(e: MouseEvent) {
    if (this.pointInFixedCell(e)) {
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const cell = this.findCellByPoint(
        x - this.scroll.x,
        y - this.scroll.y,
        false,
        false,
      );
      if (cell) {
        return cell.fixed.x;
      }
    }
    return false;
  }

  /**
   * 判断鼠标事件是否在冻结列单元格范围内
   * @param e 鼠标事件
   * @returns 是否在冻结列单元格范围内
   */
  private pointInFixedYCell(e: MouseEvent) {
    if (this.pointInFixedCell(e)) {
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const cell = this.findCellByPoint(
        x - this.scroll.x,
        y - this.scroll.y,
        false,
        false,
      );
      if (cell) {
        return cell.fixed.y;
      }
    }
    return false;
  }

  /**
   * 判断鼠标事件是否在普通单元格范围内
   * @param e 鼠标事件
   * @returns 是否在普通单元格范围内
   */
  private pointInNormalCell(e: MouseEvent) {
    if (this.pointInCellRange(e)) {
      return !this.pointInFixedCell(e);
    }
    return false;
  }

  /**
   * 判断鼠标事件是否在填充手柄范围内
   * @param e 鼠标事件
   * @returns 是否在填充手柄范围内
   */
  private pointInFillHandle(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    if (this.fillHandle) {
      return (
        offsetX >= this.fillHandle.position.leftTop.x - this.scroll.x &&
        offsetX <= this.fillHandle.position.rightBottom.x - this.scroll.x &&
        offsetY >= this.fillHandle.position.leftTop.y - this.scroll.y &&
        offsetY <= this.fillHandle.position.rightBottom.y - this.scroll.y
      );
    } else {
      return false;
    }
  }

  /**
   * 判断鼠标事件是否在行调整手柄范围内
   * @param e 鼠标事件
   * @returns 是否在行调整手柄范围内
   */
  private pointInRowResize(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    if (this.pointInFixedCell(e)) {
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const cell = this.findCellByPoint(x - this.scroll.x, y, false, false);
      if (cell) {
        if (cell.fixed.x) {
          return (
            offsetX >= cell.position!.leftTop.x &&
            offsetX <= cell.position!.rightTop.x &&
            offsetY >=
              cell.position!.leftBottom.y - this.scroll.y - RESIZE_ROW_SIZE &&
            offsetY <= cell.position!.leftBottom.y - this.scroll.y
          );
        } else {
          return false;
        }
      }
      return false;
    } else {
      return false;
    }
  }

  /**
   * 判断鼠标事件是否在列调整手柄范围内
   * @param e 鼠标事件
   * @returns 是否在列调整手柄范围内
   */
  private pointInColResize(e: MouseEvent) {
    const { offsetX, offsetY } = e;
    if (this.pointInFixedCell(e)) {
      const { x, y } = this.getCellPointByMousePosition(e.x, e.y);
      const cell = this.findCellByPoint(x, y - this.scroll.y, false, false);
      if (cell) {
        if (cell.fixed.y) {
          return (
            offsetX >=
              cell.position!.rightTop.x - this.scroll.x - RESIZE_COL_SIZE &&
            offsetX <= cell.position!.rightTop.x - this.scroll.x &&
            offsetY >= cell.position!.leftTop.y &&
            offsetY <= cell.position!.leftBottom.y
          );
        } else {
          return false;
        }
      }
      return false;
    } else {
      return false;
    }
  }

  /**
   * 检查单元格是否在合并单元格范围内
   * @param rowIndex 行索引
   * @param colIndex 列索引
   * @returns 是否在合并单元格范围内
   */
  private checkCellInMergedCells(rowIndex: number, colIndex: number): boolean {
    return this.mergedCells.some((item) => {
      const [startRowIndex, endRowIndex, startColIndex, endColIndex] = item;
      return (
        rowIndex! >= startRowIndex &&
        rowIndex! <= endRowIndex &&
        colIndex! >= startColIndex &&
        colIndex! <= endColIndex
      );
    });
  }

  /**
   * 转换合并单元格的位置
   * @returns 转换后的合并单元格范围
   */
  private transformMergedCells() {
    let rowAdjust: Record<number, number> = {};
    return this.mergedCells.map((range) => {
      const [minRowIndex, maxRowIndex, minColIndex, maxColIndex] = range;
      const leftTopCell = this.getCell(minRowIndex, minColIndex);
      const rightBottomCell = this.getCell(maxRowIndex, maxColIndex);
      const w =
        rightBottomCell.position.rightBottom.x! -
        leftTopCell.position.leftTop.x!;
      if (
        leftTopCell.meta?.type === "text" &&
        leftTopCell.value &&
        leftTopCell.width
      ) {
        rowAdjust = this.adjustCellPosition(
          leftTopCell,
          this.cells,
          minRowIndex,
          minColIndex,
          w,
          rowAdjust,
        );
      }
    });
  }

  /**
   * 转换单元格的位置
   * @param cells 单元格矩阵
   * @returns 转换后的单元格矩阵
   */
  private transformCells(
    cells: Excel.Cell.CellInstance[][],
  ): Excel.Cell.CellInstance[][] {
    const transformedCells = [...cells];
    let rowAdjust: Record<number, number> = {};

    for (let rowIndex = 0; rowIndex < transformedCells.length; rowIndex++) {
      const row = transformedCells[rowIndex];
      let currentX = 0;

      for (let colIndex = 0; colIndex < row.length; colIndex++) {
        const cell = row[colIndex];

        this.setCellPosition(
          cell,
          rowIndex,
          colIndex,
          transformedCells,
          currentX,
        );

        const cellInMergedCells = this.checkCellInMergedCells(
          this.mode === "view" ? rowIndex : rowIndex + 1,
          this.mode === "view" ? colIndex : colIndex + 1,
        );

        if (this.shouldProcessTextCell(cell, cellInMergedCells)) {
          rowAdjust = this.adjustCellPosition(
            cell,
            transformedCells,
            rowIndex,
            colIndex,
            cell.width!,
            rowAdjust,
          );
        }

        if (cell.meta?.type === "image") {
          this.processImageCell(cell, rowIndex, transformedCells, rowAdjust);
        }

        currentX += cell.width || 0;
      }
    }
    return transformedCells;
  }

  /**
   * 设置单元格的位置
   * @param cell 单元格实例
   * @param rowIndex 行索引
   * @param colIndex 列索引
   * @param transformedCells 转换后的单元格矩阵
   * @param currentX 当前X坐标
   */
  private setCellPosition(
    cell: Excel.Cell.CellInstance,
    rowIndex: number,
    colIndex: number,
    transformedCells: Excel.Cell.CellInstance[][],
    currentX: number,
  ): void {
    cell.x = currentX;

    if (rowIndex === 0) {
      cell.y = 0;
    } else {
      const previousRow = transformedCells[rowIndex - 1];
      if (previousRow && previousRow.length > 0) {
        const previousCell = previousRow[0];
        cell.y = previousCell.y! + (previousCell.height || 0);
      } else {
        cell.y = 0;
      }
    }
  }

  /**
   * 判断是否需要处理文本单元格
   * @param cell 单元格实例
   * @param cellInMergedCells 是否在合并单元格范围内
   * @returns 是否需要处理文本单元格
   */
  private shouldProcessTextCell(
    cell: Excel.Cell.CellInstance,
    cellInMergedCells: boolean,
  ): boolean {
    return (
      cell.meta?.type === "text" &&
      cell.value &&
      cell.width &&
      !cellInMergedCells
    );
  }

  /**
   * 处理图片单元格的位置调整
   * @param cell 单元格实例
   * @param rowIndex 行索引
   * @param transformedCells 转换后的单元格矩阵
   * @param rowAdjust 行调整记录
   * @returns
   */
  private processImageCell(
    cell: Excel.Cell.CellInstance,
    rowIndex: number,
    transformedCells: Excel.Cell.CellInstance[][],
    rowAdjust: Record<number, number>,
  ): void {
    if (
      !(cell.meta?.data as Excel.Cell.CellImageMetaData)?.height ||
      cell.height === (cell.meta?.data as Excel.Cell.CellImageMetaData).height
    ) {
      return;
    }

    const heightIncrease =
      (cell.meta?.data as Excel.Cell.CellImageMetaData).height -
      (cell.height || 0);

    if (heightIncrease <= 0) {
      return;
    }

    if (!rowAdjust[rowIndex] || rowAdjust[rowIndex] < heightIncrease) {
      const offset = this.calculateHeightOffset(
        rowAdjust,
        rowIndex,
        heightIncrease,
      );
      rowAdjust[rowIndex] = heightIncrease;

      this.adjustRowHeight(transformedCells, rowIndex, offset);
    }
  }

  /**
   * 计算高度调整偏移量
   * @param rowAdjust 行调整记录
   * @param rowIndex 行索引
   * @param heightIncrease 高度增加量
   * @returns 高度调整偏移量
   */
  private calculateHeightOffset(
    rowAdjust: Record<number, number>,
    rowIndex: number,
    heightIncrease: number,
  ): number {
    if (!rowAdjust[rowIndex]) {
      return heightIncrease;
    } else {
      return heightIncrease - rowAdjust[rowIndex];
    }
  }

  /**
   * 调整行高
   * @param cells 单元格矩阵
   * @param rowIndex 行索引
   * @param offset 高度调整偏移量
   */
  private adjustRowHeight(
    cells: Excel.Cell.CellInstance[][],
    rowIndex: number,
    offset: number,
  ): void {
    const currentRow = cells[rowIndex];
    currentRow.forEach((cell) => {
      cell.height! += offset;
      cell.updatePosition?.();
    });

    for (
      let adjustRowIndex = rowIndex + 1;
      adjustRowIndex < cells.length;
      adjustRowIndex++
    ) {
      const adjustRow = cells[adjustRowIndex];
      adjustRow.forEach((cell) => {
        cell.y = (cell.y || 0) + offset;
        cell.updatePosition?.();
      });
    }
  }

  /**
   * 截断文本内容以适应单元格宽度
   * @param content 要截断的内容
   * @param width 单元格宽度
   * @param fontSize 字体大小
   * @returns 截断后的内容数组
   */
  private truncateContent(
    content: string,
    width: number,
    fontSize: number,
  ): string[] {
    const value = String(content || "");

    if (!value) return [];

    const availableWidth = width - DEFAULT_CELL_PADDING * 2;

    if (availableWidth <= 0) return [];

    let currentWidth = 0;
    let result: string[] = [];
    let currentLine = "";

    for (let i = 0; i < value.length; i++) {
      const char = value[i];
      const { width: charWidth } = getTextMetrics(char, fontSize);

      if (currentWidth + charWidth > availableWidth) {
        if (currentLine) {
          result.push(currentLine);
        }

        if (charWidth > availableWidth) {
          result.push(char);
          currentLine = "";
          currentWidth = 0;
        } else {
          currentLine = char;
          currentWidth = charWidth;
        }
      } else {
        currentLine += char;
        currentWidth += charWidth;
      }
    }

    if (currentLine) {
      result.push(currentLine);
    }

    return result;
  }

  /**
   * 调整单元格位置
   * @param cell 单元格实例
   * @param cells 单元格矩阵
   * @param rowIndex 行索引
   * @param colIndex 列索引
   * @param width 单元格宽度
   * @param rowAdjust 行调整记录
   */
  private adjustCellPosition(
    cell: Excel.Cell.CellInstance,
    cells: Excel.Cell.CellInstance[][],
    rowIndex: number,
    colIndex: number,
    width: number,
    rowAdjust: Record<number, number>,
  ) {
    const { width: textWidth } = getTextMetrics(
      cell.value,
      cell.textStyle.fontSize,
    );
    if (textWidth > width) {
      if (cell.wrap === "no-wrap") {
        const widthIncrease = textWidth - width + DEFAULT_CELL_PADDING * 2;
        cell.width! += widthIncrease;
        cell.updatePosition?.();

        for (
          let adjustRowIndex = 0;
          adjustRowIndex < cells.length;
          adjustRowIndex++
        ) {
          const adjustRow = cells[adjustRowIndex];
          adjustRow[colIndex].width = cell.width;
          adjustRow[colIndex].updatePosition?.();
          for (
            let adjustColIndex = colIndex + 1;
            adjustColIndex < adjustRow.length;
            adjustColIndex++
          ) {
            const adjustCell = adjustRow[adjustColIndex];
            adjustCell.x = (adjustCell.x || 0) + widthIncrease;
            adjustCell.updatePosition?.();
          }
        }
      } else if (cell.wrap === "wrap") {
        const fontSize =
          cell.textStyle?.fontSize || DEFAULT_CELL_TEXT_FONT_SIZE;
        const valueSlices: string[] = cell.value
          .split("\n")
          .map((item: string) => this.truncateContent(item, width, fontSize));
        cell.valueSlices = valueSlices.flat();
        const heightIncrease = fontSize * cell.valueSlices.length;
        if (!rowAdjust[rowIndex] || rowAdjust[rowIndex] < heightIncrease) {
          let offset = 0;
          if (!rowAdjust[rowIndex]) {
            offset = heightIncrease - fontSize;
          } else {
            offset = heightIncrease - rowAdjust[rowIndex];
          }
          rowAdjust[rowIndex] = heightIncrease;
          cells[rowIndex].forEach((item) => {
            item.height! += offset;
            item.updatePosition?.();
          });
          for (
            let adjustRowIndex = rowIndex + 1;
            adjustRowIndex < cells.length;
            adjustRowIndex++
          ) {
            const adjustRow = cells[adjustRowIndex];
            for (
              let adjustColIndex = 0;
              adjustColIndex < adjustRow.length;
              adjustColIndex++
            ) {
              const adjustCell = adjustRow[adjustColIndex];
              adjustCell.y = (adjustCell.y || 0) + offset;
              adjustCell.updatePosition?.();
            }
          }
        }
      }
    }
    return rowAdjust;
  }
}
