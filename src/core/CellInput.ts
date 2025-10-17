import Element from "../components/Element";
import {
  DEFAULT_CELL_INPUT_COLOR,
  DEFAULT_CELL_INPUT_FONT_FAMILY,
  DEFAULT_CELL_INPUT_FONT_SIZE,
  DEFAULT_CELL_INPUT_MARGIN,
  DEFAULT_CELL_INPUT_PLACEHOLDER_COLOR,
} from "../config/index";
import getTextMetrics from "../utils/getTextMetrics";

class CellInput extends Element<HTMLDivElement> {
  layout: Excel.LayoutInfo;
  cell: Excel.Cell.CellInstance | null = null;
  constructor(layout: Excel.LayoutInfo) {
    super("div");
    this.layout = layout;
    this.init();
  }
  init() {
    if (!document.head.querySelector("#excel-cell-input-div")) {
      const styleEl = document.createElement("style");
      styleEl.id = "excel-cell-input-div";
      styleEl.innerHTML = `
                        .excel-cell-input-div {
                            position: fixed;
                            margin: 0;
                            background-color: #fff;
                            color: ${DEFAULT_CELL_INPUT_COLOR};
                            font-size: ${DEFAULT_CELL_INPUT_FONT_SIZE}px;
                            font-family: ${DEFAULT_CELL_INPUT_FONT_FAMILY};
                            border: none;
                            outline: none;
                            overflow: hidden;
                            white-space: nowrap;
                            z-index: 999;
                        }
                        .excel-cell-input-div:empty:before {
                            content: attr(placeholder);
                            color: ${DEFAULT_CELL_INPUT_PLACEHOLDER_COLOR};
                        }
                        .excel-cell-input-div:focus:before {
                            content: none;
                        }
                    `;
      document.head.appendChild(styleEl);
    }
    if (!document.body.contains(this.$el!)) {
      this.$el!.setAttribute("contenteditable", "true");
      this.$el!.setAttribute("placeholder", "");
      this.$el!.className = "excel-cell-input-div";
      this.$el!.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          this.triggerEvent(
            "input",
            (e.target as HTMLDivElement).innerText,
            this.cell
          );
        }
      });
      this.$el!.addEventListener("blur", (e) => {
        this.triggerEvent(
          "input",
          (e.target as HTMLDivElement).innerText,
          this.cell
        );
      });
      document.body.append(this.$el!);
    }
  }
  hide() {
    this.$el!.style.display = "none";
    this.$el!.innerText = "";
    this.cell = null;
  }
  setValue(value: string) {
    if (!this.$el) return;
    this.$el!.innerText = value;
  }
  render(cell: Excel.Cell.CellInstance, scrollX: number, scrollY: number) {
    const { height: wordHeight } = getTextMetrics("1", cell.textStyle.fontSize);
    this.$el!.innerText = cell.value;
    this.$el!.style.width = `${cell.width! - 2 * DEFAULT_CELL_INPUT_MARGIN}px`;
    this.$el!.style.height = `${
      cell.height! - 2 * DEFAULT_CELL_INPUT_MARGIN
    }px`;
    this.$el!.style.lineHeight = `${wordHeight}px`;
    this.$el!.style.left = `${
      cell.x! - scrollX + this.layout.x + DEFAULT_CELL_INPUT_MARGIN
    }px`;
    this.$el!.style.top = `${
      cell.y! - scrollY + this.layout.y + DEFAULT_CELL_INPUT_MARGIN
    }px`;
    this.$el!.style.display = "block";
    this.$el!.focus();
    this.cell = cell;
  }
}

export default CellInput;
