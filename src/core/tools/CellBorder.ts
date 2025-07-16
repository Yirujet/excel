import Element from "../../components/Element";
import Excel from "../../core/Excel";

declare class Drop {
  constructor(
    el: HTMLElement,
    panel: HTMLElement,
    options?: {
      eventType: "hover" | "click";
      position?: string;
      offsets?: {
        x: number;
        y: number;
      };
      onShow?: () => void;
    }
  );
}

declare class Color {
  constructor(
    el: HTMLElement,
    options?: {
      onShow?: () => void;
    }
  );
}

export default class CellBorder extends Element {
  constructor() {
    super("div");
    this.render();
  }

  render() {
    this.addClass(`${Excel.CSS_PREFIX}-cell-border`);
    this.addClass(`${Excel.CSS_PREFIX}-icon-button`);
    const icon = new Element("i");
    icon.addClass("iconfont");
    icon.addClass("icon-border-none");
    this.add(icon.$el!);
    const arrowDownIcon = new Element("i");
    arrowDownIcon.addClass("iconfont");
    arrowDownIcon.addClass("icon-arrow-down");
    this.add(arrowDownIcon.$el!);

    const panel = new Element("div");
    panel.addClass(`${Excel.CSS_PREFIX}-cell-border-drop-panel`);

    const group1 = new Element("span");
    group1.$el!.innerHTML = "边框";
    group1.addClass(`${Excel.CSS_PREFIX}-cell-border-drop-group-title`);
    panel.add(group1.$el!);

    const noBorderItem = new Element("div");
    noBorderItem.addClass(`${Excel.CSS_PREFIX}-cell-border-drop-item`);
    noBorderItem.$el!.innerHTML =
      '<i class="iconfont icon-border-none"></i>无边框';
    panel.add(noBorderItem.$el!);

    const allBorderItem = new Element("div");
    allBorderItem.addClass(`${Excel.CSS_PREFIX}-cell-border-drop-item`);
    allBorderItem.$el!.innerHTML =
      '<i class="iconfont icon-border-all"></i>所有边框';
    panel.add(allBorderItem.$el!);

    const group2 = new Element("span");
    group2.$el!.innerHTML = "绘图边框";
    group2.addClass(`${Excel.CSS_PREFIX}-cell-border-drop-group-title`);
    panel.add(group2.$el!);

    const borderColorItem = new Element("div");
    borderColorItem.addClass(`${Excel.CSS_PREFIX}-cell-border-drop-item`);
    borderColorItem.$el!.innerHTML =
      '<i class="iconfont icon-border-all"></i>线条颜色<i class="iconfont icon-arrow-right"></i>';
    panel.add(borderColorItem.$el!);
    const borderColorPanel = new Element("div");
    borderColorPanel.addClass(
      `${Excel.CSS_PREFIX}-cell-border-color-drop-panel`
    );

    const colorPicker = new Element("input");
    colorPicker.addClass("ui-color-input");
    colorPicker.$el!.setAttribute("type", "color");
    // const colorPicker = new Element("div");
    // // colorPicker.addClass("ui-color-input");
    // const color = new Color(colorPicker.$el!, {
    //   onShow() {
    //     const colorPanel = document.querySelector(".ui-color-container");
    //     borderColorPanel.add(colorPanel as HTMLElement);
    //     borderColorPanel.$el!.style.width = colorPanel!.clientWidth + "px";
    //     borderColorPanel.$el!.style.height = colorPanel!.clientHeight + "px";
    //   },
    // });
    borderColorPanel.add(colorPicker.$el!);
    const borderColorDrop = new Drop(
      borderColorItem.$el!,
      borderColorPanel.$el!,
      {
        eventType: "click",
        offsets: {
          x: 90,
          y: 0,
        },
        onShow() {
          colorPicker.$el?.click();
          setTimeout(() => {
            const colorPanel = document.querySelector(".ui-color-container");
            console.log(colorPanel);
            if (colorPanel) {
              borderColorPanel.add(colorPanel as HTMLElement);
              borderColorPanel.$el!.style.width =
                colorPanel!.clientWidth + "px";
              borderColorPanel.$el!.style.height =
                colorPanel!.clientHeight + "px";
            }
          }, 300);
        },
      }
    );

    const borderStyleItem = new Element("div");
    borderStyleItem.addClass(`${Excel.CSS_PREFIX}-cell-border-drop-item`);
    borderStyleItem.$el!.innerHTML =
      '<i class="iconfont icon-border-style"></i>线条样式<i class="iconfont icon-arrow-right"></i>';
    panel.add(borderStyleItem.$el!);
    const drop = new Drop(this.$el!, panel.$el!, {
      eventType: "click",
    });
  }
}
