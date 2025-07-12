class Tool implements Excel.Tools.ToolInstance {
  $el: HTMLElement | null = null;
  type!: Excel.Tools.ToolType;
  disabled = false;

  constructor(type: Excel.Tools.ToolType) {
    this.type = type;
    this.render();
  }

  render() {
    console.log(this.type);
  }
}

export default Tool;
