class Cell implements Excel.Cell.CellInstance {
  width: number | null = null;
  height: number | null = null;
  rowIndex: number | null = null;
  colIndex: number | null = null;
  selected = false;
  cellName: string = "";
  x: number | null = null;
  y: number | null = null;
  position: {
    leftTop: {
      x: number;
      y: number;
    };
    rightTop: {
      x: number;
      y: number;
    };
    rightBottom: {
      x: number;
      y: number;
    };
    leftBottom: {
      x: number;
      y: number;
    };
  } = {
    leftTop: {
      x: 0,
      y: 0,
    },
    rightTop: {
      x: 0,
      y: 0,
    },
    rightBottom: {
      x: 0,
      y: 0,
    },
    leftBottom: {
      x: 0,
      y: 0,
    },
  };
  textStyle = {
    fontFamily: "sans-serif",
    fontSize: 12,
    bold: false,
    italic: false,
    underline: false,
    backgroundColor: "",
    color: "",
    align: "",
  };
  borderStyle = {
    solid: false,
    color: "",
    bold: false,
  };
  meta = null;
  value = "";
  fn = null;
}

export default Cell;
