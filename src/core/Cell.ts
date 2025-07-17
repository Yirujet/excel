class Cell implements Excel.Cell.CellInstance {
  width = null;
  height = null;
  rowIndex = null;
  colIndex = null;
  selected = false;
  x = null;
  y = null;
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
