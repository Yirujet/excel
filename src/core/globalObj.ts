import { Parser } from "hot-formula-parser";

interface GlobalObj {
  SCROLL_X: number;
  SCROLL_Y: number;
  EVENT_LOCKED: boolean;
  FORMULA_PARSER: Parser | null;
  SET_CURSOR(cursor: string): void;
}

const globalObj: GlobalObj = {
  SCROLL_X: 0,
  SCROLL_Y: 0,
  EVENT_LOCKED: false,
  FORMULA_PARSER: null,
  SET_CURSOR(cursor: string) {
    document.body.style.cursor = cursor;
  },
};

export default globalObj;
