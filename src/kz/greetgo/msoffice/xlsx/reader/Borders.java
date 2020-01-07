package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.xlsx.reader.model.BorderSide;
import kz.greetgo.msoffice.xlsx.reader.model.BorderStyle;

public interface Borders {
  default BorderStyle topStyle() {
    return style(BorderSide.top);
  }

  default BorderStyle bottomStyle() {
    return style(BorderSide.bottom);
  }

  default BorderStyle leftStyle() {
    return style(BorderSide.left);
  }

  default BorderStyle rightStyle() {
    return style(BorderSide.right);
  }

  default BorderStyle diagonalStyle() {
    return style(BorderSide.diagonal);
  }

  BorderStyle style(BorderSide borderSide);
}
