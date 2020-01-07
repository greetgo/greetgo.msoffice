package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.xlsx.reader.model.Border4;
import kz.greetgo.msoffice.xlsx.reader.model.BorderSide;
import kz.greetgo.msoffice.xlsx.reader.model.BorderStyle;
import kz.greetgo.msoffice.xlsx.reader.model.CellXf;
import kz.greetgo.msoffice.xlsx.reader.model.ColData;
import kz.greetgo.msoffice.xlsx.reader.model.StylesData;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class XlsxReaderContext {

  public DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS");

  public final StoredStrings storedStrings;

  public final StylesData styles = new StylesData();

  public XlsxReaderContext(StoredStrings storedStrings) {
    this.storedStrings = storedStrings;
  }

  public String dateToStr(Date date) {
    return date == null ? null : dateFormat.format(date);
  }

  public BorderStyle getBorderStyle(BorderSide borderSide, int rowIndex, ColData col, SheetReader sheetReader) {
    {
      Border4 border4 = getBorder4(col.style);
      BorderStyle style = border4.get(borderSide).style;
      if (style != null && style != BorderStyle.NONE) {
        return style;
      }
    }

    int[] coordinates = new int[]{rowIndex, col.col};
    int[] newCoordinates = borderSide.move(coordinates);

    if (newCoordinates[0] == coordinates[0] && newCoordinates[1] == coordinates[1]) {
      return BorderStyle.NONE;
    }

    if (newCoordinates[0] < 0 || newCoordinates[1] < 0) {
      return BorderStyle.NONE;
    }

    CellReader cell = sheetReader.row(newCoordinates[0]).cell(newCoordinates[1]);

    {
      Border4 border4 = getBorder4(cell.col.style);
      BorderStyle style = border4.get(borderSide.mirror()).style;
      return style != null ? style : BorderStyle.NONE;
    }
  }

  private static final Border4 EMPTY = new Border4();

  private Border4 getBorder4(int style) {
    if (style == 0 && styles.cellXfList.isEmpty()) return EMPTY;
    CellXf cellXf = styles.cellXfList.get(style);
    int borderId = cellXf.borderId();
    if (borderId == 0 && styles.border4List.isEmpty()) return EMPTY;
    return styles.border4List.get(borderId);
  }
}
