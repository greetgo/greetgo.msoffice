package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.util.UtilOffice;
import kz.greetgo.msoffice.xlsx.reader.model.CellXf;
import kz.greetgo.msoffice.xlsx.reader.model.ColData;
import kz.greetgo.msoffice.xlsx.reader.model.NumFmtData;
import kz.greetgo.msoffice.xlsx.reader.model.StylesData;

import java.util.Date;

import static kz.greetgo.msoffice.util.UtilOffice.fnn;

public class CellReader implements Cell {
  private final StylesData styles;
  private final StoredStrings storedStrings;
  @SuppressWarnings({"FieldCanBeLocal", "unused"})
  private final int rowIndex;
  private final ColData col;

  public CellReader(StylesData styles, StoredStrings storedStrings, int rowIndex, ColData col) {
    this.styles = styles;
    this.storedStrings = storedStrings;
    this.rowIndex = rowIndex;
    this.col = col;
  }

  public static Cell empty(int rowIndex, int colIndex) {
    ColData colData = ColData.empty(colIndex);
    StylesData stylesData = new StylesData();
    return new CellReader(stylesData, null, rowIndex, colData);
  }

  @Override
  public String asText() {
    switch (col.valueType) {
      case UNKNOWN:
        return col.value;

      case STR:
        return storedStrings.get(Long.parseLong(col.value));

      default:
        throw new RuntimeException("Unknown ValueType = " + col.valueType);
    }
  }

  @Override
  public String format() {
    CellXf cellXf = styles.getCellXf(col.style);
    if (cellXf.numFmtId == null) {
      return "";
    }
    {
      NumFmtData numFmtData = styles.numFmtDataIdMap.get(cellXf.numFmtId);
      return numFmtData == null ? "" : fnn(numFmtData.formatCode, "");
    }
  }

  @Override
  public Date asDate() {
    return UtilOffice.excelToDate(col.value);
  }
}
