package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.xlsx.reader.model.SheetData;
import kz.greetgo.msoffice.xlsx.reader.model.StylesData;

import java.util.Objects;

public class SheetReader implements Sheet {
  private final StylesData styles;
  private final StoredStrings storedStrings;
  private final SheetRef sheetRef;
  private final SheetData sheetData;

  public SheetReader(StylesData styles, StoredStrings storedStrings, SheetRef sheetRef, SheetData sheetData) {
    Objects.requireNonNull(sheetData);
    this.styles = styles;
    this.storedStrings = storedStrings;
    this.sheetRef = sheetRef;
    this.sheetData = sheetData;
  }

  @Override
  public String name() {
    return sheetRef.name;
  }

  @Override
  public int index() {
    return sheetRef.index;
  }

  @Override
  public int rowCount() {
    return sheetData.rowSize();
  }

  @Override
  public Row row(int rowIndex) {
    return new RowReader(styles, storedStrings, rowIndex, sheetData.getRowData(rowIndex));
  }
}
