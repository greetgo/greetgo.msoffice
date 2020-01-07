package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.xlsx.reader.model.ColData;
import kz.greetgo.msoffice.xlsx.reader.model.RowData;
import kz.greetgo.msoffice.xlsx.reader.model.StylesData;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.function.Function;

public class RowReader implements Row {

  private final StylesData styles;
  private final StoredStrings storedStrings;
  private final Function<Date, String> dateToStr;
  private final int rowIndex;
  private final RowData rowData;
  private List<CellReader> cells = null;

  public RowReader(StylesData styles, StoredStrings storedStrings,
                   Function<Date, String> dateToStr,
                   int rowIndex, RowData rowData) {
    this.styles = styles;
    this.storedStrings = storedStrings;
    this.dateToStr = dateToStr;
    this.rowIndex = rowIndex;
    this.rowData = rowData;
  }

  private List<CellReader> cells() {

    if (this.cells != null) return this.cells;

    List<CellReader> cells = new ArrayList<>();
    for (ColData col : rowData.cols) {
      cells.add(new CellReader(styles, storedStrings, dateToStr, rowIndex, col));
    }

    return this.cells = cells;
  }

  @Override
  public double height() {
    return rowData.heightAsDouble();
  }

  @Override
  public Cell cell(int i) {
    if (i < 0) throw new RuntimeException("Illegal column index = " + i);
    if (i >= rowData.cols.size()) return CellReader.empty(rowIndex, i);
    return cells().get(i);
  }

  @Override
  public int count() {
    return rowData.cols.size();
  }
}
