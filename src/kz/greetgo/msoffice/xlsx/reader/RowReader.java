package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.xlsx.reader.model.ColData;
import kz.greetgo.msoffice.xlsx.reader.model.RowData;

import java.util.ArrayList;
import java.util.List;

public class RowReader implements Row {

  private final XlsxReaderContext context;
  private final SheetReader sheetReader;
  private final int rowIndex;
  private final RowData rowData;
  private List<CellReader> cells = null;

  public RowReader(XlsxReaderContext context, SheetReader sheetReader, int rowIndex) {

    this.context = context;
    this.sheetReader = sheetReader;
    this.rowIndex = rowIndex;
    rowData = sheetReader.sheetData.getRowData(rowIndex);
  }

  private List<CellReader> cells() {

    if (this.cells != null) return this.cells;

    if (rowData.cols.isEmpty()) {
      return this.cells = new ArrayList<>();
    }

    int colCount = rowData.cols.get(rowData.cols.size() - 1).col + 1;

    List<CellReader> cells = new ArrayList<>();

    int colIndex = 0;
    for (int col = 0; col < colCount; col++) {
      ColData colData = rowData.cols.get(colIndex);
      if (col < colData.col) {
        cells.add(CellReader.empty(context, rowIndex, col, sheetReader));
      } else {
        cells.add(new CellReader(context, rowIndex, colData, sheetReader));
        colIndex++;
      }
    }

    return this.cells = cells;
  }

  @Override
  public double height() {
    return rowData.heightAsDouble();
  }

  @Override
  public CellReader cell(int i) {
    if (i < 0) throw new RuntimeException("Illegal column index = " + i);
    return i < cells().size() ? cells().get(i) : CellReader.empty(context, rowIndex, i, sheetReader);
  }

  @Override
  public int count() {
    return rowData.cols.size();
  }
}
