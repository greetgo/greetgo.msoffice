package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.xlsx.reader.model.MergeCell;
import kz.greetgo.msoffice.xlsx.reader.model.SheetData;

import java.util.HashMap;
import java.util.Objects;

public class SheetReader implements Sheet {
  private final SheetRef sheetRef;
  final SheetData sheetData;
  private final XlsxReaderContext context;

  public SheetReader(XlsxReaderContext context, SheetRef sheetRef, SheetData sheetData) {
    Objects.requireNonNull(context);
    Objects.requireNonNull(sheetData);
    Objects.requireNonNull(sheetRef);
    this.context = context;
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
  public RowReader row(int rowIndex) {
    return new RowReader(context, this, rowIndex);
  }

  @Override
  public boolean tabSelected() {
    return sheetData.tabSelected;
  }

  @Override
  public int frozenRowCount() {
    return sheetData.frozenRowCount;
  }

  @Override
  public int cellMergeCount() {
    return sheetData.mergeCellSize();
  }

  private HashMap<Integer, MergeCell> mergeCellCache = new HashMap<>();

  @Override
  public MergeCell cellMerge(int i) {
    if (mergeCellCache.size() > 1000) mergeCellCache = new HashMap<>();
    return mergeCellCache.computeIfAbsent(i, sheetData::getMergeCell);
  }

}
