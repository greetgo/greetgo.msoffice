package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.util.UtilOffice;
import kz.greetgo.msoffice.xlsx.reader.model.CellXf;
import kz.greetgo.msoffice.xlsx.reader.model.ColData;
import kz.greetgo.msoffice.xlsx.reader.model.HorizontalAlign;
import kz.greetgo.msoffice.xlsx.reader.model.NumFmtData;
import kz.greetgo.msoffice.xlsx.reader.model.VerticalAlign;

import java.util.Date;

import static kz.greetgo.msoffice.util.UtilOffice.fnn;

public class CellReader implements Cell {
  private final XlsxReaderContext context;
  @SuppressWarnings({"FieldCanBeLocal", "unused"})
  private final int rowIndex;
  final ColData col;
  private final SheetReader sheetReader;

  public CellReader(XlsxReaderContext context, int rowIndex, ColData col, SheetReader sheetReader) {
    this.context = context;
    this.rowIndex = rowIndex;
    this.col = col;
    this.sheetReader = sheetReader;
  }

  public static CellReader empty(XlsxReaderContext context, int rowIndex, int colIndex, SheetReader sheetReader) {
    return new CellReader(context, rowIndex, ColData.empty(colIndex), sheetReader);
  }

  @Override
  public String asText() {
    switch (col.valueType) {
      case UNKNOWN:
        if (isDate()) return context.dateToStr(asDate());
        return col.value;

      case STR:
        return context.storedStrings.get(Long.parseLong(col.value));

      default:
        throw new RuntimeException("Unknown ValueType = " + col.valueType);
    }
  }

  @Override
  public String format() {
    CellXf cellXf = context.styles.getCellXf(col.style);
    if (cellXf.numFmtId == null) {
      return "";
    }
    {
      NumFmtData numFmtData = context.styles.numFmtDataIdMap.get(cellXf.numFmtId);
      return numFmtData == null ? "" : fnn(numFmtData.formatCode, "");
    }
  }

  @Override
  public boolean isDate() {
    return UtilOffice.isFormatForDate(format());
  }

  @Override
  public Date asDate() {
    return UtilOffice.excelToDate(col.value);
  }

  @Override
  public Borders borders() {
    return borderSide -> context.getBorderStyle(borderSide, rowIndex, col, sheetReader);
  }

  @Override
  public HorizontalAlign horAlign() {
    return context.styles.getCellXf(col.style).horizontalAlign;
  }

  @Override
  public VerticalAlign vertAlign() {
    return context.styles.getCellXf(col.style).verticalAlign;
  }
}
