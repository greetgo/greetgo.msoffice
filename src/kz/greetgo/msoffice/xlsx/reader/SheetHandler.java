package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.util.UtilOffice;
import kz.greetgo.msoffice.xlsx.reader.model.ColData;
import kz.greetgo.msoffice.xlsx.reader.model.ColumnInfo;
import kz.greetgo.msoffice.xlsx.reader.model.MergeCell;
import kz.greetgo.msoffice.xlsx.reader.model.RowData;
import kz.greetgo.msoffice.xlsx.reader.model.SheetData;
import kz.greetgo.msoffice.xlsx.reader.model.ValueType;
import org.xml.sax.Attributes;

import static kz.greetgo.msoffice.util.UtilOffice.parseCellCoordinate;
import static kz.greetgo.msoffice.util.UtilOffice.strToIntOr;

public class SheetHandler extends AbstractXmlHandler {

  private final SheetData sheet;
  private RowData row;
  private int rowIndex;
  private ColData col;

  public SheetHandler(SheetData sheet) {
    this.sheet = sheet;
  }

  @Override
  protected void startTag(String tagPath, Attributes attributes) {

    if ("/worksheet/cols/col".equals(tagPath)) {
      ColumnInfo columnInfo = new ColumnInfo();
      sheet.columnInfoList.add(columnInfo);
      columnInfo.minIndex = Integer.parseInt(attributes.getValue("min"));
      columnInfo.maxIndex = Integer.parseInt(attributes.getValue("max"));
      columnInfo.width = UtilOffice.strToBd(attributes.getValue("width"));
      return;
    }

    if ("/worksheet/sheetData/row".equals(tagPath)) {
      RowData row = this.row = new RowData();
      rowIndex = Integer.parseInt(attributes.getValue("r")) - 1;
      row.height = UtilOffice.strToBd(attributes.getValue("ht"));
      return;
    }

    if ("/worksheet/sheetData/row/c".equals(tagPath)) {
      ColData col = new ColData();
      this.col = col;
      row.cols.add(col);
      col.col = parseCellCoordinate(attributes.getValue("r"))[0] - 1;
      col.valueType = ValueType.parse(attributes.getValue("t"));
      col.style = strToIntOr(attributes.getValue("s"), 0);
      return;
    }

    if ("/worksheet/mergeCells/mergeCell".equals(tagPath)) {
      MergeCell mergeCell = new MergeCell();
      mergeCell.parseAndSet(attributes.getValue("ref"));
      sheet.addMergeCell(mergeCell);
      return;
    }

    if ("/worksheet/sheetData/row/c/v".equals(tagPath)) return;

    if ("/worksheet/sheetViews/sheetView".equals(tagPath)) {
      if ("1".equals(attributes.getValue("tabSelected"))) {
        sheet.tabSelected = true;
      }
      return;
    }

    if ("/worksheet/sheetViews/sheetView/pane".equals(tagPath)) {

      if ("frozen".equals(attributes.getValue("state"))) {
        String ySplit = attributes.getValue("ySplit");
        if (ySplit != null && ySplit.trim().length() > 0) {
          sheet.frozenRowCount = Math.round(Float.parseFloat(ySplit));
        }
        return;
      }

      return;
    }

//    System.out.println("j43bh25v6 ::     open " + tagPath + " " + XmlUtil.toStr(attributes) + " for " + sheet.id);
  }

  @Override
  protected void endTag(String tagPath) {
    if ("/worksheet/sheetData/row/c/v".equals(tagPath)) {
      col.value = text();
      return;
    }

    if ("/worksheet/sheetData/row".equals(tagPath)) {
      sheet.addRow(rowIndex, row);
      return;
    }

    if ("/worksheet/sheetData/row/c/is/t".equals(tagPath)) {
      if (col.valueType == ValueType.INLINE_STR) {
        col.value = text();
      }
      return;
    }
  }
}
