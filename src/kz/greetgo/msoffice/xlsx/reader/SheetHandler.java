package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.util.UtilOffice;
import kz.greetgo.msoffice.util.XmlUtil;
import kz.greetgo.msoffice.xlsx.reader.model.ColData;
import kz.greetgo.msoffice.xlsx.reader.model.ColumnInfo;
import kz.greetgo.msoffice.xlsx.reader.model.MergeCell;
import kz.greetgo.msoffice.xlsx.reader.model.RowData;
import kz.greetgo.msoffice.xlsx.reader.model.SheetData;
import kz.greetgo.msoffice.xlsx.reader.model.ValueType;
import org.xml.sax.Attributes;

public class SheetHandler extends AbstractXmlHandler {

  private final SheetData sheet;
  private RowData row;
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
      row.index = Integer.parseInt(attributes.getValue("r")) - 1;
      row.height = UtilOffice.strToBd(attributes.getValue("ht"));
      return;
    }

    if ("/worksheet/sheetData/row/c".equals(tagPath)) {
      ColData col = new ColData();
      this.col = col;
      row.cols.add(col);
      col.col = UtilOffice.parseCellCoordinate(attributes.getValue("r"))[0];
      col.valueType = ValueType.parse(attributes.getValue("t"));
      return;
    }

    if ("/worksheet/mergeCells/mergeCell".equals(tagPath)) {
      MergeCell mergeCell = new MergeCell();
      mergeCell.parseAndSet(attributes.getValue("ref"));
      sheet.addMergeCell(mergeCell);
      return;
    }

    if ("/worksheet/sheetData/row/c/v".equals(tagPath)) return;

    System.out.println("j43bh25v6 ::     open " + tagPath + " " + XmlUtil.toStr(attributes) + " for " + sheet.name);
  }

  @Override
  protected void endTag(String tagPath) {
    if ("/worksheet/sheetData/row/c/v".equals(tagPath)) {
      col.value = text();
      return;
    }

    if ("/worksheet/sheetData/row".equals(tagPath)) {
      sheet.addRow(row);
      return;
    }
  }

  @Override
  public void endDocument() {
    System.out.println("3j2n5b5b :: sheet = " + sheet);
  }
}
