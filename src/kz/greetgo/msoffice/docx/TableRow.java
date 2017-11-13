package kz.greetgo.msoffice.docx;

import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

import kz.greetgo.msoffice.docx.Table.TableInfo;

public class TableRow implements XmlWriter {
  private final TableInfo tableInfo;
  private final String ownerPartName;
  private final MSHelper msHelper;
  
  TableRow(TableInfo tableInfo, String ownerPartName, MSHelper msHelper) {
    this.tableInfo = tableInfo;
    this.ownerPartName = ownerPartName;
    this.msHelper = msHelper;
  }
  
  private List<TableCell> cells = new ArrayList<TableCell>();
  private Integer tblCellSpacing = 0;
  private Integer height;
  
  public TableCell createCell() {
    TableCell ret = new TableCell(tableInfo.getTableCol(cells.size()), ownerPartName, msHelper);
    cells.add(ret);
    return ret;
  }
  
  public TableCell getCell(int index) {
    return cells.get(index);
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<w:tr>");
    out.print("<w:trPr>");
    out.print("<w:tblCellSpacing w:w=\"" + getTblCellSpacing() + "\" w:type=\"dxa\" />");
    if (getHeight() != null) {
      out.print("<w:trHeight w:val=\"" + getHeight() + "\" />");
    }
    out.print("</w:trPr>");
    for (TableCell c : cells) {
      c.write(out);
    }
    out.print("</w:tr>");
  }
  
  public Integer getTblCellSpacing() {
    return tblCellSpacing;
  }
  
  public void setTblCellSpacing(Integer tblCellSpacing) {
    this.tblCellSpacing = tblCellSpacing;
  }
  
  public Integer getHeight() {
    return height;
  }
  
  public void setHeight(Integer height) {
    this.height = height;
  }
  
}
