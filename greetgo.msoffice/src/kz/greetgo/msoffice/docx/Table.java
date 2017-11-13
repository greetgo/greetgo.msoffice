package kz.greetgo.msoffice.docx;

import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

public class Table implements FlowElement {
  private final String ownerPartName;
  private final MSHelper msHelper;
  
  Table(String ownerPartName, MSHelper msHelper) {
    this.ownerPartName = ownerPartName;
    this.msHelper = msHelper;
  }
  
  private List<TableCol> cols = new ArrayList<TableCol>();
  
  public TableCol createCol() {
    TableCol ret = new TableCol(cols.size());
    cols.add(ret);
    return ret;
  }
  
  private List<TableRow> rows = new ArrayList<TableRow>();
  private final Borders borders = new Borders("w:tblBorders");
  private final Margins margins = new Margins("w:tblCellMar");
  private String tableLookVal;
  private Integer tableWidth;
  private Integer tblCellSpacing = 0;
  
  class TableInfo {
    TableCol getTableCol(int colIndex) {
      if (colIndex < cols.size()) {
        return cols.get(colIndex);
      }
      for (int i = cols.size(); i < colIndex; i++) {
        cols.add(new TableCol(i));
      }
      TableCol ret = new TableCol(cols.size());
      if (colIndex == cols.size()) {
        cols.add(ret);
      }
      return ret;
    }
  }
  
  private final TableInfo tableInfo = new TableInfo();
  
  public TableRow createRow() {
    TableRow ret = new TableRow(tableInfo, ownerPartName, msHelper);
    rows.add(ret);
    return ret;
  }
  
  public TableRow getRow(int index) {
    return rows.get(index);
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<w:tbl>");
    if (cols.size() > 0) {
      out.print("<w:tblGrid>");
      for (TableCol c : cols) {
        c.write(out);
      }
      out.print("</w:tblGrid>");
    }
    
    out.print("<w:tblPr>");
    if (getTableWidth() != null) {
      out.print("<w:tblW w:w=\"" + getTableWidth() + "\" w:type=\"dxa\" />");
    }
    out.print("<w:tblCellSpacing w:w=\"" + 0 + "\" w:type=\"dxa\" />");
    getBorders().write(out);
    getMargins().write(out);
    if (tableLookVal != null) {
      out.print("<w:tblLook w:val=\"" + tableLookVal + "\" />");
    }
    out.print("</w:tblPr>");
    
    for (TableRow row : rows) {
      row.write(out);
    }
    out.print("</w:tbl>");
  }
  
  public Borders getBorders() {
    return borders;
  }
  
  /**
   * Создаёт параграф для примера. Этот параграф не используется для оттображения в документе, но
   * может служить для создания реальных параграфов как пример оформления
   * 
   * @param anotherExample
   *          другой пример параграфа, или null, для значений по умолчанию
   * @return параграф-пример
   */
  public Para createExamplePara(Para anotherExample) {
    Para ret = new Para("", null);
    if (anotherExample != null) {
      ret.copyDecorationFrom(anotherExample);
    } else if (msHelper != null) {
      Para dp = msHelper.getDefaultPara();
      if (dp != null) {
        ret.copyDecorationFrom(dp);
      }
    }
    return ret;
    
  }
  
  public String getTableLookVal() {
    return tableLookVal;
  }
  
  public void setTableLookVal(String tableLookVal) {
    this.tableLookVal = tableLookVal;
  }
  
  public Integer getTableWidth() {
    return tableWidth;
  }
  
  public void setTableWidth(Integer tableWidth) {
    this.tableWidth = tableWidth;
  }
  
  public Integer getTblCellSpacing() {
    return tblCellSpacing;
  }
  
  public void setTblCellSpacing(Integer tblCellSpacing) {
    this.tblCellSpacing = tblCellSpacing;
  }
  
  public Margins getMargins() {
    return margins;
  }
}
