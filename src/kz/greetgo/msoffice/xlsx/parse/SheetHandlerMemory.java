package kz.greetgo.msoffice.xlsx.parse;

import java.math.BigDecimal;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import kz.greetgo.msoffice.UtilOffice;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

public class SheetHandlerMemory extends DefaultHandler implements Sheet {
  private final int sheetNo;
  private final String name;
  private boolean active;
  private Map<Long, Map<Integer, MyCell>> rows = new TreeMap<Long, Map<Integer, MyCell>>();
  private Map<Integer, Map<Long, Integer>> rowLength = new TreeMap<Integer, Map<Long, Integer>>();
  private final Map<Long, String> values;
  
  public SheetHandlerMemory(String sheetPathName, int sheetNo, Map<Long, String> values) {
    this.sheetNo = sheetNo;
    this.values = values;
    {
      int startLen = "xl/worksheets/".length();
      int endLen = ".xml".length();
      name = sheetPathName.substring(startLen, sheetPathName.length() - endLen);
    }
  }
  
  private final XmlIn in = new XmlIn();
  
  @Override
  public void startDocument() {}
  
  @Override
  public void endDocument() {}
  
  @Override
  public void startElement(String uri, String localName, String qName, Attributes attributes)
      throws SAXException {
    textBuilder = null;
    in.stepIn(qName);
    
    if ("worksheet/sheetData/row/c".equals(in.current())) {
      cell.r = attributes.getValue("r");
      cell.t = attributes.getValue("t");
      cell.s = attributes.getValue("s");
    }
    if ("worksheet/sheetViews/sheetView".equals(in.current())) {
      if ("1".equals(attributes.getValue("tabSelected"))) {
        this.active = true;
      }
    }
  }
  
  private static class CellInfo implements Cloneable {
    String r, t, s, v;
    
    long row;
    int col;
    
    public void calcRowCol() {
      int index = 0;
      while (true) {
        if (index >= r.length()) break;
        char c = r.charAt(index);
        if ('0' <= c && c <= '9') break;
        index++;
      }
      {
        row = Long.parseLong(r.substring(index));
        col = UtilOffice.parseLettersNumber(r.substring(0, index));
      }
    }
    
    @Override
    public int hashCode() {
      final int prime = 31;
      int result = 1;
      result = prime * result + col;
      result = prime * result + (int)(row ^ (row >>> 32));
      return result;
    }
    
    @Override
    public boolean equals(Object obj) {
      if (this == obj) return true;
      if (obj == null) return false;
      if (getClass() != obj.getClass()) return false;
      CellInfo other = (CellInfo)obj;
      if (col != other.col) return false;
      if (row != other.row) return false;
      return true;
    }
    
    protected Object clone() throws CloneNotSupportedException {
      return super.clone();
    }
    
  }
  
  final CellInfo cell = new CellInfo();
  
  @Override
  public void endElement(String uri, String localName, String qName) throws SAXException {
    if ("worksheet/sheetData/row/c/v".equals(in.current())) {
      cell.v = text();
      appendCell(cell);
    }
    in.stepOut();
  }
  
  private StringBuilder textBuilder = null;
  
  private String text() {
    if (textBuilder == null) return "";
    return textBuilder.toString();
  }
  
  @Override
  public void characters(char[] ch, int start, int length) throws SAXException {
    if (textBuilder == null) textBuilder = new StringBuilder();
    textBuilder.append(ch, start, length);
  }
  
  private void appendCell(CellInfo cell) {
    try {
      putIntoMap(cell);
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }
  
  private void putIntoMap(CellInfo cell) throws CloneNotSupportedException {
    cell.calcRowCol();
    if (!rows.containsKey((cell.row - 1))) {
      rows.put(cell.row - 1, new TreeMap<Integer, MyCell>());
    }
    Map<Integer, MyCell> row = rows.get(cell.row - 1);
    if (!row.containsKey(cell.col)) {
      row.put(cell.col, new MyCell(sheetNo, (int)(cell.row - 1), cell.col, cell.t, cell.s, cell.v));
    }
    
    if (!rowLength.containsKey(sheetNo)) {
      rowLength.put(sheetNo, new TreeMap<Long, Integer>());
    }
    Map<Long, Integer> rows = rowLength.get(sheetNo);
    if (!rows.containsKey(cell.row - 1)) {
      rows.put(cell.row - 1, cell.col);
    } else {
      rows.put(cell.row - 1, Math.max(rows.get(cell.row - 1), cell.col));
    }
  }
  
  @Override
  public String name() {
    return name;
  }
  
  @Override
  public String toString() {
    return name();
  }
  
  @Override
  public boolean isActive() {
    return active;
  }
  
  private class MyCell extends Cell {
    @SuppressWarnings("unused")
    public int sheet_id;
    public String value;
    
    @SuppressWarnings("unused")
    MyCell() {}
    
    MyCell(int sheet_id, int row, int col, String t, String s, String v) {
      this.sheet_id = sheet_id;
      this.row = row;
      this.col = col;
      this.t = t;
      this.s = s;
      this.v = v;
    }
    
    @SuppressWarnings("unused")
    MyCell(int row, int col) {
      this.row = row;
      this.col = col;
    }
    
    @Override
    public String asStr() {
      if (isStr()) return value;
      return v;
    }
    
    @Override
    public Date asDate() {
      if (isStr()) return null;
      return UtilOffice.excelToDate(asStr());
    }
    
    @Override
    public Integer asInt() {
      BigDecimal bd = asBigDecimal();
      return bd == null ? null :bd.intValue();
    }
    
    @Override
    public Long asLong() {
      BigDecimal bd = asBigDecimal();
      return bd == null ? null :bd.longValue();
    }
    
    @Override
    public BigDecimal asBigDecimal() {
      String s = value;
      if (isStr()) return null;
      if (s == null) return null;
      return new BigDecimal(s);
    }
    
  }
  
  @Override
  public void scanCells(CellHandler handler) {
    try {
      scanCellsEx(handler);
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException)e;
      }
      throw new RuntimeException(e);
    }
  }
  
  private void scanCellsEx(CellHandler handler) throws Exception {
    for (Map<Integer, MyCell> row : rows.values()) {
      for (MyCell cell : row.values()) {
        handler.handle(cell);
      }
    }
  }
  
  @Override
  public List<Cell> loadRow(int row) {
    try {
      return loadRowEx(row);
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException)e;
      }
      throw new RuntimeException(e);
    }
  }
  
  private List<Cell> loadRowEx(int row) throws SQLException {
    List<Cell> ret = new ArrayList<Cell>();
    if (rows.get((long)row) != null) {
      Map<Integer, MyCell> r = rows.get((long)row);
      Map<Long, Integer> m = rowLength.get(sheetNo);
      Integer maxColumn = m != null ? m.get((long)row) :null;
      if (maxColumn != null) {
        for (int i = 0; i <= maxColumn; i++) {
          MyCell myCell = r.get(i);
          if (myCell != null && myCell.isStr()) myCell.value = values.get(Long.parseLong(myCell.v));
          ret.add(myCell);
        }
      }
    }
    return ret;
  }
  
  @Override
  public void scanRows(int colCountInRow, RowHandler handler) {
    try {
      scanRowsEx(colCountInRow, handler);
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException)e;
      }
      throw new RuntimeException(e);
    }
  }
  
  private void scanRowsEx(int colCountInRow, RowHandler handler) throws Exception {
    for (Map.Entry<Long, Map<Integer, MyCell>> row : rows.entrySet()) {
      List<Cell> r = new ArrayList<Cell>();
      for (int i = 0; i < colCountInRow; i++) {
        MyCell myCell = row.getValue().get(i);
        if (myCell != null && myCell.isStr()) myCell.value = values.get(Long.parseLong(myCell.v));
        r.add(myCell);
      }
      handler.handle(r, row.getKey().intValue());
    }
  }
}
