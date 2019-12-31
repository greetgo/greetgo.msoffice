package kz.greetgo.msoffice.xlsx.parse;

import kz.greetgo.msoffice.UtilOffice;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.math.BigDecimal;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class SheetHandler extends DefaultHandler implements Sheet {
  private final Connection connection;
  private final int sheetNo;
  private final String name;
  private boolean active;

  public SheetHandler(Connection connection, String sheetPathName, int sheetNo) {
    this.connection = connection;
    this.sheetNo = sheetNo;

    {
      int startLen = "xl/worksheets/".length();
      int endLen = ".xml".length();
      name = sheetPathName.substring(startLen, sheetPathName.length() - endLen);

      try {
        PreparedStatement ps = connection.prepareStatement("insert into sheets (id, name) values (?, ?)");
        try {
          ps.setInt(1, sheetNo);
          ps.setString(2, name);
          ps.executeUpdate();
          connection.commit();
        } finally {
          ps.close();
        }
      } catch (SQLException e) {
        throw new RuntimeException(e);
      }
    }
  }

  private final XmlIn in = new XmlIn();

  PreparedStatement cellsInsertPS = null;
  private int currentBachSize = 0;
  private static final int MAX_BATCH_SIZE = 2000;

  @Override
  public void startDocument() throws SAXException {
    try {
      cellsInsertPS = connection
        .prepareStatement("insert into cells (sheet_id, n_row, n_col, t, s, v) values (?, ?, ?, ?, ?, ?)");
      currentBachSize = 0;
    } catch (SQLException e) {
      throw new RuntimeException(e);
    }
  }

  @Override
  public void endDocument() throws SAXException {
    finishInsertion();
  }

  public void finishInsertion() {
    if (cellsInsertPS == null) return;
    try {
      if (currentBachSize > 0) {
        commit();
      }
      cellsInsertPS.close();
      cellsInsertPS = null;
    } catch (SQLException e) {
      throw new RuntimeException(e);
    }
  }

  @Override
  public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
    textBuilder = null;
    in.stepIn(qName);

    if ("worksheet/sheetData/row/c".equals(in.current())) {
      cell.r = attributes.getValue("r");
      cell.t = attributes.getValue("t");
      cell.s = attributes.getValue("s");
    }
    if ("worksheet/sheetViews/sheetView".equals(in.current())) {
      if ("1".equals(attributes.getValue("tabSelected"))) {
        setSheetIsActive(true);
      }
    }
  }

  private static class CellInfo {
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
      appendCellEx(cell);
    } catch (SQLException e) {
      throw new RuntimeException(e);
    }
  }

  private void appendCellEx(CellInfo cell) throws SQLException {
    cell.calcRowCol();
    {
      cellsInsertPS.setLong(1, sheetNo);
      cellsInsertPS.setLong(2, cell.row - 1);
      cellsInsertPS.setInt(3, cell.col);
      cellsInsertPS.setString(4, cell.t);
      cellsInsertPS.setString(5, cell.s);
      cellsInsertPS.setString(6, cell.v);
    }
    cellsInsertPS.addBatch();
    currentBachSize++;
    if (currentBachSize >= MAX_BATCH_SIZE) {
      commit();
    }
  }

  private void commit() throws SQLException {
    cellsInsertPS.executeBatch();
    connection.commit();
    currentBachSize = 0;
  }

  private void setSheetIsActive(boolean active) {
    try {
      setSheetIsActiveEx(active);
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException) e;
      }
      throw new RuntimeException(e);
    }
  }

  private void setSheetIsActiveEx(boolean active) throws Exception {
    try (PreparedStatement ps = connection.prepareStatement("update sheets set is_active = ? where id = ?")) {
      ps.setInt(1, active ? 1 : 0);
      ps.setLong(2, sheetNo);
      ps.executeUpdate();
      connection.commit();
    }
    this.active = active;
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
    MyCell() {}

    MyCell(int row, int col) {
      this.row = row;
      this.col = col;
    }

    @Override
    public String asStr() {
      if (isStr()) return getStrByNo(Long.parseLong(v));
      return v;
    }

    @Override
    public Date asDate() {
      if (isStr()) return null;
      return UtilOffice.excelToDate(v);
    }

    @Override
    public Integer asInt() {
      BigDecimal bd = asBigDecimal();
      return bd == null ? null : bd.intValue();
    }

    @Override
    public Long asLong() {
      BigDecimal bd = asBigDecimal();
      return bd == null ? null : bd.longValue();
    }

    @Override
    public BigDecimal asBigDecimal() {
      if (isStr()) return null;
      if (v == null) return null;
      return new BigDecimal(v);
    }

    public void read(ResultSet rs) throws SQLException {
      row = rs.getInt("n_row");
      col = rs.getInt("n_col");
      readData(rs);
    }

    public void readData(ResultSet rs) throws SQLException {
      s = rs.getString("s");
      t = rs.getString("t");
      v = rs.getString("v");
    }
  }

  @Override
  public void scanCells(CellHandler handler) {
    try {
      scanCellsEx(handler);
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException) e;
      }
      throw new RuntimeException(e);
    }
  }

  private void scanCellsEx(CellHandler handler) throws Exception {
    PreparedStatement ps = connection.prepareStatement(//
      "select * from cells where sheet_id = ? order by sheet_id, n_row, n_col");
    try {
      ps.setInt(1, sheetNo);
      ResultSet rs = ps.executeQuery();
      try {
        MyCell cell = new MyCell();
        while (rs.next()) {
          cell.read(rs);
          handler.handle(cell);
        }
      } finally {
        rs.close();
      }
    } finally {
      ps.close();
    }
    connection.commit();
  }

  private String getStrByNo(long strNo) {
    try {
      return getStrByNoEx(strNo);
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException) e;
      }
      throw new RuntimeException(e);
    }
  }

  private String getStrByNoEx(long strNo) throws SQLException {
    PreparedStatement ps = connection.prepareStatement("select value from strs where nom = ?");
    try {
      ps.setLong(1, strNo);
      ResultSet rs = ps.executeQuery();
      try {
        String ret = null;
        if (rs.next()) ret = rs.getString(1);
        return ret;
      } finally {
        rs.close();
      }
    } finally {
      ps.close();
    }
  }

  @Override
  public List<Cell> loadRow(int row) {
    try {
      return loadRowEx(row);
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException) e;
      }
      throw new RuntimeException(e);
    }
  }

  private List<Cell> loadRowEx(int row) throws SQLException {
    List<Cell> ret = new ArrayList<Cell>();
    PreparedStatement ps = connection
      .prepareStatement("select * from cells where sheet_id = ? and n_row = ?"
        + " order by sheet_id, n_row, n_col");
    try {
      ps.setInt(1, sheetNo);
      ps.setInt(2, row);
      ResultSet rs = ps.executeQuery();
      try {
        int lastCol = -1;
        while (rs.next()) {
          MyCell cell = new MyCell();
          cell.read(rs);
          while (++lastCol < cell.col) {
            ret.add(new MyCell(row, lastCol));
          }
          ret.add(cell);
        }
      } finally {
        rs.close();
      }
    } finally {
      ps.close();
    }
    connection.commit();
    return ret;
  }

  @Override
  public void scanRows(int colCountInRow, RowHandler handler) {
    try {
      scanRowsEx(colCountInRow, handler);
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException) e;
      }
      throw new RuntimeException(e);
    } finally {
      try {
        connection.commit();
      } catch (SQLException e) {
        e.printStackTrace();
      }
    }
  }

  private void scanRowsEx(int colCountInRow, RowHandler handler) throws Exception {
    List<Cell> row = new ArrayList<Cell>();
    for (int i = 0; i < colCountInRow; i++) {
      MyCell cell = new MyCell();
      cell.col = i;
      row.add(cell);
    }
    PreparedStatement ps = connection.prepareStatement(//
      "select * from cells where sheet_id = ? order by sheet_id, n_row, n_col");
    try {
      ps.setInt(1, sheetNo);
      ResultSet rs = ps.executeQuery();
      try {
        int currRowIndex = -1;
        int currColIndex = 213213;
        while (rs.next()) {
          int colIndex = rs.getInt("n_col");
          if (colIndex >= row.size()) continue;
          int rowIndex = rs.getInt("n_row");
          if (currRowIndex < rowIndex) {
            currColIndex = -1;
            if (currRowIndex > -1) {
              handler.handle(row, currRowIndex);
            }
            for (Cell cell : row) {
              cell.cleanData();
            }
            while (++currRowIndex < rowIndex) {
              for (Cell cell : row) {
                cell.row = currRowIndex;
              }
              handler.handle(row, currRowIndex);
            }
            assert currRowIndex == rowIndex;
            for (Cell cell : row) {
              cell.row = rowIndex;
            }
          }
          while (++currColIndex < colIndex) {
            row.get(currColIndex).cleanData();
          }
          ((MyCell) row.get(colIndex)).readData(rs);
        }
        handler.handle(row, currRowIndex);
      } finally {
        rs.close();
      }
    } finally {
      ps.close();
    }
  }
}
