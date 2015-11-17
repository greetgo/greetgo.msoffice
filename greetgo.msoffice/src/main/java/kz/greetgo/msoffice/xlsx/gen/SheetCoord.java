package kz.greetgo.msoffice.xlsx.gen;

import kz.greetgo.msoffice.UtilOffice;

/**
 * Координаты на листе. Используются в графиках и для изображений.
 */
public class SheetCoord {
  
  /** Столбец ячейки, нумерация с 1 */
  public int col;
  /** Строка ячейки, нумерация с 1 */
  public int row;
  /** Отступ по горизонтали от ячейки */
  public int coloff;
  /** Отступ по вертикали от ячейки */
  public int rowoff;
  
  public SheetCoord() {}
  
  public SheetCoord(int col, int row) {
    this.col = col;
    this.row = row;
  }
  
  public SheetCoord(String coord) {
    int[] coordp = UtilOffice.parseCellCoord(coord);
    col = coordp[0];
    row = coordp[1];
  }
  
  public SheetCoord(int col, int row, int coloff, int rowoff) {
    this.col = col;
    this.row = row;
    this.coloff = coloff;
    this.rowoff = rowoff;
  }
  
  public SheetCoord(String coord, int coloff, int rowoff) {
    int[] coordp = UtilOffice.parseCellCoord(coord);
    col = coordp[0];
    row = coordp[1];
    this.coloff = coloff;
    this.rowoff = rowoff;
  }
  
  public SheetCoord(String col, int row) {
    this.col = UtilOffice.parseLettersNumber(col) + 1;
    this.row = row;
  }
  
  public SheetCoord(String col, int row, int coloff, int rowoff) {
    this.col = UtilOffice.parseLettersNumber(col) + 1;
    this.row = row;
    this.coloff = coloff;
    this.rowoff = rowoff;
  }
  
  @Override
  public String toString() {
    
    StringBuilder builder = new StringBuilder();
    
    builder.append("SheetCoord [col=");
    builder.append(col);
    builder.append(", row=");
    builder.append(row);
    builder.append(", coloff=");
    builder.append(coloff);
    builder.append(", rowoff=");
    builder.append(rowoff);
    builder.append("]");
    
    return builder.toString();
  }
}