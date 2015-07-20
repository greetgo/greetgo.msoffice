package kz.greetgo.msoffice.xlsx.gen;

import kz.greetgo.msoffice.UtilOffice;

/**
 * Координаты на листе. Используются в графиках.
 */
public class SheetCoord {
  
  /** Столбец ячейки */
  public int col;
  /** Строка ячейки */
  public int row;
  /** Отступ по горизонтали от ячейки */
  public int coloff;
  /** Отступ по вертикали от ячейки */
  public int rowoff;
  
  public SheetCoord(int col, int row) {
    
    this.col = col;
    this.row = row;
  }
  
  public SheetCoord(int col, int row, int coloff, int rowoff) {
    
    this.col = col;
    this.row = row;
    this.coloff = coloff;
    this.rowoff = rowoff;
  }
  
  public SheetCoord(String col, int row) {
    
    this.col = UtilOffice.parseLettersNumber(col);
    this.row = row;
  }
  
  public SheetCoord(String col, int row, int coloff, int rowoff) {
    
    this.col = UtilOffice.parseLettersNumber(col);
    this.row = row;
    this.coloff = coloff;
    this.rowoff = rowoff;
  }
}