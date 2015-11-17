package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

/**
 * Графический объект листа, связанный с двумя ячейками и перемещяющийся при их перемещении
 * (например при изменении ширины столбцов)
 */
abstract class TwoCellAnchor {
  
  // верхний левый угол
  protected final int col1; // колонка, нумерация с 1
  protected final int row1; // строка, нумерация с 1
  protected final long col1off; // отступ от начала колонки
  protected final long row1off; // отступ от начала строки
  
  // правый нижний угол
  protected final int col2; // колонка, нумерация с 1
  protected final int row2; // строка, нумерация с 1
  protected final long col2off; // отступ от начала колонки
  protected final long row2off; // отступ от начала строки
  
  public TwoCellAnchor(SheetCoord coordFrom, SheetCoord coordTo) {
    
    col1 = coordFrom.col;
    col1off = coordFrom.coloff;
    row1 = coordFrom.row;
    row1off = coordFrom.rowoff;
    
    col2 = coordTo.col;
    col2off = coordTo.coloff;
    row2 = coordTo.row;
    row2off = coordTo.rowoff;
  }
  
  abstract int getRelId();
  
  abstract String getType();
  
  abstract void print(PrintStream os);
}