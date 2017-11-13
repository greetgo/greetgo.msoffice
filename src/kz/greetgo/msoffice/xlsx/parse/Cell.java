package kz.greetgo.msoffice.xlsx.parse;

import java.math.BigDecimal;
import java.util.Date;

/**
 * Предоставлет доступ к определённой ячейке
 * 
 * @author pompei
 */
public abstract class Cell {
  int row, col;
  
  String s, t, v;
  
  void cleanData() {
    s = t = v = null;
  }
  
  @Override
  public String toString() {
    return asStr();
  }
  
  public boolean isStr() {
    return "s".equals(t);
  }
  
  public abstract String asStr();
  
  public abstract Date asDate();
  
  public abstract Integer asInt();
  
  public abstract Long asLong();
  
  public abstract BigDecimal asBigDecimal();
  
  /**
   * 
   * Получает индекс строки (0 - первая строка, 1 - вторая, ...)
   * 
   * 
   * @return индекс строки
   */
  public int row() {
    return row;
  }
  
  /**
   * Получает индекс колонки (0 - колонка А, 1 - колонка B, ...)
   * 
   * @return индекс колонки
   */
  public int col() {
    return col;
  }
  
}
