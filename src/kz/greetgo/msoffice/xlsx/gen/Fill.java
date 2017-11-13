package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

/**
 * <p>
 * Заливка ячейки
 * </p>
 * 
 * <p>
 * Смотрите {@link PatternFill}
 * </p>
 * 
 * @author pompei
 */
public abstract class Fill {
  abstract Fill copy();
  
  abstract void print(PrintStream out);
  
  public static Fill copy(Fill fill) {
    if (fill == null) return null;
    return fill.copy();
  }
}
