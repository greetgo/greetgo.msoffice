package kz.greetgo.msoffice.xlsx.parse;

import java.util.List;

/**
 * Обрабочик сканирования строк
 * 
 * @author pompei
 */
public interface RowHandler {
  /**
   * Вызывается при появлении очередной строки
   * 
   * @param row
   *          список ячеек очередной строки
   * @param rowIndex
   *          индекс очередной строки: 0 - первая строка, 1 - вторая, и т.д.
   */
  void handle(List<Cell> row, int rowIndex) throws Exception;
}
