package kz.greetgo.msoffice.xlsx.parse;

import java.util.List;

/**
 * Лист электронной таблицы
 *
 * @author pompei
 */
public interface Sheet {
  /**
   * Внутренее имя листа
   *
   * @return имя
   */
  String name();

  /**
   * Проверяет лист на активность. Активным может быть только один лист
   *
   * @return true - активный, false - иначе
   */
  boolean isActive();

  /**
   * <p>
   * Сканирует ячейки. Сканирование производиться сверху-в-низ, слева-на-право (т.е. вначале все
   * ячейки слева-на-право первой сверху непустой строки, потом - второй сверху непустой и т.д.).
   * </p>
   * <p>
   * CellHandler предоставляет всегда один и тот-же объект, только с разными значениями.
   * </p>
   * <p>
   * Пустые ячейки не предоставляются.
   * </p>
   *
   * @param handler обработчик новой ячейки
   */
  void scanCells(CellHandler handler);

  /**
   * <p>
   * Загружает определённую строку
   * </p>
   *
   * <p>
   * Этот метод медленный (правда от величины номера строки скорость не зависит, т.е. первая строка
   * читается с такой-же скоростью как и последняя). Этот метод пригоден для получения нескольких
   * строк. Для полного сканирования следует использовать методы {@link #scanRows(int, RowHandler)}
   * или {@link #scanCells(CellHandler)}
   * </p>
   *
   * @param row индекс строки: самая верхнаяя имеет индекс 0, следующая - 1, потом - 2, и т.д.
   * @return список ячеек. Первый элемент списка (с индексом 0) предоставляет ячейку из колонки A,
   * второй элемент списка (с индексом 1) - из B, следующий - из C, и т.д.
   */
  List<Cell> loadRow(int row);

  /**
   * <p>
   * Сканирует строки сверху-в-низ
   * </p>
   *
   * <p>
   * RowHandler предоставляет всегда один и тот-же объект, только с разными значениями.
   * </p>
   *
   * @param colCountInRow запрашиваемый размер строки (остальные ячейки будут обрезаться и игнорироваться)
   * @param handler       обработчик новой строки
   */
  void scanRows(int colCountInRow, RowHandler handler);
}
