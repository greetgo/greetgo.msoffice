package kz.greetgo.msoffice;

import java.awt.*;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.MathContext;
import java.nio.charset.StandardCharsets;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class UtilOffice {
  /**
   * Стандартное представление даты и времени по w3c
   */
  @SuppressWarnings("SpellCheckingInspection")
  public static final SimpleDateFormat W3CDTF = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'");

  /**
   * Количество миллисекунд в сутках
   */
  public static final int MILLIS_IN_DAY = 24 * 60 * 60 * 1000;

  /**
   * Преобразует дату в формат W3CDTF
   *
   * @param date исходная дата
   * @return строка даты в формате W3CDTF, или пустая строка если на вход пришел null
   */
  @SuppressWarnings("SpellCheckingInspection")
  public static String toW3CDTF(Date date) {
    if (date == null) return null;
    return W3CDTF.format(date);
  }

  /**
   * Преобразует строку формата W3CDTF в дату
   *
   * @param str строка содержащая дату в формате W3CDTF
   * @return дата или null, если str == null
   */
  @SuppressWarnings("SpellCheckingInspection")
  public static Date parseW3CDTF(String str) {
    if (str == null) return null;
    try {
      return W3CDTF.parse(str);
    } catch (ParseException e) {
      throw new RuntimeException(e);
    }
  }

  /**
   * Увеличивает длину строки до указанной добавлением нуля в начало строки
   *
   * @param len требуемая длина строки
   * @param s   исходная строка
   * @return результирующая строка
   */
  public static String toLenZero(int len, String s) {
    StringBuilder sb = new StringBuilder();
    if (s != null) sb.append(s);
    while (len > sb.length()) {
      sb.insert(0, ' ');
    }
    return sb.toString();
  }

  /**
   * Представляет цвет в виде шеснадцатиричной строки
   *
   * @param color исходный цвет
   * @return шеснадцатиричное представление цвета
   */
  public static String toHEX(Color color) {
    if (color == null) return "";
    return toLenZero(2, Integer.toHexString(color.getRed()).toUpperCase())
      + toLenZero(2, Integer.toHexString(color.getGreen()).toUpperCase())
      + toLenZero(2, Integer.toHexString(color.getBlue()).toUpperCase());
  }

  public static void appendToSB(InputStream in, StringBuilder sb) {
    try {
      appendToSB0(in, sb);
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }

  public static void appendToSB0(InputStream in, StringBuilder sb) throws Exception {
    BufferedReader br = new BufferedReader(new InputStreamReader(in, StandardCharsets.UTF_8));
    String line;
    while ((line = br.readLine()) != null) {
      sb.append(line);
      sb.append(System.getProperty("line.separator"));
    }
    br.close();
  }

  /**
   * Принимает поток как текстовый в кодировке UTF-8, и представляет его как строку
   *
   * @param in принимаемый поток. Будет считан до конца и закрыт.
   * @return полученная строка
   */
  public static String streamToStr(InputStream in) {
    StringBuilder sb = new StringBuilder();
    appendToSB(in, sb);
    return sb.toString();
  }

  @SuppressWarnings("unused")
  public static String streamToStr0(InputStream in) throws Exception {
    StringBuilder sb = new StringBuilder();
    appendToSB0(in, sb);
    return sb.toString();
  }

  public static void copyStreams(InputStream in, OutputStream out, int bufferSize)
    throws IOException {
    byte[] buffer = new byte[bufferSize];
    int readBytes;
    while ((readBytes = in.read(buffer)) != -1) {
      out.write(buffer, 0, readBytes);
    }
  }

  /**
   * Копирует данные из входного потока в выходной через буфер размером 2048 байт
   *
   * @param in  входной поток
   * @param out выходной поток
   * @throws IOException происходт в случае ошибки ввода/вывода
   */
  public static void copyStreams(InputStream in, OutputStream out) throws IOException {
    copyStreams(in, out, 2048);
  }

  /**
   * Удаляет слэш из начала строки, если он там есть, иначе ни чего не делает
   *
   * @param s исходная строка
   * @return результирующая строка
   */
  public static String killFirstSlash(String s) {
    if (s == null) return null;
    if (!s.startsWith("/")) return s;
    return s.substring(1);
  }

  /**
   * <p>
   * Вырезает из полного пути с именем файла только его имя. Если в конце слэш, то он игнорируется
   * </p>
   * <p>
   * Например:
   * </p>
   * <code>
   * asd/asd/wow.xml -&gt; wow.xml<br>
   * /wow/asd/dsa/uu/mama/ -&gt; mama
   * </code>
   *
   * @param fullName полный путь к файлу или папке
   * @return только имя файла или папки
   */
  public static String extractBaseName(String fullName) {
    if (fullName.endsWith("/")) {
      fullName = fullName.substring(0, fullName.length() - 1);
    }
    int index = fullName.lastIndexOf('/');
    if (index < 0) return fullName;
    return fullName.substring(index + 1);
  }

  /**
   * Удаляет папку вместе со всем её содержимым
   *
   * @param dir путь к удаляемой папке (абсолютный или относительный)
   * @return признак произведения операции: если файловая система не изменилась, то возвращается
   * лож; если же хоть что-то удалилось, то возвращается истина
   */
  public static boolean removeDir(String dir) {
    cleanDir(dir);
    return new File(dir).delete();
  }

  /**
   * Очищает папку - удаляет все её файлы и подпапки
   *
   * @param dir путь к очищаемой папке (абсолютный или относительный)
   * @return признак произведения операции: если файловая система не изменилась, то возвращается
   * лож; если же хоть что-то удалилось, то возвращается истина
   */
  @SuppressWarnings("UnusedReturnValue")
  public static boolean cleanDir(String dir) {
    String[] subNames = new File(dir).list();
    if (subNames == null) return false;
    boolean ret = false;
    for (String name : subNames) {
      if (".".equals(name)) continue;
      if ("..".equals(name)) continue;
      File f = new File(dir + System.getProperty("file.separator") + name);
      if (f.isFile()) {
        boolean q = f.delete();
        //ни в коем случае нельзя сокращать переменную q
        ret = ret || q;
        continue;
      }
      if (f.isDirectory()) {
        boolean q = removeDir(f.getPath());
        //ни в коем случае нельзя сокращать переменную q
        ret = ret || q;
        continue;
      }
      throw new LeftFileException("Left file " + f);
    }
    return ret;
  }

  private static final String[] LETTER_BASE = new String[]{"A", "B", "C", "D", "E", "F", "G",
    "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};

  /**
   * Формирует "буквенное" число: 0-&gt;A, 1-&gt;B, ..., 25-&gt;Z, 26-&gt;AA, 27-&gt;AB, .... (Такие числа
   * используются для обозначения колонок в электронных таблицах)
   *
   * @param n исходное число (должно быть больше или равно нулю)
   * @return результирующее "буквенное" число
   */
  public static String toLettersNumber(int n) {
    StringBuilder ret = new StringBuilder();
    appendLettersNumber(ret, n);
    return ret.toString();
  }

  public static void appendLettersNumber(StringBuilder sb, int n) {
    if (n < 0) throw new IllegalArgumentException("n must be >= 0");

    int base = LETTER_BASE.length;
    sb.append(LETTER_BASE[n % base]);
    n = n / base;

    while (n > 0) {
      sb.insert(0, LETTER_BASE[(n - 1) % base]);
      n = n / (base + 1);
    }
  }

  /**
   * Парсит буквенное число, преобразуя его в обычное: A-&gt;0, B-&gt;1, C-&gt;2, ..., Z-&gt;25, AA-&gt;26, AB-&gt;27,
   * ..., ZZ-&gt;701, AAA-&gt;702, ... (Такие числа используются для обозначения колонок в электронных
   * таблицах)
   *
   * @param s "буквенное" число
   * @return Распарсеное число
   */
  public static int parseLettersNumber(String s) {
    if (s == null) {
      throw new NullPointerException("parseLettersNumber cannot parse null string");
    }
    s = s.toUpperCase();
    if (s.length() == 0) {
      throw new IllegalArgumentException("parseLettersNumber cannot parse empty string");
    }
    int ret = s.charAt(0) - 'A';
    for (int i = 1, C = s.length(); i < C; i++) {
      int current = s.charAt(i) - 'A';
      ret = 26 * (ret + 1) + current;
    }
    return ret;
  }

  /**
   * Парсит представление координат ячейки.
   *
   * @param coordinate Строка с координатой: символьное обозначение столбца и номер строки ("B6", "C8", "AB12").
   * @return [ col, row ], нумерация с 1
   */
  public static int[] parseCellCoordinate(String coordinate) {
    if (coordinate == null) throw new NullPointerException("Координата не должна быть пустой");
    coordinate = coordinate.trim();
    if (coordinate.length() < 1) throw new IllegalArgumentException("Координата не должна быть пустой");

    int[] coordinateInt = new int[2];

    coordinate = coordinate.toUpperCase();

    int i = 0;
    while (i < coordinate.length()) {
      if (Character.isDigit(coordinate.charAt(i))) break;
      i++;
    }

    String col = coordinate.substring(0, i);
    String row = "0";
    if (i < coordinate.length()) row = coordinate.substring(i);

    coordinateInt[0] = parseLettersNumber(col) + 1;
    coordinateInt[1] = Integer.parseInt(row);

    return coordinateInt;
  }

  /**
   * <p>
   * Формирует копию InputStream
   * </p>
   * <p>
   * Необходим например для предотвращения закрытия оригинального потока, при закрытии копии
   * </p>
   *
   * @param in оригинальный поток
   * @return копия
   */
  public static InputStream copy(final InputStream in) {
    return new InputStream() {
      @Override
      public int read() throws IOException {
        return in.read();
      }

      @Override
      public int read(byte[] b, int off, int len) throws IOException {
        return in.read(b, off, len);
      }
    };
  }

  /**
   * <p>
   * Магическое число для преобразования excel-евского представления даты в java-кое представление
   * (и обратно)
   * </p>
   * <p>
   * <p>
   * это магическое число выведено подбором
   * </p>
   */
  private static final int EXCEL_DATE_MAGIC = 70 * 365 + 50;

  public static Calendar epochStart() {
    Calendar c = new GregorianCalendar();
    c.set(Calendar.YEAR, 1970);
    c.set(Calendar.MONTH, 1);
    c.set(Calendar.DAY_OF_MONTH, 1);
    c.set(Calendar.HOUR_OF_DAY, 0);
    c.set(Calendar.MINUTE, 0);
    c.set(Calendar.SECOND, 0);
    c.set(Calendar.MILLISECOND, 0);
    return c;
  }

  /**
   * Преобразует excel-евское представление даты-времени в java-представление (java.util.Date)
   *
   * @param excelValue excel-евское представление даты-времени
   * @return java-представление даты-времени
   */
  public static Date excelToDate(String excelValue) {
    if (excelValue == null) return null;
    Calendar c = epochStart();

    BigDecimal value = new BigDecimal(excelValue);

    c.add(Calendar.DAY_OF_YEAR, value.setScale(0, BigDecimal.ROUND_DOWN).intValueExact() - EXCEL_DATE_MAGIC);

    BigDecimal afterZero = value.subtract(value.setScale(0, BigDecimal.ROUND_DOWN));

    BigDecimal millis = afterZero.multiply(new BigDecimal(MILLIS_IN_DAY));

    millis = millis.setScale(0, BigDecimal.ROUND_HALF_UP);

    return new Date(c.getTimeInMillis() + millis.longValueExact());
  }

  /**
   * Пребразует java-представление даты-времени в excel-евское
   *
   * @param date java-представление даты-времени
   * @return excel-евское представление даты-времени в виде строки
   */
  public static String toExcelDateTime(Date date) {
    if (date == null) return null;

    Calendar c = epochStart();

    long totalMillis = date.getTime() - c.getTimeInMillis();

    int days = (int) (totalMillis / MILLIS_IN_DAY) + EXCEL_DATE_MAGIC;

    c.add(Calendar.DAY_OF_YEAR, days - EXCEL_DATE_MAGIC);

    int millis = (int) (date.getTime() - c.getTimeInMillis());

    BigDecimal afterZero = new BigDecimal(millis).divide(new BigDecimal(MILLIS_IN_DAY),
      MathContext.DECIMAL64);
    BigDecimal bdDays = new BigDecimal(days);
    BigDecimal result = bdDays.add(afterZero);

    return result.toString();
  }

  /**
   * <p>
   * Получает величину 16-чного символа
   * </p>
   * <p>
   * (0-&gt;0, 1-&gt;1, ..., 9-&gt;9, A-&gt;10, B-&gt;11, ..., F-&gt;15)
   * </p>
   *
   * @param c 16-чный символ
   * @return его величина
   */
  public static int charAsHex(char c) {
    if ('0' <= c && c <= '9') return c - '0';
    if ('A' <= c && c <= 'F') return c - 'A' + 10;
    if ('a' <= c && c <= 'f') return c - 'a' + 10;
    throw new IllegalArgumentException("Unknown hex digit " + c + " (code is " + (int) c + ")");
  }

  /**
   * Парсит строку как целое представленое в 16-чных цыфрах
   *
   * @param hex   входная строка
   * @param index индекс первого читаемого символа
   * @param count количество читаемых символов
   * @return полученое целое
   */
  public static int parsePartAsHex(String hex, int index, int count) {
    int ret = 0;
    int base = 1;
    for (; count-- > 0; ) {
      ret += base * charAsHex(hex.charAt(index + count));
      base *= 16;
    }
    return ret;
  }

  /**
   * <p>
   * Преобразует координаты ячейки таблицы в электронно-табличную форму
   * </p>
   * <p>
   * Например для <br>
   * (row,col) = (1,0) -&gt; "A1"<br>
   * (row,col) = (4,1) -&gt; "A4"<br>
   * (row,col) = (4,2) -&gt; "B4"<br>
   * (row,col) = (7,3) -&gt; "C7"<br>
   * (row,col) = (7,10) -&gt; "J7"<br>
   * </p>
   *
   * @param row индекс строки (1 - первая строка, 2 - вторая, ...)
   * @param col индекс колонки (0 - первая колонка, 1 - вторая, ...)
   * @return электронно-табличная форма координат
   */
  public static String toTablePosition(int row, int col) {
    return toLettersNumber(col) + row;
  }

  /**
   * Зиппует указанную папку со всем её содержимым рекурсивно и отправляет в указанный поток
   *
   * @param dir зиппуемая папка
   * @param out выводимый поток
   */
  public static void zipDir(String dir, OutputStream out) {
    try {
      zipDirEx(dir, out);
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException) e;
      }
      throw new RuntimeException(e);
    }
  }

  /**
   * Зиппует указанную папку со всем её содержимым рекурсивно и отправляет в указанный поток
   *
   * @param dir зиппуемая папка
   * @param out выводимый поток
   * @throws Exception утобы не делать try/catch-блок
   */
  public static void zipDirEx(String dir, OutputStream out) throws Exception {
    final ZipOutputStream zOut;
    if (out instanceof ZipOutputStream) {
      zOut = (ZipOutputStream) out;
    } else {
      zOut = new ZipOutputStream(out);
    }

    appendDir(dir, "", zOut);

    zOut.close();
  }

  private static void appendDir(String dir, String localPath, ZipOutputStream zOut)
    throws Exception {
    String[] list = new File(dir).list();
    if (list != null) for (String localName : list) {
      String fullName = dir + "/" + localName;
      String localFullName = localName;
      if (localPath.length() > 0) {
        localFullName = localPath + "/" + localName;
      }
      if (new File(fullName).isDirectory()) {
        appendDir(fullName, localFullName, zOut);
      } else {
        appendFile(fullName, localFullName, zOut);
      }
    }
  }

  private static void appendFile(String fullName, String localFullName, ZipOutputStream zOut)
    throws Exception {

    zOut.putNextEntry(new ZipEntry(localFullName));

    FileInputStream in = new FileInputStream(fullName);

    UtilOffice.copyStreams(in, zOut);

    in.close();

    zOut.closeEntry();
  }

  public static Integer strToInt(String str) {
    if (str == null) return null;
    return Integer.valueOf(str);
  }
}
