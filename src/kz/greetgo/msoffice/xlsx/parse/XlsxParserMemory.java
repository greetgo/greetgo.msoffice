package kz.greetgo.msoffice.xlsx.parse;

import kz.greetgo.msoffice.util.UtilOffice;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * <p>
 * Основной файл для парсинга xlsx-файлов
 * </p>
 * <p>
 * Пример использования:
 * </p>
 *
 * {@code XlsxParser p = new XlsxParser();}<br>
 * {@code p.load(inputStreamOfXlsxFileContent);}<br>
 * <br>
 * {@code p.loadRow();}<br>
 * {@code p.scanCells();}<br>
 * {@code p.scanRows();}<br>
 * <br>
 * {@code //Необходимо обязательно закрывать парсер}<br>
 * {@code p.close();}<br>
 *
 * @author zhandos
 */
public class XlsxParserMemory {
  private SAXParser saxParser;
  public boolean removeWorkDirOnClose = true;
  private final Map<Long, String> values = new TreeMap<Long, String>();

  private final List<Sheet> sheets = new ArrayList<Sheet>();

  private SAXParser saxParser() throws Exception {
    if (saxParser == null) {
      SAXParserFactory fact = SAXParserFactory.newInstance();
      saxParser = fact.newSAXParser();
    }
    return saxParser;
  }

  public void load(InputStream in) {
    try {
      loadEx(in);
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException) e;
      }
      throw new RuntimeException(e);
    }
  }

  public void loadEx(InputStream in) throws Exception {

    final ZipInputStream zin;
    if (in instanceof ZipInputStream) {
      zin = (ZipInputStream) in;
    } else {
      zin = new ZipInputStream(in);
    }

    ZipEntry entry;
    int sheetNo = 1;
    while ((entry = zin.getNextEntry()) != null) {
      if ("xl/sharedStrings.xml".equals(entry.getName())) {
        saxParser().parse(UtilOffice.copy(zin), new SharedStringsHandlerMemory(values));
        zin.closeEntry();
        continue;
      }
      if (entry.getName().startsWith("xl/worksheets/") && entry.getName().endsWith(".xml")) {
        SheetHandlerMemory handler = new SheetHandlerMemory(entry.getName(), sheetNo++, values);
        saxParser().parse(UtilOffice.copy(zin), handler);
        zin.closeEntry();
        sheets.add(handler);
        continue;
      }
      zin.closeEntry();
    }

    zin.close();
  }

  public List<Sheet> sheets() {
    return sheets;
  }

  public Sheet activeSheet() {
    for (Sheet sheet : sheets) {
      if (sheet.isActive()) return sheet;
    }
    if (sheets.size() > 0) return sheets.get(0);
    return null;
  }

}
