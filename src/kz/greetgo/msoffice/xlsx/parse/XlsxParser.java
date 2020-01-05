package kz.greetgo.msoffice.xlsx.parse;

import kz.greetgo.msoffice.util.UtilOffice;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.File;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * <p>
 * Основной файл для парсинга xlsx-файлов
 * </p>
 * <p>
 * Пример использования:
 * </p>
 * <p>
 * <code>
 * XlsxParser p = new XlsxParser();
 * p.load(inputStreamOfXlsxFileContent);
 * <p>
 * p.loadRow();
 * p.scanCells();
 * p.scanRows();
 * <p>
 * //Необходимо обязательно закрывать парсер
 * p.close();
 * </code>
 * </p>
 *
 * @author pompei
 */
public class XlsxParser {
  private String tmpDirBase = System.getProperty("java.io.tmpdir", ".");
  private SAXParser saxParser;
  public boolean removeWorkDirOnClose = true;

  private final List<Sheet> sheets = new ArrayList<Sheet>();

  private SAXParser saxParser() throws Exception {
    if (saxParser == null) {
      SAXParserFactory fact = SAXParserFactory.newInstance();
      saxParser = fact.newSAXParser();
    }
    return saxParser;
  }

  public void setTmpDirBase(String tmpDirBase) {
    if (workDir != null) {
      throw new IllegalStateException("Cannot change this property on active parser. "
        + "Call method 'close' to deactivate this parser");
    }
    this.tmpDirBase = tmpDirBase;
  }

  private String workDir = null;
  private Connection connection;

  private String workDir() {
    if (workDir == null) {
      String newName = "xlsxParser-" + System.currentTimeMillis() + '-' + new Random().nextLong();
      //newName = newName.replaceAll("-", "");
      workDir = tmpDirBase + '/' + newName;
      new File(workDir).mkdirs();
    }
    return workDir;
  }

  public void closeEx() throws Exception {
    if (workDir == null) return;
    if (connection != null) {
      connection.rollback();
      connection.close();
      try {
        connection = DriverManager.getConnection("jdbc:derby:" + workDir() + "/db;shutdown=true");
      } catch (SQLException sqlError) {
        if (!"08006".equals(sqlError.getSQLState())) throw sqlError;
      }
      connection.close();
    }
    if (removeWorkDirOnClose) {
      UtilOffice.removeDir(workDir);
    }
    workDir = null;
  }

  public void close() {
    try {
      closeEx();
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException) e;
      }
      throw new RuntimeException(e);
    }
  }

  public void load(InputStream in) {
    try {
      loadEx(in);
    } catch (Exception e) {

      if (connection != null) {
        try {
          connection.close();
        } catch (SQLException ignore) {}
        connection = null;
      }

      if (e instanceof RuntimeException) {
        throw (RuntimeException) e;
      }
      throw new RuntimeException(e);
    }
  }

  public void loadEx(InputStream in) throws Exception {
    prepareDB();

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
        saxParser().parse(UtilOffice.copy(zin), new SharedStringsHandler(connection));
        zin.closeEntry();
        continue;
      }
      if (entry.getName().startsWith("xl/worksheets/") && entry.getName().endsWith(".xml")) {
        SheetHandler handler = new SheetHandler(connection, entry.getName(), sheetNo++);
        try {
          saxParser().parse(UtilOffice.copy(zin), handler);
        } finally {
          handler.finishInsertion();
        }
        zin.closeEntry();
        sheets.add(handler);
        continue;
      }
      zin.closeEntry();
    }

    zin.close();
  }

  private void prepareDB() throws Exception {
    Class.forName("org.apache.derby.jdbc.EmbeddedDriver");
    connection = DriverManager.getConnection("jdbc:derby:" + workDir() + "/db;create=true");

    {
      PreparedStatement ps = connection.prepareStatement(
        "create table strs (nom bigint not null primary key, value clob)");
      ps.execute();
      ps.close();
    }
    {
      PreparedStatement ps = connection.prepareStatement(UtilOffice.streamToStr(getClass()
        .getResourceAsStream("create_table_SHEETS.sql")));
      ps.execute();
      ps.close();
    }
    {
      PreparedStatement ps = connection.prepareStatement(UtilOffice.streamToStr(getClass()
        .getResourceAsStream("create_table_CELLS.sql")));
      ps.execute();
      ps.close();
    }
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
