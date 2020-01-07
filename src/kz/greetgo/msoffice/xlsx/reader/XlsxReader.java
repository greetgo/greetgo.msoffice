package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.util.UtilOffice;
import kz.greetgo.msoffice.xlsx.reader.model.SheetData;
import kz.greetgo.msoffice.xlsx.reader.model.WorkbookData;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.DateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Random;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * Reads data from xlsx file and give random access to this data.
 * <p>
 * Data does NOT load to memory - data in row saved to a tmp file, because you must not get OutOfMemoryError
 * <p>
 * Please use try-with-resource statement using this class for example
 * <p>
 * <pre>
 * try (XlsxReader xlsxReader = new XlsxReader()) {
 *   xlsxReader.read(inputStream);
 *
 *   Sheet sheet = xlsxReader.tabSelectedSheet();
 *
 *   System.out.println("sheet name is " + sheet.name());
 *   System.out.println("sheet row count is " + sheet.rowCount());
 *   System.out.println(sheet.cell(3, 7).asText());
 *   System.out.println("" + sheet.cell(3, 7).borders().topStyle());
 *   // etc
 * }
 * </pre>
 */
public class XlsxReader implements AutoCloseable {

  private final XlsxReaderContext context;

  private final Map<Integer, SheetData> sheetDataMap = new HashMap<>();
  private final WorkbookData workbook = new WorkbookData();

  private final Path tempDir;
  private Random random = new Random();
  private final List<Path> tmpFiles = new ArrayList<>();

  private final Map<Integer, SheetReader> sheetReaderMap = new HashMap<>();
  private Sheet tabSelectedSheet = null;

  private Path tmp(Path tmp) {
    tmpFiles.add(tmp);
    return tmp;
  }

  @SuppressWarnings("FinalPrivateMethod")
  private final Path createTmpFile(String prefixName) {

    if (tempDir == null) {
      try {
        return tmp(Files.createTempFile(prefixName, ".bin"));
      } catch (IOException e) {
        throw new RuntimeException(e);
      }
    }

    if (random == null) random = new Random();

    return tmp(tempDir.resolve(prefixName + "-" + Math.abs(random.nextLong()) + ".bin"));
  }

  public XlsxReader() {
    this(null);
  }

  public XlsxReader(Path tempDir) {
    this.tempDir = tempDir;

    StoredStrings storedStrings = new StoredStrings(
      createTmpFile("stored-strings-ref"), createTmpFile("stored-strings-content"));

    context = new XlsxReaderContext(storedStrings);
  }

  @Override
  public void close() {
    context.storedStrings.close();
    tmpFiles.stream().map(Path::toFile).forEach(File::delete);
  }

  private static ZipInputStream wrapInZip(InputStream inputStream) {
    return inputStream instanceof ZipInputStream
      ? (ZipInputStream) inputStream
      : new ZipInputStream(inputStream, StandardCharsets.UTF_8);
  }

  private static final Pattern SHEET_FILE = Pattern.compile("xl/worksheets/sheet(\\d+)\\.xml");

  public void read(InputStream inputStream) {
    try (ZipInputStream zipInputStream = wrapInZip(inputStream)) {

      while (true) {
        ZipEntry entry = zipInputStream.getNextEntry();
        if (entry == null) {
          break;
        }
        String entryName = entry.getName();

        try {
          if ("xl/sharedStrings.xml".equals(entryName)) {
            doParsing(zipInputStream, new StringsHandler(context.storedStrings));
            continue;
          }
          if ("xl/styles.xml".equals(entryName)) {
            doParsing(zipInputStream, new StylesHandler(context.styles));
            continue;
          }
          if ("xl/workbook.xml".equals(entryName)) {
            doParsing(zipInputStream, new WorkbookHandler(workbook));
            continue;
          }

          {
            Matcher matcher = SHEET_FILE.matcher(entryName);
            if (matcher.matches()) {
              SheetData sheetData = new SheetData(Integer.parseInt(matcher.group(1)), this::createTmpFile);
              sheetDataMap.put(sheetData.id, sheetData);
              doParsing(zipInputStream, new SheetHandler(sheetData));
              continue;
            }
          }

//          System.out.println("j253bv235 :: " + entryName);
        } finally {
          zipInputStream.closeEntry();
        }

      }

    } catch (RuntimeException e) {
      throw e;
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }

  private static void doParsing(InputStream inputStream, ContentHandler contentHandler) throws SAXException, IOException {
    //noinspection deprecation
    XMLReader reader = XMLReaderFactory.createXMLReader();
    reader.setContentHandler(contentHandler);
    reader.parse(new InputSource(UtilOffice.copy(inputStream)));
  }

  public int sheetCount() {
    return workbook.sheetRefList.size();
  }

  public void setDateFormat(DateFormat dateFormat) {
    Objects.requireNonNull(dateFormat);
    context.dateFormat = dateFormat;
  }

  public Sheet sheet(int index) {
    {
      SheetReader sheetReader = sheetReaderMap.get(index);
      if (sheetReader != null) return sheetReader;
    }
    {
      SheetRef ref = workbook.sheetRefList.get(index);
      SheetData sheetData = sheetDataMap.get(ref.id);
      Objects.requireNonNull(sheetData, "index = " + index + ", sheet id = " + ref.id);
      SheetReader sheetReader = new SheetReader(context, ref, sheetData);
      sheetReaderMap.put(index, sheetReader);
      return sheetReader;
    }
  }

  public Sheet tabSelectedSheet() {
    if (tabSelectedSheet != null) {
      return tabSelectedSheet;
    }

    {
      for (int i = 0; i < sheetCount(); i++) {
        Sheet sheet = sheet(i);
        if (sheet.tabSelected()) {
          return tabSelectedSheet = sheet;
        }
      }
      throw new RuntimeException("No tab selected sheet: sheet count = " + sheetCount());
    }
  }

  @SuppressWarnings("SameParameterValue")
  void printSharedStrings(PrintStream out) {
    long strCount = context.storedStrings.strCount();
    if (strCount == 0) return;
    int len = ("" + (strCount - 1)).length();

    for (int i = 0; i < strCount; i++) {
      out.println(UtilOffice.toLenLeft(len, "" + i, " ") + " - " + context.storedStrings.get(i));
    }
  }
}
