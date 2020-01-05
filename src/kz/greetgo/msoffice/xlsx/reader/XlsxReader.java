package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.util.UtilOffice;
import kz.greetgo.msoffice.xlsx.reader.model.SheetData;
import kz.greetgo.msoffice.xlsx.reader.model.StylesData;
import kz.greetgo.msoffice.xlsx.reader.model.WorkbookData;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
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

public class XlsxReader implements AutoCloseable {

  private final StoredStrings storedStrings;
  private final StylesData styles = new StylesData();

  private final Map<Integer, SheetData> sheetDataMap = new HashMap<>();
  private final WorkbookData workbook = new WorkbookData();

  private final Path tempDir;
  private Random random = new Random();
  private final List<Path> tmpFiles = new ArrayList<>();

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
    storedStrings = new StoredStrings(createTmpFile("stored-strings-ref"), createTmpFile("stored-strings-content"));
  }

  @Override
  public void close() {
    storedStrings.close();
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
            doParsing(zipInputStream, new StringsHandler(storedStrings));
            continue;
          }
          if ("xl/styles.xml".equals(entryName)) {
            doParsing(zipInputStream, new StylesHandler(styles));
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
    XMLReader reader = XMLReaderFactory.createXMLReader();
    reader.setContentHandler(contentHandler);
    reader.parse(new InputSource(UtilOffice.copy(inputStream)));
  }

  public int sheetCount() {
    return workbook.sheetRefList.size();
  }

  private final Map<Integer, SheetReader> sheetReaderMap = new HashMap<>();

  public Sheet sheet(int index) {
    {
      SheetReader sheetReader = sheetReaderMap.get(index);
      if (sheetReader != null) return sheetReader;
    }
    {
      SheetRef ref = workbook.sheetRefList.get(index);
      SheetData sheetData = sheetDataMap.get(ref.id);
      Objects.requireNonNull(sheetData, "index = " + index + ", sheet id = " + ref.id);
      SheetReader sheetReader = new SheetReader(styles, storedStrings, ref, sheetData);
      sheetReaderMap.put(index, sheetReader);
      return sheetReader;
    }
  }

}
