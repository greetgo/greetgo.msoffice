package kz.greetgo.msoffice.xlsx.xlsx_reader;

import kz.greetgo.msoffice.UtilOffice;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class XlsxReader implements AutoCloseable {

  private final StoredStrings storedStrings;

  public XlsxReader() {
    try {
      storedStrings = new StoredStrings(
        Files.createTempFile("stored-strings-ref", ".bin"),
        Files.createTempFile("stored-strings-content", ".bin")
      );
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  @Override
  public void close() {
    storedStrings.close();
  }

  private static ZipInputStream wrapInZip(InputStream inputStream) {
    return inputStream instanceof ZipInputStream
      ? (ZipInputStream) inputStream
      : new ZipInputStream(inputStream, StandardCharsets.UTF_8);
  }

  public void read(InputStream inputStream) {
    try (ZipInputStream zipInputStream = wrapInZip(inputStream)) {

      while (true) {
        ZipEntry entry = zipInputStream.getNextEntry();
        if (entry == null) {
          break;
        }
        try {

          if ("xl/sharedStrings.xml".equals(entry.getName())) {
            doParsing(zipInputStream, new StringsHandler());
            continue;
          }
          if ("xl/styles.xml".equals(entry.getName())) {
            doParsing(zipInputStream, new StylesHandler());
            continue;
          }

          System.out.println("j253bv235 :: " + entry.getName());
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

  private class StringsHandler extends AbstractXmlHandler {
    @Override
    protected void startTag(String tagPath, Attributes attributes) {}

    @Override
    protected void endTag(String tagPath) {
      if ("/sst/si/t".equals(tagPath)) {
        storedStrings.append(text());
      }
    }

  }

}
