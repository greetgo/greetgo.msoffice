package kz.greetgo.msoffice.xlsx.reader;

import org.testng.annotations.Test;

import java.io.InputStream;

public class XlsxReaderTest {

  @Test
  public void ok() {
    InputStream inputStream = getClass().getResourceAsStream("ok.xlsx");
    try (XlsxReader xlsxReader = new XlsxReader()) {
      xlsxReader.read(inputStream);
    }
  }

  @Test
  public void book1() {
    InputStream inputStream = getClass().getResourceAsStream("book-1.xlsx");
    try (XlsxReader xlsxReader = new XlsxReader()) {
      xlsxReader.read(inputStream);
    }
  }

}
