package kz.greetgo.msoffice.xlsx.reader;

import org.testng.annotations.Test;

import java.io.InputStream;

import static org.fest.assertions.api.Assertions.assertThat;

public class XlsxReaderTest {

  @Test
  public void ok() {
    InputStream inputStream = getClass().getResourceAsStream("xlsx/ok.xlsx");
    try (XlsxReader xlsxReader = new XlsxReader()) {
      xlsxReader.read(inputStream);
    }
  }

  @Test
  public void book1() {
    InputStream inputStream = getClass().getResourceAsStream("xlsx/book-1.xlsx");
    try (XlsxReader xlsxReader = new XlsxReader()) {
      xlsxReader.read(inputStream);

      assertThat(xlsxReader.sheetCount()).isEqualTo(1);

      Sheet sheet = xlsxReader.sheet(0);

      assertThat(sheet.rowCount()).isEqualTo(31);
      assertThat(sheet.row(7).cell(2).asText()).isEqualTo("B1");
      assertThat(sheet.cell(7, 2).asText()).isEqualTo("B1");

    }
  }

  @Test
  public void simple() {
    InputStream inputStream = getClass().getResourceAsStream("xlsx/simple.xlsx");
    try (XlsxReader xlsxReader = new XlsxReader()) {
      xlsxReader.read(inputStream);

      System.out.println("h4b325v3 :: sheet count = " + xlsxReader.sheetCount() + "\n");
      for (int i = 0; i < xlsxReader.sheetCount(); i++) {
        Sheet sheet = xlsxReader.sheet(i);
        System.out.println("h4b325v3 :: sheet index " + sheet.index());
        System.out.println("h4b325v3 ::       name  " + sheet.name());
        System.out.println("h4b325v3 ::                " + sheet.cell(0, 0).asText() + " (" + sheet.cell(0, 0).format() + ")" + " | " + sheet.cell(0, 1).asText() + " (" + sheet.cell(0, 1).format() + ")");
        System.out.println("h4b325v3 ::                " + sheet.cell(1, 0).asText() + " (" + sheet.cell(1, 0).format() + ")" + " | " + sheet.cell(1, 1).asText() + " (" + sheet.cell(1, 1).format() + ")");
        System.out.println();
      }

    }
  }

}
