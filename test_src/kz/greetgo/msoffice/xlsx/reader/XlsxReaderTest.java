package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.xlsx.reader.model.BorderStyle;
import kz.greetgo.msoffice.xlsx.reader.model.HorizontalAlign;
import kz.greetgo.msoffice.xlsx.reader.model.VerticalAlign;
import org.testng.annotations.Test;

import java.io.InputStream;

import static org.fest.assertions.api.Assertions.assertThat;

public class XlsxReaderTest {

  @Test
  public void ok() {
    InputStream inputStream = getClass().getResourceAsStream("xlsx/ok.xlsx");
    try (XlsxReader xlsxReader = new XlsxReader()) {
      xlsxReader.read(inputStream);

      Sheet sheet = xlsxReader.tabSelectedSheet();

      System.out.println("j436v6v27 :: tabSelectedSheet().cell(2, 0).borders().bottomStyle() = " + sheet.cell(2, 0).borders().bottomStyle());

      System.out.println("j436v6v27 :: tabSelectedSheet().name() = " + sheet.name());
      System.out.println("j436v6v27 :: tabSelectedSheet().frozenRowCount() = " + sheet.frozenRowCount());
      System.out.println("j436v6v27 :: tabSelectedSheet().rowCount() = " + sheet.rowCount());
      System.out.println("j436v6v27 :: tabSelectedSheet().cellMergeCount() = " + sheet.cellMergeCount());
      for (int i = 0; i < sheet.cellMergeCount(); i++) {
        System.out.println("j436v6v27 :: tabSelectedSheet().sheet.cellMerge(" + i + ") = " + sheet.cellMerge(i));
      }

      assertThat(sheet.frozenRowCount()).isEqualTo(3);
      assertThat(sheet.name()).isEqualTo("Это имя бизнес-объекта");
      assertThat(sheet.cell(2, 0).borders().bottomStyle()).isEqualTo(BorderStyle._double);
      assertThat(sheet.cell(2, 1).borders().bottomStyle()).isEqualTo(BorderStyle._double);
      assertThat(sheet.cell(2, 2).borders().bottomStyle()).isEqualTo(BorderStyle._double);
      assertThat(sheet.cell(2, 3).borders().bottomStyle()).isEqualTo(BorderStyle._double);

      for (int i = 0; i < sheet.rowCount(); i++) {

        StringBuilder sb = new StringBuilder();
        for (int j = 0; j < 16; j++) {
          sb.append(sheet.cell(i, j).asText()).append(' ');
        }
        sb.append('\n');
        System.out.println(sb);

      }

    }
  }

  @Test
  public void book1() {
    InputStream inputStream = getClass().getResourceAsStream("xlsx/book-1.xlsx");
    try (XlsxReader xlsxReader = new XlsxReader()) {
      xlsxReader.read(inputStream);

//      System.out.println("34jh2b54234 :: shared strings...");
//      xlsxReader.printSharedStrings(System.out);
//      System.out.println("34jh2b54234 :: shared strings end");

      assertThat(xlsxReader.sheetCount()).isEqualTo(1);

      Sheet sheet = xlsxReader.tabSelectedSheet();

      assertThat(sheet.rowCount()).isEqualTo(31);

      assertThat(sheet.row(7).cell(2).asText()).isEqualTo("B1");
      assertThat(sheet.cell(7, 2).asText()).isEqualTo("B1");
      assertThat(sheet.cell(8, 2).asText()).isEqualTo("B2");

//      System.out.println("wej1hbn43bn :: 8 2 topStyle    = " + sheet.cell(8, 2).borders().topStyle());
//      System.out.println("wej1hbn43bn :: 8 2 bottomStyle = " + sheet.cell(8, 2).borders().bottomStyle());
//      System.out.println("wej1hbn43bn :: 0 0 bottomStyle = " + sheet.cell(0, 0).borders().bottomStyle());

      assertThat(sheet.cell(8, 2).borders().topStyle()).isEqualTo(BorderStyle.dotted);
      assertThat(sheet.cell(8, 2).borders().bottomStyle()).isEqualTo(BorderStyle.dashDotDot);
      assertThat(sheet.cell(0, 0).borders().bottomStyle()).isEqualTo(BorderStyle.NONE);

//      System.out.println("wej1hbn43bn :: 4 9 hor  align = " + sheet.cell(4, 9).horAlign());
//      System.out.println("wej1hbn43bn :: 4 9 vert align = " + sheet.cell(4, 9).vertAlign());

      assertThat(sheet.cell(4, 9).horAlign()).isEqualTo(HorizontalAlign.center);
      assertThat(sheet.cell(4, 9).vertAlign()).isEqualTo(VerticalAlign.center);

      assertThat(sheet.cell(6, 7).horAlign()).isEqualTo(HorizontalAlign.right);
      assertThat(sheet.cell(6, 7).vertAlign()).isEqualTo(VerticalAlign.top);

      assertThat(sheet.cell(110, 110).horAlign()).isEqualTo(HorizontalAlign.left);
      assertThat(sheet.cell(110, 110).vertAlign()).isEqualTo(VerticalAlign.bottom);

      assertThat(sheet.cell(16, 8).asText()).isEqualTo("Note");
      assertThat(sheet.cell(16, 8).horAlign()).isEqualTo(HorizontalAlign.left);
      assertThat(sheet.cell(16, 8).vertAlign()).isEqualTo(VerticalAlign.bottom);
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

      assertThat(xlsxReader.tabSelectedSheet().name()).isEqualTo("Даты подполья");

      {
        Sheet sheet = xlsxReader.tabSelectedSheet();
        assertThat(sheet.cell(10, 10).asText()).isNull();
      }

      {
        Sheet sheet = xlsxReader.sheet(2);
        System.out.println("32gv4v :: cell 3 0 = " + sheet.cell(3, 0).asText() + " :: " + sheet.cell(3, 0).format());
        System.out.println("          cell 3 0 asDate = " + sheet.cell(3, 0).asDate());
      }

    }
  }

}
