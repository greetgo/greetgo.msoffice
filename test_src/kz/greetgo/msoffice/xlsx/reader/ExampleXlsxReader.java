package kz.greetgo.msoffice.xlsx.reader;

import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;

public class ExampleXlsxReader {
  public static void main(String[] args) throws Exception {

    Path fileToScan = Paths.get("/path/to/xlsx-file.xlsx");

    InputStream inputStream = new FileInputStream(fileToScan.toFile());

    try (XlsxReader xlsxReader = new XlsxReader()) {

      xlsxReader.read(inputStream);

      // inputStream here already is closed

      Sheet tabSelectedSheet = xlsxReader.tabSelectedSheet();

      int sheetCount = xlsxReader.sheetCount();

      Sheet lastSheet = xlsxReader.sheet(sheetCount - 1);

      String textOfCell1_3 = lastSheet.cell(1, 3).asText();

      // etc

    }

  }
}
