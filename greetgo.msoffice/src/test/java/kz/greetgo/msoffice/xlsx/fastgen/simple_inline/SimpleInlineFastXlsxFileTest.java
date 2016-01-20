package kz.greetgo.msoffice.xlsx.fastgen.simple_inline;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Random;

import org.testng.annotations.Test;

import kz.greetgo.msoffice.xlsx.fastgen.simple.SimpleRowStyle;
import kz.greetgo.msoffice.xlsx.parse.Cell;
import kz.greetgo.msoffice.xlsx.parse.RowHandler;
import kz.greetgo.msoffice.xlsx.parse.Sheet;
import kz.greetgo.msoffice.xlsx.parse.XlsxParser;

public class SimpleInlineFastXlsxFileTest {
  @Test
  public void createAndParseBigFile() throws Exception {
    final int count = 9000, page = 1000;
    
    SimpleInlineFastXlsxFile x = new SimpleInlineFastXlsxFile("build/tmp");
    
    String[] s = new String[27];
    for (int j = 0; j < 27; j++) {
      s[j] = "14352435zdsfgsxdfhyujsxdru65zhderfgh";
    }
    
    x.newSheet(new double[] { s[0].length(), s[0].length(), s[0].length(), s[0].length(),
        s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(),
        s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(),
        s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(),
        s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length() });
    for (int i = 0; i < count; i++) {
      x.appendRow(SimpleRowStyle.GREEN, s);
      if (i % page == 0) System.out.println("i = " + i);
    }
    
    System.out.println("sheet 1 ok");
    
    x.newSheet(new double[] { s[0].length(), s[0].length(), s[0].length(), s[0].length(),
        s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(),
        s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(),
        s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(),
        s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length() });
    for (int i = 0, C = count / 2; i < C; i++) {
      x.appendRow(SimpleRowStyle.GREEN, s);
      if (i % page == 0) System.out.println("i = " + i);
    }
    
    System.out.println("sheet 2 ok");
    
    String filename = "build/dsa" + new Random().nextLong() + ".xlsx";
    
    System.out.println("save to " + filename + "...");
    
    x.complete(new FileOutputStream(filename));
    
    System.out.println("File created ok: " + filename);
    System.out.println("Parsing...");
    
    XlsxParser parser = new XlsxParser();
    try {
      System.out.println("Loading...");
      parser.load(new FileInputStream(filename));
      
      System.out.println("Loaded ok. Scanning...");
      
      Sheet sheet = parser.sheets().get(0);
      sheet.scanRows(30, new RowHandler() {
        @Override
        public void handle(List<Cell> row, int rowIndex) throws Exception {
          if (rowIndex % page == 0) {
            System.out.println("Scanned row " + rowIndex);
            System.out.println(row);
          }
        }
      });
    } finally {
      parser.close();
    }
    
    System.out.println("Finish");
  }
}
