package kz.greetgo.msoffice.xlsx.fastgen.simple_inline;

import java.io.FileOutputStream;

import kz.greetgo.msoffice.xlsx.fastgen.simple.SimpleRowStyle;

import org.testng.annotations.Test;

public class SimpleInlineFastXlsxFileTest {
  @Test
  public void test1() throws Exception {
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
    for (int i = 0; i < 900000; i++) {
      x.appendRow(SimpleRowStyle.GREEN, s);
      if (i % 10000 == 0) System.out.println("i = " + i);
    }
    
    System.out.println("save...");
    
    x.newSheet(new double[] { s[0].length(), s[0].length(), s[0].length(), s[0].length(),
        s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(),
        s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(),
        s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length(),
        s[0].length(), s[0].length(), s[0].length(), s[0].length(), s[0].length() });
    for (int i = 0; i < 500000; i++) {
      x.appendRow(SimpleRowStyle.GREEN, s);
      if (i % 10000 == 0) System.out.println("i = " + i);
    }
    
    System.out.println("save...");
    x.complete(new FileOutputStream("build/dsa3.xlsx"));
    System.out.println("ok");
  }
}
