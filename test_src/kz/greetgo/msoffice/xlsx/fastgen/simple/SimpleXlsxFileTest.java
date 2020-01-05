package kz.greetgo.msoffice.xlsx.fastgen.simple;

import org.testng.annotations.Test;

import java.io.FileOutputStream;

public class SimpleXlsxFileTest {
  @Test
  public void test1() throws Exception {
    SimpleFastXlsxFile x = new SimpleFastXlsxFile("build/tmp");

    x.newSheet(new double[]{23.3, 11.1, 21.12});

    x.appendRow(SimpleRowStyle.NORMAL, new String[]{"sad", "fdsfdsf", "fdsfds"});
    x.appendRow(SimpleRowStyle.GREEN, new String[]{"sad1", "fdsfdsf1", "fdsfds1"});
    x.appendRow(SimpleRowStyle.NORMAL, new String[]{"sad2", "fdsfdsf2", "fdsfds2"});

    x.newSheet(new double[]{11.3, 11.1, 21.12});

    x.appendRow(SimpleRowStyle.NORMAL, new String[]{"1sad", "2fdsfdsf", "3fdsfds"});
    x.appendRow(SimpleRowStyle.GREEN, new String[]{"1sad1", "2fdsfdsf1", "3fdsfds1"});
    x.appendRow(SimpleRowStyle.NORMAL, new String[]{"1sad2", "2fdsfdsf2", "3fdsfds2"});

    x.complete(new FileOutputStream("build/SimpleXlsxFileTest_test1.xlsx"));
  }

  @Test
  public void test2() throws Exception {
    SimpleFastXlsxFile x = new SimpleFastXlsxFile("build/tmp");

    long time1 = System.currentTimeMillis();

    for (int sheet = 1; sheet <= 3; sheet++) {
      x.newSheet(new double[]{23.3, 11.1, 21.12, 23.3, 11.1, 21.12, 23.3, 11.1, 21.12, 23.3,
        11.1, 21.12, 23.3, 11.1, 21.12, 23.3, 11.1, 21.12, 23.3, 11.1, 21.12, 23.3, 11.1, 21.12,});

      for (int i = 0; i < 50; i++) {
        x.appendRow(SimpleRowStyle.NORMAL, new String[]{sheet + "sad" + i, sheet + "fdsfdsf" + i,
          sheet + "fdsfds" + i, sheet + "1sad" + i, sheet + "7fdsfdsf" + i,
          sheet + "1fdsfds" + i, sheet + "2sad" + i, sheet + "6fdsfdsf" + i,
          sheet + "2fdsfds" + i, sheet + "3sad" + i, sheet + "5fdsfdsf" + i,
          sheet + "3fdsfds" + i, sheet + "4sad" + i, sheet + "4fdsfdsf" + i,
          sheet + "4fdsfds" + i, sheet + "5sad" + i, sheet + "3fdsfdsf" + i,
          sheet + "5fdsfds" + i, sheet + "6sad" + i, sheet + "2fdsfdsf" + i,
          sheet + "6fdsfds" + i, sheet + "7sad" + i, sheet + "1fdsfdsf" + i,
          sheet + "7fdsfds" + i,});
        x.appendRow(SimpleRowStyle.GREEN, new String[]{"sad1d" + i, "fdsfdsf1x" + i,
          "fdsfds1x" + i, "1sad1d" + i, "7fdsfdsf1x" + i, "1fdsfds1x" + i, "2sad1d" + i,
          "6fdsfdsf1x" + i, "2fdsfds1x" + i, "3sad1d" + i, "5fdsfdsf1x" + i, "3fdsfds1x" + i,
          "4sad1d" + i, "4fdsfdsf1x" + i, "4fdsfds1x" + i, "5sad1d" + i, "3fdsfdsf1x" + i,
          "5fdsfds1x" + i, "6sad1d" + i, "2fdsfdsf1x" + i, "6fdsfds1x" + i, "7sad1d" + i,
          "1fdsfdsf1x" + i, "7fdsfds1x" + i,});
      }

    }

    long time2 = System.currentTimeMillis();

    System.out.println("data come " + (time2 - time1) / 1000.0 + " s");

    x.complete(new FileOutputStream("build/SimpleXlsxFileTest_test2.xlsx"));

    long time3 = System.currentTimeMillis();
    System.out.println("copy files " + (time3 - time2) / 1000.0 + " s");
    System.out.println("total time " + (time3 - time1) / 1000.0 + " s");
  }
}
