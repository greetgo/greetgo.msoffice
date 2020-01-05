package kz.greetgo.msoffice;

import kz.greetgo.msoffice.util.UtilOffice;
import kz.greetgo.msoffice.xlsx.reader.model.ColData;
import kz.greetgo.msoffice.xlsx.reader.model.MergeCell;
import kz.greetgo.msoffice.xlsx.reader.model.RowData;
import kz.greetgo.msoffice.xlsx.reader.model.ValueType;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Random;

import static org.fest.assertions.api.Assertions.assertThat;

public class UtilTest {

  private final Map<String, String> map = new HashMap<>();

  public static ColData rndColData(Random rnd) {
    ColData colData = new ColData();
    colData.valueType = ValueType.values()[rnd.nextInt(ValueType.values().length)];
    colData.value = UtilTest.rndStr(rnd, 30 + rnd.nextInt(30));
    colData.col = rnd.nextInt();
    return colData;
  }

  public static BigDecimal rndBd(Random rnd) {
    byte[] intBytes = new byte[30 + rnd.nextInt(30)];
    rnd.nextBytes(intBytes);

    BigInteger bi = new BigInteger(intBytes);
    int scale = rnd.nextInt();
    return new BigDecimal(bi, scale);
  }

  public static RowData rndRowData(Random rnd) {
    return rndRowData(rnd, null);
  }

  public static RowData rndRowData(Random rnd, Integer index) {
    RowData rowData = new RowData();
    rowData.index = index == null ? rnd.nextInt() : index;
    rowData.height = rndBd(rnd);

    for (int i = 0, c = 10 + rnd.nextInt(10); i < c; i++) {
      rowData.cols.add(rndColData(rnd));
    }
    return rowData;
  }

  public static MergeCell rndMergeCell(Random rnd) {
    MergeCell ret = new MergeCell();
    ret.rowFrom = rnd.nextInt();
    ret.colFrom = rnd.nextInt();
    ret.colFrom = rnd.nextInt();
    ret.colTo = rnd.nextInt();
    return ret;
  }

  @BeforeMethod
  public void prepareExcelMap() {
    map.clear();

    map.put("2011-11-11T11:11:11Z", "40858.466099537036");
    map.put("2011-11-11T17:55:07Z", "40858.7466099537036");
  }

  @Test
  public void excelToDate() {
    for (Entry<String, String> e : map.entrySet()) {
      Date d = UtilOffice.excelToDate(e.getValue());
      assertThat(e.getKey())
        .describedAs("Для excel-даты [" + e.getValue() + "]")
        .isEqualTo(UtilOffice.toW3CDTF(d));
    }
  }

  @Test
  public void toExcelDateTime() {
    for (Entry<String, String> e : map.entrySet()) {
      String sActual = UtilOffice.toExcelDateTime(UtilOffice.parseW3CDTF(e.getKey()));
      String sExpected = e.getValue();
      assertAbout("Для java-даты [" + e.getKey() + "]", 8e-6, sExpected, sActual);
    }
  }

  @SuppressWarnings("SameParameterValue")
  private static void assertAbout(String message, double eps, String expected, String actual) {
    double delta = new BigDecimal(expected).subtract(new BigDecimal(actual)).abs().doubleValue();
    assertThat(delta <= eps)
      .describedAs(message + ": expected = " + expected + ", actual = " + actual + ", delta = " + delta
        + ", but eps = " + eps)
      .isTrue();
  }

  @SuppressWarnings("SpellCheckingInspection")
  private static final String ENG = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  private static final String DEG = "0123456789";
  private static final String RUS = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ";
  private static final String CHIN = "⻨⻩⻪⻫⻬⻭⻮⻯⻰⻱⻲⻳⻠⻡⻢⻣⻤⻥⻦⻧⾀⾁⾂⾃⾄⾅⾆⾇⾈⾉⾊⾋⾌⾍⾎⾏⾐⾑⾒⾓⾔⾕⾖⾗⾘" +
    "⾙⾚⾛⾜⾝⾞⾟⾠⾡⾢⾣⾤⾥⾦⾧⾨⾩⾪⾫⾬⾭⾮⾯⾰⾱⾲⾳⾴⾵⾶⾷⾸⾹⾺⾻⾼⾽⾾⾿⿀⿁⿂⿃⿄⿅⿆⿇⿈⿉⿊⿋⿌⿍⿎⿏⿐⿑⿒⿓⿔⿕";

  private static final char[] ALL = (ENG.toUpperCase() + ENG.toLowerCase() + DEG + RUS.toLowerCase()
    + RUS.toUpperCase() + CHIN).toCharArray();

  public static String rndStr(Random rnd, int length) {
    char[] chars = new char[length];
    for (int i = 0; i < length; i++) {
      chars[i] = ALL[rnd.nextInt(ALL.length)];
    }
    return new String(chars);
  }
}
