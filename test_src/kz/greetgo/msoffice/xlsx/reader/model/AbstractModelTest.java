package kz.greetgo.msoffice.xlsx.reader.model;

import java.util.Random;

import static org.fest.assertions.api.Assertions.assertThat;

public abstract class AbstractModelTest {

  protected Random rnd = new Random();

  public static void assertRowDataEquals(RowData actual, RowData expected, String d) {
    assertThat(actual).describedAs(d).isNotNull();
    assertThat(actual.height).describedAs(d).isEqualTo(expected.height);
    assertThat(actual.cols).describedAs(d).hasSameSizeAs(expected.cols);
    for (int i = 0, s = actual.cols.size(); i < s; i++) {

      ColData actualCol = actual.cols.get(i);
      ColData expectedCol = expected.cols.get(i);

      String d2 = d + ", col index = " + i;
      assertThat(actualCol).describedAs(d2).isNotNull();
      assertThat(actualCol.valueType).describedAs(d2).isEqualTo(expectedCol.valueType);
      assertThat(actualCol.col).describedAs(d2).isEqualTo(expectedCol.col);
      assertThat(actualCol.value).describedAs(d2).isEqualTo(expectedCol.value);

    }
  }

  public static void assertMergeCellEquals(MergeCell actual, MergeCell expected, String d) {
    assertThat(actual).describedAs(d).isNotNull();
    assertThat(actual.colFrom).describedAs(d).isEqualTo(expected.colFrom);
    assertThat(actual.colTo).describedAs(d).isEqualTo(expected.colTo);
    assertThat(actual.rowFrom).describedAs(d).isEqualTo(expected.rowFrom);
    assertThat(actual.rowTo).describedAs(d).isEqualTo(expected.rowTo);
  }
}
