package kz.greetgo.msoffice.xlsx.xlsx_reader.model;

import org.testng.annotations.Test;

import static kz.greetgo.msoffice.UtilTest.rndRowData;

public class RowDataTest extends AbstractModelTest {

  @Test
  public void writeTo__readFrom() {

    RowData rowData = rndRowData(rnd);

    byte[] buffer = new byte[5000];
    rnd.nextBytes(buffer);

    //
    //
    rowData.writeTo(buffer, 10);
    RowData actual = RowData.readFrom(buffer, 10);
    //
    //

    assertEquals(actual, rowData, "");
  }

  @Test
  public void byteSize() {

    RowData rowData = rndRowData(rnd);

    //
    //
    int size = rowData.byteSize();
    //
    //

    byte[] buffer = new byte[size];
    rnd.nextBytes(buffer);

    rowData.writeTo(buffer, 0);
    RowData actual = RowData.readFrom(buffer, 0);

    assertEquals(actual, rowData, "");
  }

}
