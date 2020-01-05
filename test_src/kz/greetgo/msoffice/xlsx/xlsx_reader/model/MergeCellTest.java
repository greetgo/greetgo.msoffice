package kz.greetgo.msoffice.xlsx.xlsx_reader.model;

import org.testng.annotations.Test;

import static kz.greetgo.msoffice.UtilTest.rndMergeCell;

public class MergeCellTest extends AbstractModelTest {

  @Test
  public void writeBytes__readFrom() {

    MergeCell mergeCell = rndMergeCell(rnd);

    byte[] buffer = new byte[100];

    //
    //
    mergeCell.writeBytes(buffer, 10);
    MergeCell actual = MergeCell.readFrom(buffer, 10);
    //
    //

    assertMergeCellEquals(actual, mergeCell, "");

  }


  @Test
  public void byteSize() {

    MergeCell mergeCell = rndMergeCell(rnd);

    //
    //
    int byteSize = mergeCell.byteSize();
    //
    //

    byte[] buffer = new byte[byteSize];

    mergeCell.writeBytes(buffer, 0);
    MergeCell actual = MergeCell.readFrom(buffer, 0);

    assertMergeCellEquals(actual, mergeCell, "");

  }


}
