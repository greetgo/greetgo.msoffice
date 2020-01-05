package kz.greetgo.msoffice.xlsx.xlsx_reader.model;

import kz.greetgo.msoffice.util.UtilOffice;

public class MergeCell {

  /**
   * from zero
   */
  public int rowFrom;

  /**
   * from zero
   */
  public int rowTo;

  /**
   * from zero
   */
  public int colFrom;

  /**
   * from zero
   */
  public int colTo;

  public void parseAndSet(String ref) {
    String[] split = ref.split(":");
    int[] from = UtilOffice.parseCellCoordinate(split[0]);
    int[] to = UtilOffice.parseCellCoordinate(split[1]);

    colFrom = from[0] - 1;
    rowFrom = from[1] - 1;
    colTo = to[0] - 1;
    rowTo = to[1] - 1;
  }

  @Override
  public String toString() {
    return "MergeCell{" + "rowFrom=" + rowFrom + ", rowTo=" + rowTo + ", colFrom=" + colFrom + ", colTo=" + colTo + '}';
  }
}
