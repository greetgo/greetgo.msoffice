package kz.greetgo.msoffice.xlsx.xlsx_reader.model;

import java.math.BigDecimal;

public class ColumnInfo {
  public int minIndex, maxIndex;
  public BigDecimal width;

  @Override
  public String toString() {
    return "ColumnInfo{" + "minIndex=" + minIndex +
      ", maxIndex=" + maxIndex +
      ", width=" + width +
      '}';
  }
}
