package kz.greetgo.msoffice.xlsx.reader.model;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import static kz.greetgo.msoffice.util.BinUtil.SIZEOF_INT;
import static kz.greetgo.msoffice.util.BinUtil.readBd;
import static kz.greetgo.msoffice.util.BinUtil.readInt;
import static kz.greetgo.msoffice.util.BinUtil.sizeOfBd;
import static kz.greetgo.msoffice.util.BinUtil.writeBd;
import static kz.greetgo.msoffice.util.BinUtil.writeInt;

public class RowData {
  /**
   * index of row from zero
   */
  public BigDecimal height;
  public final List<ColData> cols = new ArrayList<>();

  public double heightAsDouble() {
    return height == null ? 15 : height.doubleValue();
  }

  private static final RowData EMPTY = new RowData();

  public static RowData empty() {
    return EMPTY;
  }

  public void writeTo(byte[] buffer, int offset) {
    int i = offset;

    writeBd(height, buffer, i);
    i += sizeOfBd(height);

    writeInt(cols.size(), buffer, i);
    i += SIZEOF_INT;

    for (ColData col : cols) {
      col.writeBytes(buffer, i);
      i += col.byteSize();
    }
  }

  public static RowData readFrom(byte[] buffer, int offset) {

    int i = offset;

    RowData ret = new RowData();

    ret.height = readBd(buffer, i);
    i += sizeOfBd(ret.height);

    int size = readInt(buffer, i);
    i += SIZEOF_INT;

    for (int j = 0; j < size; j++) {
      ColData colData = ColData.readBytes(buffer, i);
      i += colData.byteSize();
      ret.cols.add(colData);
    }

    return ret;
  }

  public int byteSize() {
    int ret = sizeOfBd(height) + SIZEOF_INT;
    for (ColData col : cols) {
      ret += col.byteSize();
    }
    return ret;
  }

}
