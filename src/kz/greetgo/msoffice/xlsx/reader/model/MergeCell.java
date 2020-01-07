package kz.greetgo.msoffice.xlsx.reader.model;

import kz.greetgo.msoffice.util.BinUtil;
import kz.greetgo.msoffice.util.UtilOffice;

import static kz.greetgo.msoffice.util.BinUtil.SIZEOF_INT;
import static kz.greetgo.msoffice.util.BinUtil.readInt;
import static kz.greetgo.msoffice.util.BinUtil.writeInt;

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
    return "MergeCell{rows=" + rowFrom + ".." + rowTo + ", cols=" + colFrom + ".." + colTo + '}';
  }

  public void writeBytes(byte[] buffer, int offset) {
    writeInt(rowFrom, buffer, offset);
    writeInt(rowTo, buffer, offset + SIZEOF_INT);
    writeInt(colFrom, buffer, offset + SIZEOF_INT * 2);
    writeInt(colTo, buffer, offset + SIZEOF_INT * 3);
  }

  public static MergeCell readFrom(byte[] buffer, int offset) {
    MergeCell ret = new MergeCell();
    ret.rowFrom = readInt(buffer, offset);
    ret.rowTo = readInt(buffer, offset + SIZEOF_INT);
    ret.colFrom = readInt(buffer, offset + SIZEOF_INT * 2);
    ret.colTo = readInt(buffer, offset + SIZEOF_INT * 3);
    return ret;
  }

  public int byteSize() {
    return BinUtil.SIZEOF_INT * 4;
  }
}
