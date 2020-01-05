package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.util.BinUtil;

import java.util.Objects;

import static kz.greetgo.msoffice.util.BinUtil.readInt;
import static kz.greetgo.msoffice.util.BinUtil.readLong;
import static kz.greetgo.msoffice.util.BinUtil.writeInt;
import static kz.greetgo.msoffice.util.BinUtil.writeLong;

public class Ref {
  public final long offset;
  public final int length;

  private Ref(long offset, int length) {
    this.offset = offset;
    this.length = length;
  }

  @Override
  public String toString() {
    return "Ref{" + "offset=" + offset + ", length=" + length + '}';
  }


  @Override
  public boolean equals(Object o) {
    if (this == o) return true;
    if (o == null || getClass() != o.getClass()) return false;
    Ref ref = (Ref) o;
    return offset == ref.offset && length == ref.length;
  }

  @Override
  public int hashCode() {
    return Objects.hash(offset, length);
  }

  public static Ref of(long offset, int length) {
    return new Ref(offset, length);
  }

  public static final int SIZE = BinUtil.SIZEOF_LONG + BinUtil.SIZEOF_INT;

  public static Ref readFrom(byte[] buffer, int offset) {
    long refOffset = readLong(buffer, offset);
    int refLength = readInt(buffer, offset + BinUtil.SIZEOF_LONG);
    return Ref.of(refOffset, refLength);
  }

  public void writeTo(byte[] buffer, int offset) {
    writeLong(this.offset, buffer, offset);
    writeInt(this.length, buffer, offset + BinUtil.SIZEOF_LONG);
  }

}
