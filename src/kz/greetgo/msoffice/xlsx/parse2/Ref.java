package kz.greetgo.msoffice.xlsx.parse2;

import java.util.Objects;

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

  private static final int SIZEOF_LONG = 8;
  private static final int SIZEOF_INT = 4;
  public static final int SIZE = SIZEOF_LONG + SIZEOF_INT;

  @SuppressWarnings("PointlessArithmeticExpression")
  public static Ref readFrom(byte[] buffer, int offset) {

    //@formatter:off
    long refOffset = ((((long) buffer[ offset +  0 ]) & 0xFFL) << (8 * 0))
                   | ((((long) buffer[ offset +  1 ]) & 0xFFL) << (8 * 1))
                   | ((((long) buffer[ offset +  2 ]) & 0xFFL) << (8 * 2))
                   | ((((long) buffer[ offset +  3 ]) & 0xFFL) << (8 * 3))
                   | ((((long) buffer[ offset +  4 ]) & 0xFFL) << (8 * 4))
                   | ((((long) buffer[ offset +  5 ]) & 0xFFL) << (8 * 5))
                   | ((((long) buffer[ offset +  6 ]) & 0xFFL) << (8 * 6))
                   | ((((long) buffer[ offset +  7 ]) & 0xFFL) << (8 * 7))
    ;

    int refLength =  ((((int) buffer[ offset +  8 ]) & 0xFF) << (8 * 0))
                   | ((((int) buffer[ offset +  9 ]) & 0xFF) << (8 * 1))
                   | ((((int) buffer[ offset + 10 ]) & 0xFF) << (8 * 2))
                   | ((((int) buffer[ offset + 11 ]) & 0xFF) << (8 * 3))
    ;
    //@formatter:on

    return Ref.of(refOffset, refLength);
  }

  @SuppressWarnings("PointlessArithmeticExpression")
  public void writeTo(byte[] buffer, int offset) {
    //@formatter:off
    buffer[offset +  0] = (byte) ((this.offset & 0x00000000000000FFL) >> (8 * 0));
    buffer[offset +  1] = (byte) ((this.offset & 0x000000000000FF00L) >> (8 * 1));
    buffer[offset +  2] = (byte) ((this.offset & 0x0000000000FF0000L) >> (8 * 2));
    buffer[offset +  3] = (byte) ((this.offset & 0x00000000FF000000L) >> (8 * 3));
    buffer[offset +  4] = (byte) ((this.offset & 0x000000FF00000000L) >> (8 * 4));
    buffer[offset +  5] = (byte) ((this.offset & 0x0000FF0000000000L) >> (8 * 5));
    buffer[offset +  6] = (byte) ((this.offset & 0x00FF000000000000L) >> (8 * 6));
    buffer[offset +  7] = (byte) ((this.offset & 0xFF00000000000000L) >> (8 * 7));

    buffer[offset +  8] = (byte) ((this.length & 0x000000FF) >> (8 * 0));
    buffer[offset +  9] = (byte) ((this.length & 0x0000FF00) >> (8 * 1));
    buffer[offset + 10] = (byte) ((this.length & 0x00FF0000) >> (8 * 2));
    buffer[offset + 11] = (byte) ((this.length & 0xFF000000) >> (8 * 3));
    //@formatter:on
  }

}
