package kz.greetgo.msoffice.xlsx.reader.model;

import java.util.Objects;

public enum BorderSide {
  top {
    @Override
    public int[] move(int[] c) {
      assertCoordinates(c);
      return new int[]{c[0] - 1, c[1]};
    }

    @Override
    public BorderSide mirror() {
      return bottom;
    }
  },

  bottom {
    @Override
    public int[] move(int[] c) {
      assertCoordinates(c);
      return new int[]{c[0] + 1, c[1]};
    }

    @Override
    public BorderSide mirror() {
      return top;
    }
  },

  left {
    @Override
    public int[] move(int[] c) {
      assertCoordinates(c);
      return new int[]{c[0], c[1] - 1};
    }

    @Override
    public BorderSide mirror() {
      return right;
    }
  },

  right {
    @Override
    public int[] move(int[] c) {
      assertCoordinates(c);
      return new int[]{c[0], c[1] + 1};
    }

    @Override
    public BorderSide mirror() {
      return left;
    }
  },

  diagonal {
    @Override
    public int[] move(int[] c) {
      assertCoordinates(c);
      return c;
    }

    @Override
    public BorderSide mirror() {
      return diagonal;
    }
  };

  private static void assertCoordinates(int[] coordinates) {
    Objects.requireNonNull(coordinates);
    if (coordinates.length != 2) {
      throw new IllegalArgumentException("Illegal coordinates length");
    }
  }

  public abstract int[] move(int[] coordinates);

  public abstract BorderSide mirror();
}
