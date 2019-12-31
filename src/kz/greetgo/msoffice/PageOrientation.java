package kz.greetgo.msoffice;

public enum PageOrientation {
  PORTRAIT(11906,16838),
  LANDSCAPE(16838, 11906);

  public final int width;
  public final int height;

  PageOrientation(int width, int height) {
    this.width = width;
    this.height = height;
  }

}
