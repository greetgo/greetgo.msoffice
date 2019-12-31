package kz.greetgo.msoffice.xlsx.xlsx_reader.model;

public class Border4 {
  public final Border top = new Border();
  public final Border left = new Border();
  public final Border right = new Border();
  public final Border bottom = new Border();
  public final Border diagonal = new Border();

  public Border get(BorderSide borderSide) {
    switch (borderSide) {
      //@formatter:off
      case top      : return top      ;
      case left     : return left     ;
      case right    : return right    ;
      case bottom   : return bottom   ;
      case diagonal : return diagonal ;
      //@formatter:on
      default:
        throw new IllegalArgumentException("Unknown borderSide = " + borderSide);
    }
  }
}
