package kz.greetgo.msoffice.docx;

public enum VertAlign {
  TOP("top"), CENTER("center"), BOTTOM("bottom");

  private final String code;

  VertAlign(String code) {
    this.code = code;
  }

  public String getCode() {
    return code;
  }
}
