package kz.greetgo.msoffice.docx;

public enum Align {
  LEFT("left"), CENTER("center"), RIGHT("right"), BOTH("both");
  
  private final String code;
  
  private Align(String code) {
    this.code = code;
  }
  
  public String getCode() {
    return code;
  }
}
