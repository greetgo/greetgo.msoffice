package kz.greetgo.msoffice.docx;

public enum Underline {
  NONE(""), WORDS("words"), SINGLE("single"), DOUBLE("double"), THICK("thick"), DOTTED("dotted"),
  DOTTED_HEAVY("dottedHeavy"), DASH("dash");
  
  private final String code;
  
  private Underline(String code) {
    this.code = code;
  }
  
  public String getCode() {
    return code;
  }
}
