package kz.greetgo.msoffice.docx;

public enum TextVertAlign {
  NORMAL(""), SUPERSCRIPT("superscript"), SUBSCRIPT("subscript");
  
  private final String code;
  
  private TextVertAlign(String code) {
    this.code = code;
  }
  
  public String getCode() {
    return code;
  }
}
