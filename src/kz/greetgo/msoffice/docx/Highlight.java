package kz.greetgo.msoffice.docx;

public enum Highlight {
  NONE(""), YELLOW("yellow"), GREEN("green"), CYAN("cyan"), MAGENTA("magenta"), BLUE("blue"), RED(
      "red"), DARK_BLUE("darkBlue"), DARK_CYAN("darkCyan"), DARK_GREEN("darkGreen"), DARK_MAGENTA(
      "darkMagenta"), DARK_RED("darkRed"), DARK_YELLOW("darkYellow"), DARK_GRAY("darkGray"),
  LIGHT_GRAY("lightGray"), BLACK("black");
  
  private final String code;
  
  private Highlight() {
    code = null;
  }
  
  private Highlight(String code) {
    this.code = code;
  }
  
  public String getCode() {
    return code;
  }
}
