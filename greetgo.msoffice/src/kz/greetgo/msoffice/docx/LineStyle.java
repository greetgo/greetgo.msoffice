package kz.greetgo.msoffice.docx;

public enum LineStyle {
  NONE(null), SINGLE("single"), DOUBLE("double"), TRIPLE("triple"), DOTTED("dotted"), DOT_DASH(
      "dotDash"), DASHED("dashed"), DOT_DOT_DASH("dotDotDash"), DASH_SMALL_GAP("dashSmallGap"),
  WAVE("wave"), DOUBLE_WAVE("doubleWave");
  
  private final String code;
  
  private LineStyle(String code) {
    this.code = code;
  }
  
  public String getCode() {
    return code;
  }
}
