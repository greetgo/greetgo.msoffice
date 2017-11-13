package kz.greetgo.msoffice.xlsx.gen;

public enum Align {
  
  left("l", "low"), //
  top("t", "low"), //
  right("r", "nextTo"), //
  bottom("b", "low"), //
  center("c", "low");
  
  private String tag;
  private String tagOY;
  
  private Align(String tag, String tagOY) {
    this.tag = tag;
    this.tagOY = tagOY;
  }
  
  /** Значение в xml для выравнивания */
  String getTag() {
    return tag;
  }
  
  /** Значение в xml для выравнивания подписей по оси OY */
  String getTagOY() {
    return tagOY;
  }
  
  public String str() {
    return name();
  }
}