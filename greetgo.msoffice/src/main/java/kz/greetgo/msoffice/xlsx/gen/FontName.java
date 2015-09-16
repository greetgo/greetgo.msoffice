package kz.greetgo.msoffice.xlsx.gen;

public enum FontName {
  
  Calibri("2", "204", "minor"), //
  Times_New_Roman("1", "204", null), //
  Arial("2", "204", null), //
  Arial_Unicode_MS("2", "204", null);
  
  private final String family;
  private final String charset;
  private final String scheme;
  
  private FontName(String family, String charset, String scheme) {
    this.family = family;
    this.charset = charset;
    this.scheme = scheme;
  }
  
  public String str() {
    return name().replace('_', ' ');
  }
  
  public String getFamily() {
    return family;
  }
  
  public String getCharset() {
    return charset;
  }
  
  public String getScheme() {
    return scheme;
  }
}