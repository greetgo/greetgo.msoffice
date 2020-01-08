package kz.greetgo.msoffice.xlsx.reader.model;

public enum ValueType {
  UNKNOWN, STR, INLINE_STR;

  public static ValueType parse(String t) {
    if (t == null) return UNKNOWN;
    t = t.toLowerCase();
    if ("s".equals(t)) return STR;
    if ("inlineStr".toLowerCase().equals(t)) return INLINE_STR;
    return UNKNOWN;
  }
}
