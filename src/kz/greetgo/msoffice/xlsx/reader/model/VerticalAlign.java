package kz.greetgo.msoffice.xlsx.reader.model;

public enum VerticalAlign {
  top, center, bottom;

  public static VerticalAlign valueOrNull(String str) {
    if (str == null) return null;
    for (VerticalAlign value : values()) {
      if (value.name().equals(str)) return value;
    }
    throw new IllegalArgumentException("Unknown name `" + str + "` of " + VerticalAlign.class.getSimpleName());
  }
}
