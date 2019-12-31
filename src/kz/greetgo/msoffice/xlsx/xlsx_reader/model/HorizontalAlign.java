package kz.greetgo.msoffice.xlsx.xlsx_reader.model;

public enum HorizontalAlign {
  left, center, right;

  public static HorizontalAlign valueOrNull(String str) {
    if (str == null) return null;
    for (HorizontalAlign value : values()) {
      if (value.name().equals(str)) return value;
    }
    throw new IllegalArgumentException("Unknown name `" + str + "` of " + HorizontalAlign.class.getSimpleName());
  }

}
