package kz.greetgo.msoffice.xlsx.reader.model;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

public enum BorderStyle {
  hair, dotted, dashDotDot, dashDot, dashed, thin, mediumDashDotDot,
  slantDashDot, mediumDashDot, mediumDashed, medium, thick, _double,

  NONE;

  private static final Map<String, BorderStyle> map;

  static {
    Map<String, BorderStyle> map1 = new HashMap<>();
    for (BorderStyle value : values()) {
      String name = value.name();
      if (name.startsWith("_")) {
        name = name.substring(1);
      }
      map1.put(name, value);
    }
    map = Collections.unmodifiableMap(map1);
  }

  public static Optional<BorderStyle> from(String str) {
    if (str == null) return Optional.empty();
    return Optional.ofNullable(map.get(str));
  }

}
