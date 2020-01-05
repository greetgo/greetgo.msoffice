package kz.greetgo.msoffice.util;

import org.xml.sax.Attributes;

public class XmlUtil {
  public static String toStr(Attributes attributes) {
    StringBuilder sb = new StringBuilder();

    for (int i = 0; i < attributes.getLength(); i++) {
      String localName = attributes.getLocalName(i);
      String value = attributes.getValue(i);
      if (i > 0) sb.append(' ');
      sb.append(localName).append('=').append(value);
    }

    return sb.toString();
  }
}
