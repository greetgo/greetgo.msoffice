package kz.greetgo.msoffice.xlsx.parse;

import java.sql.SQLException;
import java.util.Map;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

public class SharedStringsHandlerMemory extends DefaultHandler {
  private final Map<Long, String> values;
  private long nom = 0;
  
  public SharedStringsHandlerMemory(Map<Long, String> values) {
    this.values = values;
  }
  
  private final XmlIn in = new XmlIn();
  
  @Override
  public void startDocument() throws SAXException {}
  
  @Override
  public void endDocument() throws SAXException {}
  
  @Override
  public void startElement(String uri, String localName, String qName, Attributes attributes)
      throws SAXException {
    textBuilder = null;
    in.stepIn(qName);
  }
  
  @Override
  public void endElement(String uri, String localName, String qName) throws SAXException {
    if ("sst/si/t".equals(in.current())) {
      appendString(text());
    }
    in.stepOut();
  }
  
  private StringBuilder textBuilder = null;
  
  private String text() {
    if (textBuilder == null) return "";
    return textBuilder.toString();
  }
  
  @Override
  public void characters(char[] ch, int start, int length) throws SAXException {
    if (textBuilder == null) textBuilder = new StringBuilder();
    textBuilder.append(ch, start, length);
  }
  
  private void appendString(String string) {
    try {
      appendStringEx(string);
    } catch (SQLException e) {
      throw new RuntimeException(e);
    }
  }
  
  private void appendStringEx(String string) throws SQLException {
    values.put(nom++, string);
  }
}
