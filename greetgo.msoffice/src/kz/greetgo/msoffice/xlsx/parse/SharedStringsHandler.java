package kz.greetgo.msoffice.xlsx.parse;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

public class SharedStringsHandler extends DefaultHandler {
  private final Connection connection;
  private long nom = 0;
  
  public SharedStringsHandler(Connection connection) {
    this.connection = connection;
  }
  
  private final XmlIn in = new XmlIn();
  
  PreparedStatement strInsertPS = null;
  private int currentBachSize = 0;
  private static final int MAX_BATCH_SIZE = 2000;
  
  @Override
  public void startDocument() throws SAXException {
    try {
      strInsertPS = connection.prepareStatement("insert into strs (nom, value) values (?, ?)");
      currentBachSize = 0;
    } catch (SQLException e) {
      throw new RuntimeException(e);
    }
  }
  
  @Override
  public void endDocument() throws SAXException {
    try {
      if (currentBachSize > 0) {
        commit();
      }
      strInsertPS.close();
      strInsertPS = null;
    } catch (SQLException e) {
      throw new RuntimeException(e);
    }
  }
  
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
    strInsertPS.setLong(1, nom++);
    strInsertPS.setString(2, string);
    strInsertPS.addBatch();
    currentBachSize++;
    if (currentBachSize >= MAX_BATCH_SIZE) {
      commit();
    }
  }
  
  private void commit() throws SQLException {
    strInsertPS.executeBatch();
    connection.commit();
    currentBachSize = 0;
  }
}
