package kz.greetgo.msoffice.xlsx.xlsx_reader;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.List;

public abstract class AbstractXmlHandler extends DefaultHandler {

  private final List<String> tags = new ArrayList<>();

  @Override
  public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
    tags.add(localName);
    characters.setLength(0);
    try {
      startTag(tagPath(), attributes);
    } catch (SAXException | RuntimeException e) {
      throw e;
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }

  protected String tagPath() {
    return "/" + String.join("/", tags);
  }

  protected abstract void startTag(String tagPath, Attributes attributes) throws Exception;

  private final StringBuilder characters = new StringBuilder();

  @Override
  public void characters(char[] ch, int start, int length) {
    characters.append(ch, start, length);
  }

  protected String text() {
    return characters.toString();
  }

  @Override
  public void endElement(String uri, String localName, String qName) throws SAXException {
    try {
      endTag(tagPath());
      tags.remove(tags.size() - 1);
      characters.setLength(0);
    } catch (SAXException | RuntimeException e) {
      throw e;
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }

  protected abstract void endTag(String tagPath) throws Exception;

}
