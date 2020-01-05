package kz.greetgo.msoffice.xlsx.reader;

import org.xml.sax.Attributes;

public class StringsHandler extends AbstractXmlHandler {

  private final StoredStrings storedStrings;

  StringsHandler(StoredStrings storedStrings) {
    this.storedStrings = storedStrings;
  }


  @Override
  protected void startTag(String tagPath, Attributes attributes) {}

  @Override
  protected void endTag(String tagPath) {
    if ("/sst/si/t".equals(tagPath)) {
      storedStrings.append(text());
    }
  }

}
