package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

class Reference implements XmlWriter {
  
  private String id;
  private String tagName;
  
  @Override
  public void write(PrintStream out) {
    out.print("<" + getTagName() + " w:type=\"default\" r:id=\"" + getId() + "\" />");
  }
  
  public void setTagName(String tagName) {
    this.tagName = tagName;
  }
  
  public String getTagName() {
    return tagName;
  }
  
  public void setId(String id) {
    this.id = id;
  }
  
  public String getId() {
    return id;
  }
  
}
