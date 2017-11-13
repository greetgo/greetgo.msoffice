package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

class Relationship implements XmlWriter {
  public static enum Type {
    HEADER("http://schemas.openxmlformats.org/officeDocument/2006/" + "relationships/header"),
    
    FOOTER("http://schemas.openxmlformats.org/officeDocument/2006/" + "relationships/footer"),
    
    OFFICE_DOCUMENT("http://schemas.openxmlformats.org/"
        + "officeDocument/2006/relationships/officeDocument"),
    
    CORE_PROPERTIES("http://schemas.openxmlformats.org/package/2006/"
        + "relationships/metadata/core-properties"),
    
    EXTENDED_PROPERTIES("http://schemas.openxmlformats.org/"
        + "officeDocument/2006/relationships/extended-properties"),
    
    IMAGE("http://schemas.openxmlformats.org/" + "officeDocument/2006/relationships/image"),
    
    FONT_TABLE("http://schemas.openxmlformats.org/" + "officeDocument/2006/relationships/fontTable"),
    
    ;
    private final String xmlns;
    
    private Type(String xmlns) {
      this.xmlns = xmlns;
    }
    
    public String getXmlns() {
      return xmlns;
    }
  }
  
  private Type type;
  private String id;
  private String target;
  
  public Type getType() {
    return type;
  }
  
  public void setType(Type type) {
    this.type = type;
  }
  
  public String getId() {
    return id;
  }
  
  public void setId(String id) {
    this.id = id;
  }
  
  public String getTarget() {
    return target;
  }
  
  public void setTarget(String target) {
    this.target = target;
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<Relationship Id=\"" + getId() + "\" Type=\"" + getType().getXmlns()
        + "\" Target=\"" + getTarget() + "\" />");
  }
  
}
