package kz.greetgo.msoffice.docx;

public enum ContentType {
  DOCUMENT("application/vnd.openxmlformats-officedocument" + ".wordprocessingml.document.main+xml"),
  
  STYLES("application/vnd.openxmlformats-officedocument" + ".wordprocessingml.styles+xml"),
  
  ENDNOTES("application/vnd.openxmlformats-officedocument" + ".wordprocessingml.endnotes+xml"),
  
  EXTENDED_PROPERTIES("application/vnd.openxmlformats-officedocument" + ".extended-properties+xml"),
  
  SETTINGS("application/vnd.openxmlformats-officedocument" + ".wordprocessingml.settings+xml"),
  
  FOOTER("application/vnd.openxmlformats-officedocument" + ".wordprocessingml.footer+xml"),
  
  THEME("application/vnd.openxmlformats-officedocument.theme+xml"),
  
  FONT_TABLE("application/vnd.openxmlformats-officedocument" + ".wordprocessingml.fontTable+xml"),
  
  WEB_SETTINGS("application/vnd.openxmlformats-officedocument"
      + ".wordprocessingml.webSettings+xml"),
  
  HEADER("application/vnd.openxmlformats-officedocument" + ".wordprocessingml.header+xml"),
  
  CORE_PROPERTIES("application/vnd.openxmlformats-package" + ".core-properties+xml"),
  
  APPLICATION_XML("application/xml"),
  
  RELATIONSHIPS("application/vnd.openxmlformats-package.relationships+xml"),
  
  IMAGE_PNG("image/png"),
  
  ;
  
  private final String xmlns;
  
  private ContentType(String xmlns) {
    this.xmlns = xmlns;
  }
  
  public String getXmlns() {
    return xmlns;
  }
  
}
