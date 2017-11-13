package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;
import java.text.SimpleDateFormat;
import java.util.Date;

public class CoreProperties {
  private String creator = "XlsxGenerator";
  private String lastModifiedBy = "XlsxGenerator";
  private Date created = new Date();
  private Date modified = new Date();
  
  public String getCreator() {
    return creator;
  }
  
  public void setCreator(String creator) {
    this.creator = creator;
  }
  
  public Date getCreated() {
    return created;
  }
  
  public void setCreated(Date created) {
    this.created = created;
  }
  
  public String getLastModifiedBy() {
    return lastModifiedBy;
  }
  
  public Date getModified() {
    return modified;
  }
  
  public void setModified(Date modified) {
    this.modified = modified;
  }
  
  public void setLastModifiedBy(String lastModifiedBy) {
    this.lastModifiedBy = lastModifiedBy;
  }
  
  private static final SimpleDateFormat DF = new SimpleDateFormat("yyyy-MM-dd'T'hh:mm:ss'Z'");
  
  public void print(PrintStream out) {
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    
    out.println("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/"
        + "2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\""
        + " xmlns:dcterms=\"http://purl.org/dc/terms/\""
        + " xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\""
        + " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
    
    out.println("<dc:creator>" + creator + "</dc:creator>");
    out.println("<cp:lastModifiedBy>" + lastModifiedBy + "</cp:lastModifiedBy>");
    out.println("<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + DF.format(created)
        + "</dcterms:created>");
    out.println("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + DF.format(modified)
        + "</dcterms:modified>");
    
    out.println("</cp:coreProperties>");
  }
}
