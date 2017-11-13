package kz.greetgo.msoffice.docx;

import java.io.PrintStream;
import java.util.Date;

import kz.greetgo.msoffice.UtilOffice;

public class CoreProperties implements ContentElement {
  private final String partName;
  private String title = "";
  private String subject = "";
  private String creator = "greetgo";
  private String keywords = "";
  private String description = "";
  private String lastModifiedBy = "greetgo";
  private String revision = "1";
  private Date created = new Date();
  private Date modified = new Date();
  
  CoreProperties(String partName) {
    this.partName = partName;
  }
  
  @Override
  public String getPartName() {
    return partName;
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<?xml version=\"1.0\" encoding=\"UTF-8\" " + "standalone=\"yes\"?>\n");
    out.print("<cp:coreProperties" + " xmlns:cp=\"http://schemas.openxmlformats.org/"
        + "package/2006/metadata/core-properties\""
        + " xmlns:dc=\"http://purl.org/dc/elements/1.1/\""
        + " xmlns:dcterms=\"http://purl.org/dc/terms/\""
        + " xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\""
        + " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
    
    out.print("<dc:title>" + getTitle() + "</dc:title>");
    out.print("<dc:subject>" + getSubject() + "</dc:subject>");
    out.print("<dc:creator>" + getCreator() + "</dc:creator>");
    out.print("<cp:keywords>" + getKeywords() + "</cp:keywords>");
    out.print("<dc:description>" + getDescription() + "</dc:description>");
    out.print("<cp:lastModifiedBy>" + getLastModifiedBy() + "</cp:lastModifiedBy>");
    out.print("<cp:revision>" + getRevision() + "</cp:revision>");
    out.print("<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + UtilOffice.toW3CDTF(getCreated())
        + "</dcterms:created>");
    out.print("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + UtilOffice.toW3CDTF(getModified())
        + "</dcterms:modified>");
    
    out.print("</cp:coreProperties>");
  }
  
  public String getTitle() {
    return title;
  }
  
  public void setTitle(String title) {
    this.title = title;
  }
  
  public String getSubject() {
    return subject;
  }
  
  public void setSubject(String subject) {
    this.subject = subject;
  }
  
  public String getKeywords() {
    return keywords;
  }
  
  public void setKeywords(String keywords) {
    this.keywords = keywords;
  }
  
  public String getDescription() {
    return description;
  }
  
  public void setDescription(String description) {
    this.description = description;
  }
  
  public String getLastModifiedBy() {
    return lastModifiedBy;
  }
  
  public void setLastModifiedBy(String lastModifiedBy) {
    this.lastModifiedBy = lastModifiedBy;
  }
  
  public String getRevision() {
    return revision;
  }
  
  public void setRevision(String revision) {
    this.revision = revision;
  }
  
  public Date getCreated() {
    return created;
  }
  
  public void setCreated(Date created) {
    this.created = created;
  }
  
  public Date getModified() {
    return modified;
  }
  
  public void setModified(Date modified) {
    this.modified = modified;
  }
  
  public void setCreator(String creator) {
    this.creator = creator;
  }
  
  public String getCreator() {
    return creator;
  }
  
  @Override
  public ContentType getContentType() {
    return ContentType.CORE_PROPERTIES;
  }
}
