package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

public class ExtendedProperties implements ContentElement {
  private final String partName;
  
  private String template = "Normal.dotm";
  private int totalTime = 5;
  private int pages = 1;
  private int words = 1;
  private int characters = 1;
  private String application = "Microsoft Office Word";
  private int docSecurity = 0;
  private int lines = 1;
  private int paragraphs = 1;
  private boolean scaleCrop = false;
  private String company = "greetgo!";
  private boolean linksUpToDate = false;
  private int charactersWithSpaces = 1;
  private boolean sharedDoc = false;
  private boolean hyperlinksChanged = false;
  private String appVersion = "12.0000";
  
  ExtendedProperties(String partName) {
    this.partName = partName;
  }
  
  @Override
  public ContentType getContentType() {
    return ContentType.EXTENDED_PROPERTIES;
  }
  
  @Override
  public String getPartName() {
    return partName;
  }
  
  public String getTemplate() {
    return template;
  }
  
  public void setTemplate(String template) {
    this.template = template;
  }
  
  public int getTotalTime() {
    return totalTime;
  }
  
  public void setTotalTime(int totalTime) {
    this.totalTime = totalTime;
  }
  
  public int getPages() {
    return pages;
  }
  
  public void setPages(int pages) {
    this.pages = pages;
  }
  
  public int getWords() {
    return words;
  }
  
  public void setWords(int words) {
    this.words = words;
  }
  
  public int getCharacters() {
    return characters;
  }
  
  public void setCharacters(int characters) {
    this.characters = characters;
  }
  
  public String getApplication() {
    return application;
  }
  
  public void setApplication(String application) {
    this.application = application;
  }
  
  public int getDocSecurity() {
    return docSecurity;
  }
  
  public void setDocSecurity(int docSecurity) {
    this.docSecurity = docSecurity;
  }
  
  public int getLines() {
    return lines;
  }
  
  public void setLines(int lines) {
    this.lines = lines;
  }
  
  public int getParagraphs() {
    return paragraphs;
  }
  
  public void setParagraphs(int paragraphs) {
    this.paragraphs = paragraphs;
  }
  
  public boolean isScaleCrop() {
    return scaleCrop;
  }
  
  public void setScaleCrop(boolean scaleCrop) {
    this.scaleCrop = scaleCrop;
  }
  
  public String getCompany() {
    return company;
  }
  
  public void setCompany(String company) {
    this.company = company;
  }
  
  public boolean isLinksUpToDate() {
    return linksUpToDate;
  }
  
  public void setLinksUpToDate(boolean linksUpToDate) {
    this.linksUpToDate = linksUpToDate;
  }
  
  public int getCharactersWithSpaces() {
    return charactersWithSpaces;
  }
  
  public void setCharactersWithSpaces(int charactersWithSpaces) {
    this.charactersWithSpaces = charactersWithSpaces;
  }
  
  public boolean isSharedDoc() {
    return sharedDoc;
  }
  
  public void setSharedDoc(boolean sharedDoc) {
    this.sharedDoc = sharedDoc;
  }
  
  public boolean isHyperlinksChanged() {
    return hyperlinksChanged;
  }
  
  public void setHyperlinksChanged(boolean hyperlinksChanged) {
    this.hyperlinksChanged = hyperlinksChanged;
  }
  
  public String getAppVersion() {
    return appVersion;
  }
  
  public void setAppVersion(String appVersion) {
    this.appVersion = appVersion;
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<?xml version=\"1.0\" encoding=\"UTF-8\"" + " standalone=\"yes\"?>\n");
    out.print("<Properties xmlns=\"http://schemas.openxmlformats.org/"
        + "officeDocument/2006/extended-properties\" "
        + "xmlns:vt=\"http://schemas.openxmlformats.org/" + "officeDocument/2006/docPropsVTypes\">");
    
    out.print("<Template>" + getTemplate() + "</Template>");
    out.print("<TotalTime>" + getTotalTime() + "</TotalTime>");
    out.print("<Pages>" + getPages() + "</Pages>");
    out.print("<Words>" + getWords() + "</Words>");
    out.print("<Characters>" + getCharacters() + "</Characters>");
    out.print("<Application>" + getApplication() + "</Application>");
    out.print("<DocSecurity>" + getDocSecurity() + "</DocSecurity>");
    out.print("<Lines>" + getLines() + "</Lines>");
    out.print("<Paragraphs>" + getParagraphs() + "</Paragraphs>");
    out.print("<ScaleCrop>" + isScaleCrop() + "</ScaleCrop>");
    out.print("<Company>" + getCompany() + "</Company>");
    out.print("<LinksUpToDate>" + isLinksUpToDate() + "</LinksUpToDate>");
    out.print("<CharactersWithSpaces>" + getCharactersWithSpaces() + "</CharactersWithSpaces>");
    out.print("<SharedDoc>" + isSharedDoc() + "</SharedDoc>");
    out.print("<HyperlinksChanged>" + isHyperlinksChanged() + "</HyperlinksChanged>");
    out.print("<AppVersion>" + getAppVersion() + "</AppVersion>");
    
    out.print("</Properties>");
  }
  
}
