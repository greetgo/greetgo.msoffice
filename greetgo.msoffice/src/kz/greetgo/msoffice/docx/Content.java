package kz.greetgo.msoffice.docx;

import java.io.PrintStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

class Content implements FilePart {
  private final MSHelper msHelper;
  
  static class Default implements XmlWriter {
    private final String extention;
    private final ContentType contentType;
    
    public Default(ContentType contentType, String extention) {
      this.contentType = contentType;
      this.extention = extention;
      assert contentType != null;
      assert extention != null && extention.length() > 0;
    }
    
    @Override
    public void write(PrintStream out) {
      out.print("<Default Extension=\"" + extention + "\" ContentType=\"" + contentType.getXmlns()
          + "\" />");
    }
    
    public String getExtention() {
      return extention;
    }
    
    public ContentType getContentType() {
      return contentType;
    }
  }
  
  private boolean defaultImagePngExists = false;
  
  void checkExistsDefaultImagePng() {
    if (defaultImagePngExists) return;
    defaults.add(new Default(ContentType.IMAGE_PNG, "png"));
    defaultImagePngExists = true;
  }
  
  static class ContOverride implements XmlWriter {
    private final ContentElement contentElement;
    
    private ContOverride(ContentElement contentElement) {
      this.contentElement = contentElement;
    }
    
    public String getPartName() {
      return contentElement.getPartName();
    }
    
    public ContentType getContentType() {
      return contentElement.getContentType();
    }
    
    @Override
    public void write(PrintStream out) {
      out.print("<Override PartName=\"" + getPartName() + "\" ContentType=\""
          + getContentType().getXmlns() + "\" />");
    }
    
    public ContentElement getContentElement() {
      return contentElement;
    }
  }
  
  private Content(MSHelper msHelper) {
    this.msHelper = msHelper;
  }
  
  private List<Default> defaults = new ArrayList<Default>();
  private List<ContOverride> overrides = new ArrayList<ContOverride>();
  private FontTableContentElement fontTableContentElement;
  
  static Content createDefaultContent(MSHelper msHelper) {
    Content ret = new Content(msHelper);
    ret.initDefault();
    return ret;
  }
  
  private void initDefault() {
    {
      defaults.add(new Default(ContentType.APPLICATION_XML, "xml"));
      defaults.add(new Default(ContentType.RELATIONSHIPS, "rels"));
    }
    {
      overrides.add(new ContOverride(new CoreProperties("/docProps/core.xml")));
    }
    {
      overrides.add(new ContOverride(new ExtendedProperties("/docProps/app.xml")));
    }
    {
      overrides.add(new ContOverride(new Document("/word/document.xml", msHelper)));
    }
  }
  
  FontTableContentElement getFontTableContentElement() {
    if (fontTableContentElement == null) {
      fontTableContentElement = FontTableContentElement.createDefault("/word/fontTable.xml");
      overrides.add(new ContOverride(fontTableContentElement));
    }
    return fontTableContentElement;
  }
  
  void addFontTableOverride() {
    overrides.add(new ContOverride(new CoreProperties("/docProps/core.xml")));
  }
  
  public CoreProperties getCoreProperties() {
    for (ContOverride o : overrides) {
      if (ContentType.CORE_PROPERTIES == o.getContentType()) {
        return (CoreProperties)o.getContentElement();
      }
    }
    throw new IllegalStateException("No core properties");
  }
  
  public ExtendedProperties getExtendedProperties() {
    for (ContOverride o : overrides) {
      if (ContentType.EXTENDED_PROPERTIES == o.getContentType()) {
        return (ExtendedProperties)o.getContentElement();
      }
    }
    throw new IllegalStateException("No extended properties");
  }
  
  public Document getDocument() {
    for (ContOverride o : overrides) {
      if (ContentType.DOCUMENT == o.getContentType()) {
        return (Document)o.getContentElement();
      }
    }
    throw new IllegalStateException("No document");
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<?xml version=\"1.0\" encoding=\"UTF-8\" " + "standalone=\"yes\"?>\n"
        + "<Types xmlns=\"http://schemas.openxmlformats.org/" + "package/2006/content-types\">");
    for (Default d : defaults) {
      d.write(out);
    }
    for (ContOverride o : overrides) {
      o.write(out);
    }
    out.print("</Types>");
  }
  
  DocumentHeader getOrCreateHeader() {
    for (ContOverride o : overrides) {
      if (ContentType.HEADER == o.getContentType()) {
        return (DocumentHeader)o.getContentElement();
      }
    }
    {
      DocumentHeader o = new DocumentHeader("/word/header1.xml", msHelper);
      overrides.add(new ContOverride(o));
      return o;
    }
  }
  
  DocumentFooter getOrCreateFooter() {
    for (ContOverride o : overrides) {
      if (ContentType.FOOTER == o.getContentType()) {
        return (DocumentFooter)o.getContentElement();
      }
    }
    {
      DocumentFooter o = new DocumentFooter("/word/footer1.xml", msHelper);
      overrides.add(new ContOverride(o));
      return o;
    }
  }
  
  @Override
  public String getPartName() {
    return "[Content_Types].xml";
  }
  
  public void addAllFileParts(Collection<FilePart> pull) {
    for (ContOverride o : overrides) {
      pull.add(o.getContentElement());
    }
  }
}
