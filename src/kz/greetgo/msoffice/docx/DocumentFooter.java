package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

public class DocumentFooter extends DocumentFlow {

  public DocumentFooter(String partName, MSHelper msHelper) {
    super(partName, msHelper);
  }

  @Override
  public ContentType getContentType() {
    return ContentType.FOOTER;
  }

  @Override
  protected void writeTopXml(PrintStream out) {
    out.print("<?xml version=\"1.0\" encoding=\"UTF-8\"" + " standalone=\"yes\"?>\n");
    out.print("<w:ftr xmlns:ve=\"http://schemas.openxmlformats.org/"
      + "markup-compatibility/2006\"" + " xmlns:o=\"urn:schemas-microsoft-com:office:office\""
      + " xmlns:r=\"http://schemas.openxmlformats.org/" + "officeDocument/2006/relationships\""
      + " xmlns:m=\"http://schemas.openxmlformats.org/" + "officeDocument/2006/math\""
      + " xmlns:v=\"urn:schemas-microsoft-com:vml\""
      + " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/"
      + "2006/wordprocessingDrawing\"" + " xmlns:w10=\"urn:schemas-microsoft-com:office:word\""
      + " xmlns:w=\"http://schemas.openxmlformats.org/" + "wordprocessingml/2006/main\""
      + " xmlns:wne=\"http://schemas.microsoft.com/" + "office/word/2006/wordml\">");
  }

  @Override
  protected void writeBottomXml(PrintStream out) {
    out.print("</w:ftr>");
  }

}