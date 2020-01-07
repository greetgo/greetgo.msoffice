package kz.greetgo.msoffice.docx;

import kz.greetgo.msoffice.PageOrientation;

import java.io.PrintStream;

public class Document extends DocumentFlow {

  private PageOrientation pageOrientation = PageOrientation.PORTRAIT;
  private Integer top = 1134;
  private Integer bottom = 1134;
  private Integer left = 1701;
  private Integer right = null;

  public PageOrientation getPageOrientation() {
    return pageOrientation;
  }

  public void setPageOrientation(PageOrientation pageOrientation) {
    this.pageOrientation = pageOrientation;
  }

  public Integer getTop() {
    return top;
  }

  public void setTop(Integer top) {
    this.top = top;
  }

  public Integer getBottom() {
    return bottom;
  }

  public void setBottom(Integer bottom) {
    this.bottom = bottom;
  }

  public Integer getLeft() {
    return left;
  }

  public void setLeft(Integer left) {
    this.left = left;
  }

  public Integer getRight() {
    return right;
  }

  public void setRight(Integer right) {
    this.right = right;
  }

  Document(String partName, MSHelper msHelper) {
    super(partName, msHelper);
  }

  @Override
  public ContentType getContentType() {
    return ContentType.DOCUMENT;
  }

  @Override
  protected void writeTopXml(PrintStream out) {
    out.print("<?xml version=\"1.0\" encoding=\"UTF-8\"" + " standalone=\"yes\"?>\n");
    out.print("<w:document xmlns:ve=\"http://schemas.openxmlformats.org/"
      + "markup-compatibility/2006\"" + " xmlns:o=\"urn:schemas-microsoft-com:office:office\""
      + " xmlns:r=\"http://schemas.openxmlformats.org/" + "officeDocument/2006/relationships\""
      + " xmlns:m=\"http://schemas.openxmlformats.org/" + "officeDocument/2006/math\""
      + " xmlns:v=\"urn:schemas-microsoft-com:vml\""
      + " xmlns:wp=\"http://schemas.openxmlformats.org/"
      + "drawingml/2006/wordprocessingDrawing\""
      + " xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\""
      + " xmlns:w10=\"urn:schemas-microsoft-com:office:word\""
      + " xmlns:w=\"http://schemas.openxmlformats.org/" + "wordprocessingml/2006/main\""
      + " xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/" + "wordml\">");
    out.print("<w:body>");
  }

  @Override
  protected void writeBottomXml(PrintStream out) {

    out.print("<w:sectPr>");
    writeReferences(out);
    out.print("<w:pgSz w:w=\"" + pageOrientation.width + "\" w:h=\"" + pageOrientation.height + "\" />");
    out.print("<w:pgMar w:top=\"" + getTop() + "\" w:bottom=\"" + getBottom() + "\"" + " w:left=\""
      + getLeft() + "\"");
    if (right != null) out.print(" w:right=\"" + getRight() + "\"");
    out.print(" w:header=\"708\" w:footer=\"708\"" + " w:gutter=\"0\" />");
    out.print("<w:cols w:space=\"708\" />");
    out.print("<w:docGrid w:linePitch=\"360\" />");
    out.print("</w:sectPr>");

    out.print("</w:body>");
    out.print("</w:document>");
  }

  public void addPageBreak() {
    createPara().addPageBreak();
  }
}
