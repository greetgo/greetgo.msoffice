package kz.greetgo.msoffice.docx;

import java.util.List;

public class FieldAppender {
  
  private final List<RunElement> elements;
  
  FieldAppender(List<RunElement> elements) {
    this.elements = elements;
  }
  
  public void addPageNomer() {
    elements.add(new RunFldChar(FldCharType.begin));
    elements.add(RunInstrText.PAGE);
    elements.add(new RunFldChar(FldCharType.end));
  }
  
  public void addPageCount() {
    elements.add(new RunFldChar(FldCharType.begin));
    elements.add(RunInstrText.NUMPAGES);
    elements.add(new RunFldChar(FldCharType.end));
  }
}
