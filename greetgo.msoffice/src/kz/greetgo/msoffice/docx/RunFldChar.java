package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

public class RunFldChar implements RunElement {
  private final FldCharType type;
  
  RunFldChar(FldCharType type) {
    assert type != null;
    this.type = type;
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<w:fldChar w:fldCharType=\"" + type.name() + "\" />");
  }
}
