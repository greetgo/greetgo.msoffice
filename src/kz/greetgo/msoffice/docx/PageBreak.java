package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

public class PageBreak implements RunElement {
  PageBreak() {}
  
  @Override
  public void write(PrintStream out) {
    out.print("<w:br w:type=\"page\" />");
  }
  
}
