package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

public class RunTab implements RunElement {
  RunTab() {}
  
  @Override
  public void write(PrintStream out) {
    out.print("<w:tab />");
  }
}
