package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

public class RunBr implements RunElement {
  RunBr() {}
  
  @Override
  public void write(PrintStream out) {
    out.print("<w:br />");
  }
  
}
