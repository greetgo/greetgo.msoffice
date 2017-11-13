package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

public class RunInstrText implements RunElement {
  private final String text;
  
  static final RunInstrText PAGE = new RunInstrText("PAGE");
  static final RunInstrText NUMPAGES = new RunInstrText("NUMPAGES");
  
  private RunInstrText(String text) {
    this.text = text;
    assert text != null;
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<w:instrText>" + text + "</w:instrText>");
  }
}
