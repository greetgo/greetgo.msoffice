package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

public class RunText implements RunElement {
  private final String text;
  
  RunText(String text) {
    this.text = text;
    assert text != null;
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<w:t xml:space=\"preserve\"><![CDATA[" + text + "]]></w:t>");
  }
}
