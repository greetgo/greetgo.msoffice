package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

public class Font implements XmlWriter {
  String name;
  
  Font() {}
  
  public String getName() {
    return name;
  }
  
  @Override
  public void write(PrintStream out) {
    String n = getName();
    out.print("<w:rFonts w:ascii=\"" + n + "\" w:hAnsi=\"" + n + "\" w:cs=\"" + n + "\" />");
  }
}
