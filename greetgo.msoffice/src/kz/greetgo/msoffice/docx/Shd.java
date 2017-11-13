package kz.greetgo.msoffice.docx;

import java.awt.Color;
import java.io.PrintStream;

import kz.greetgo.msoffice.UtilOffice;

public class Shd implements XmlWriter {
  private String val;
  private String fill;
  private Color color = null;
  
  @Override
  public void write(PrintStream out) {
    if (getVal() == null) return;
    if (getFill() == null) return;
    String sColor = "auto";
    if (getColor() != null) {
      sColor = UtilOffice.toHEX(getColor());
    }
    out.print("<w:shd w:val=\"" + val + "\" w:color=\"" + sColor + "\" w:fill=\"" + fill + "\" />");
  }
  
  public String getVal() {
    return val;
  }
  
  public void setVal(String val) {
    this.val = val;
  }
  
  public Color getColor() {
    return color;
  }
  
  public void setColor(Color color) {
    this.color = color;
  }
  
  public String getFill() {
    return fill;
  }
  
  public void setFill(String fill) {
    this.fill = fill;
  }
}
