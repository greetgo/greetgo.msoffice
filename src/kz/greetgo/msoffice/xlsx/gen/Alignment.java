package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

public class Alignment {
  private Align horizontal, vertical;
  private boolean wrapText = false;
  
  public void clean() {
    horizontal = null;
    vertical = null;
    wrapText = false;
  }
  
  public Alignment() {}
  
  public Alignment(Alignment x) {
    horizontal = x.horizontal;
    vertical = x.vertical;
    wrapText = x.wrapText;
  }
  
  boolean isStandart() {
    return horizontal == null && vertical == null && (wrapText == false);
  }
  
  @Override
  public int hashCode() {
    final int prime = 31;
    int result = 1;
    result = prime * result + ((horizontal == null) ? 0 :horizontal.hashCode());
    result = prime * result + ((vertical == null) ? 0 :vertical.hashCode());
    result = prime * result + (wrapText ? 1231 :1237);
    return result;
  }
  
  @Override
  public boolean equals(Object obj) {
    if (this == obj) return true;
    if (obj == null) return false;
    if (getClass() != obj.getClass()) return false;
    Alignment other = (Alignment)obj;
    if (horizontal != other.horizontal) return false;
    if (vertical != other.vertical) return false;
    if (wrapText != other.wrapText) return false;
    return true;
  }
  
  public Align horizontal() {
    return horizontal;
  }
  
  public Align vertical() {
    return vertical;
  }
  
  public void setHorizontal(Align horizontal) {
    this.horizontal = horizontal;
  }
  
  public void setVertical(Align vertical) {
    this.vertical = vertical;
  }
  
  public Alignment horizontalCenter() {
    horizontal = Align.center;
    return this;
  }
  
  public Alignment verticalCenter() {
    vertical = Align.center;
    return this;
  }
  
  public Alignment horizontalNull() {
    horizontal = null;
    return this;
  }
  
  public Alignment verticalNull() {
    vertical = null;
    return this;
  }
  
  public Alignment left() {
    horizontal = Align.left;
    return this;
  }
  
  public Alignment right() {
    horizontal = Align.right;
    return this;
  }
  
  public Alignment top() {
    vertical = Align.top;
    return this;
  }
  
  public Alignment bottom() {
    vertical = Align.bottom;
    return this;
  }
  
  public void print(PrintStream out) {
    out.print("<alignment");
    if (horizontal != null) {
      out.print(" horizontal=\"" + horizontal.name() + "\"");
    }
    if (vertical != null) {
      out.print(" vertical=\"" + vertical.name() + "\"");
    }
    if (isWrapText()) {
      out.print(" wrapText=\"1\"");
    }
    out.println(" />");
  }
  
  public boolean isWrapText() {
    return wrapText;
  }
  
  public void setWrapText(boolean wrapText) {
    this.wrapText = wrapText;
  }
  
  public Alignment wrap() {
    setWrapText(true);
    return this;
  }
  
  public Alignment noWrap() {
    setWrapText(false);
    return this;
  }
}
