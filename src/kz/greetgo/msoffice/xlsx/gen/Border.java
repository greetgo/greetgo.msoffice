package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

public class Border {
  private BorderStyle style;
  private Color color;
  
  public Border() {}
  
  public Border(Border x) {
    style = x.style;
    color = Color.copy(x.color);
  }
  
  public static Border copy(Border x) {
    if (x == null) return null;
    return new Border(x);
  }
  
  public BorderStyle getStyle() {
    return style;
  }
  
  public void setStyle(BorderStyle style) {
    this.style = style;
  }
  
  public Color getColor() {
    return color;
  }
  
  public void setColor(Color color) {
    this.color = color;
  }
  
  @Override
  public int hashCode() {
    final int prime = 31;
    int result = 1;
    result = prime * result + ((color == null) ? 0 :color.hashCode());
    result = prime * result + ((style == null) ? 0 :style.hashCode());
    return result;
  }
  
  @Override
  public boolean equals(Object obj) {
    if (this == obj) return true;
    if (obj == null) return false;
    if (getClass() != obj.getClass()) return false;
    Border other = (Border)obj;
    if (style != other.style) return false;
    
    if (style == null && other.style == null) return true;
    
    if (color == null) {
      if (other.color != null) return false;
    } else if (!color.equals(other.color)) return false;
    
    return true;
  }
  
  public void print(PrintStream out, String tagName) {
    if (style == null) {
      out.println("<" + tagName + " />");
      return;
    }
    out.print('<');
    out.print(tagName);
    out.println(" style=\"" + style.str() + "\">");
    if (color == null) {
      out.println("<color auto=\"1\" />");
    } else {
      out.println("<color rgb=\"" + color.str() + "\" />");
    }
    out.println("</" + tagName + ">");
  }
  
  public Border noStyle() {
    style = null;
    return this;
  }
  
  public Border noColor() {
    color = null;
    return this;
  }
  
  public Border clean() {
    style = null;
    color = null;
    return this;
  }
  
  public Border medium() {
    style = BorderStyle.medium;
    return this;
  }
  
  public Border thin() {
    style = BorderStyle.thin;
    return this;
  }
  
  public Border thick() {
    style = BorderStyle.thick;
    return this;
  }
  
  public Border style(BorderStyle style) {
    this.style = style;
    return this;
  }
}
