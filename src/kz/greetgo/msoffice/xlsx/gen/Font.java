package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

public class Font {
  private FontName name = FontName.Calibri;
  private int size = 11;
  private Color color = null;
  private boolean bold = false;
  private boolean italic = false;
  private boolean underlined = false;
  
  public Font() {}
  
  public Font(Font x) {
    name = x.name;
    size = x.size;
    color = Color.copy(x.color);
    bold = x.bold;
    italic = x.italic;
    underlined = x.underlined;
  }
  
  public static Font copy(Font font) {
    if (font == null) return null;
    return new Font(font);
  }
  
  @Override
  public int hashCode() {
    final int prime = 31;
    int result = 1;
    result = prime * result + (bold ? 1231 :1237);
    result = prime * result + ((color == null) ? 0 :color.hashCode());
    result = prime * result + (italic ? 1231 :1237);
    result = prime * result + ((name == null) ? 0 :name.hashCode());
    result = prime * result + size;
    result = prime * result + (underlined ? 1231 :1237);
    return result;
  }
  
  @Override
  public boolean equals(Object obj) {
    if (this == obj) return true;
    if (obj == null) return false;
    if (getClass() != obj.getClass()) return false;
    Font other = (Font)obj;
    if (bold != other.bold) return false;
    if (color == null) {
      if (other.color != null) return false;
    } else if (!color.equals(other.color)) return false;
    if (italic != other.italic) return false;
    if (name != other.name) return false;
    if (size != other.size) return false;
    if (underlined != other.underlined) return false;
    return true;
  }
  
  public FontName getName() {
    return name;
  }
  
  public void setName(FontName name) {
    this.name = name;
  }
  
  public int getSize() {
    return size;
  }
  
  public void setSize(int size) {
    this.size = size;
  }
  
  public Color getColor() {
    return color;
  }
  
  public void setColor(Color color) {
    this.color = color;
  }
  
  public boolean isBold() {
    return bold;
  }
  
  public void setBold(boolean bold) {
    this.bold = bold;
  }
  
  public boolean isItalic() {
    return italic;
  }
  
  public void setItalic(boolean italic) {
    this.italic = italic;
  }
  
  public boolean isUnderlined() {
    return underlined;
  }
  
  public void setUnderlined(boolean underlined) {
    this.underlined = underlined;
  }
  
  public void print(PrintStream out) {
    out.println("<font>");
    if (italic) out.println("<i />");
    if (underlined) out.println("<u />");
    if (bold) out.println("<b />");
    out.println("<sz val=\"" + size + "\" />");
    if (color != null) out.println("<color rgb=\"" + color.str() + "\" />");
    out.println("<name val=\"" + name.str() + "\" />");
    out.println("<family val=\"" + name.getFamily() + "\" />");
    out.println("<charset val=\"" + name.getCharset() + "\" />");
    if (name.getScheme() != null) out.println("<scheme val=\"" + name.getScheme() + "\" />");
    out.println("</font>");
  }
  
  public Font bold() {
    setBold(true);
    return this;
  }
  
  public Font noBold() {
    setBold(false);
    return this;
  }
  
  public Font italic() {
    setItalic(true);
    return this;
  }
  
  public Font noItalic() {
    setItalic(false);
    return this;
  }
  
  public Font underlined() {
    setUnderlined(true);
    return this;
  }
  
  public Font noUnderlined() {
    setUnderlined(false);
    return this;
  }
  
  public Font color(Color color) {
    setColor(color);
    return this;
  }
  
  public Font size(int size) {
    setSize(size);
    return this;
  }
  
  public Font name(FontName name) {
    setName(name);
    return this;
  }
}
