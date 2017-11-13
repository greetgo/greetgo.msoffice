package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

/**
 * Заливка узором
 * 
 * @author pompei
 */
public class PatternFill extends Fill {
  private PatternFillType type = PatternFillType.none;
  private Color fgColor, bgColor;
  
  public PatternFill() {}
  
  public PatternFill(PatternFillType type) {
    this.type = type;
  }
  
  public PatternFill(PatternFill x) {
    type = x.type;
    fgColor = Color.copy(x.fgColor);
    bgColor = Color.copy(x.bgColor);
  }
  
  public PatternFill none() {
    type = PatternFillType.none;
    fgColor = null;
    bgColor = null;
    return this;
  }
  
  public PatternFill solid() {
    type = PatternFillType.solid;
    return this;
  }
  
  public PatternFill type(PatternFillType type) {
    this.type = type;
    return this;
  }
  
  @Override
  Fill copy() {
    return new PatternFill(this);
  }
  
  @Override
  public int hashCode() {
    final int prime = 31;
    int result = 1;
    result = prime * result + ((bgColor == null) ? 0 :bgColor.hashCode());
    result = prime * result + ((fgColor == null) ? 0 :fgColor.hashCode());
    result = prime * result + ((type == null) ? 0 :type.hashCode());
    return result;
  }
  
  @Override
  public boolean equals(Object obj) {
    if (this == obj) return true;
    if (obj == null) return false;
    if (getClass() != obj.getClass()) return false;
    PatternFill other = (PatternFill)obj;
    if (bgColor == null) {
      if (other.bgColor != null) return false;
    } else if (!bgColor.equals(other.bgColor)) return false;
    if (fgColor == null) {
      if (other.fgColor != null) return false;
    } else if (!fgColor.equals(other.fgColor)) return false;
    if (type != other.type) return false;
    return true;
  }
  
  public PatternFillType getType() {
    return type;
  }
  
  public void setType(PatternFillType type) {
    this.type = type;
  }
  
  public Color getFgColor() {
    return fgColor;
  }
  
  public void setFgColor(Color fgColor) {
    this.fgColor = fgColor;
  }
  
  public PatternFill fgColor(Color fgColor) {
    setFgColor(fgColor);
    return this;
  }
  
  public PatternFill bgColor(Color bgColor) {
    setBgColor(bgColor);
    return this;
  }
  
  public Color getBgColor() {
    return bgColor;
  }
  
  public void setBgColor(Color bgColor) {
    this.bgColor = bgColor;
  }
  
  @Override
  void print(PrintStream out) {
    if (fgColor == null && bgColor == null) {
      out.println("<fill>");
      out.println("<patternFill patternType=\"" + type.name() + "\" />");
      out.println("</fill>");
      return;
    }
    out.println("<fill>");
    out.println("<patternFill patternType=\"" + type.name() + "\" >");
    if (fgColor != null) out.println("<fgColor rgb=\"" + fgColor.str() + "\" />");
    if (bgColor != null) out.println("<bgColor rgb=\"" + bgColor.str() + "\" />");
    out.println("</patternFill>");
    out.println("</fill>");
  }
  
}
