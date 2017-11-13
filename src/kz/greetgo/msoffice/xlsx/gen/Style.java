package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

public class Style {
  NumFmt numFmt;
  Font font;
  Fill fill;
  final Alignment alignment;
  final Borders borders;
  
  private final Styles styles;
  
  public void clean() {
    font = null;
    fill = null;
    alignment.clean();
    borders.clean();
    clearNumFmt();
  }
  
  public Font font() {
    if (font == null) {
      font = new Font();
    }
    return font;
  }
  
  public PatternFill patternFill() {
    if (fill instanceof PatternFill) {
      return (PatternFill)fill;
    }
    {
      PatternFill ret = new PatternFill();
      fill = ret;
      return ret;
    }
  }
  
  public Fill getFill() {
    if (fill == null) return new PatternFill(PatternFillType.none);
    return fill;
  }
  
  public void noFill() {
    fill = null;
  }
  
  int index() {
    return styles.styleIndex(this);
  }
  
  Style(Styles styles) {
    this.styles = styles;
    if (styles == null) throw new NullPointerException("styles == null");
    alignment = new Alignment();
    borders = new Borders();
  }
  
  Style(Style x) {
    numFmt = x.numFmt;
    font = Font.copy(x.font);
    fill = Fill.copy(x.fill);
    alignment = new Alignment(x.alignment);
    borders = Borders.copy(x.borders);
    styles = x.styles;
  }
  
  public Style copy() {
    return new Style(this);
  }
  
  @Override
  public int hashCode() {
    final int prime = 31;
    int result = 1;
    result = prime * result + ((alignment == null) ? 0 :alignment.hashCode());
    result = prime * result + ((borders == null) ? 0 :borders.hashCode());
    result = prime * result + ((fill == null) ? 0 :fill.hashCode());
    result = prime * result + ((font == null) ? 0 :font.hashCode());
    result = prime * result + ((numFmt == null) ? 0 :numFmt.hashCode());
    return result;
  }
  
  @Override
  public boolean equals(Object obj) {
    if (this == obj) return true;
    if (obj == null) return false;
    if (getClass() != obj.getClass()) return false;
    Style other = (Style)obj;
    if (alignment == null) {
      if (other.alignment != null) return false;
    } else if (!alignment.equals(other.alignment)) return false;
    if (borders == null) {
      if (other.borders != null) return false;
    } else if (!borders.equals(other.borders)) return false;
    if (fill == null) {
      if (other.fill != null) return false;
    } else if (!fill.equals(other.fill)) return false;
    if (font == null) {
      if (other.font != null) return false;
    } else if (!font.equals(other.font)) return false;
    if (numFmt != other.numFmt) return false;
    return true;
  }
  
  void clearNumFmt() {
    numFmt = null;
  }
  
  public Borders borders() {
    return borders;
  }
  
  public void print(PrintStream out) {
    out.print("<xf");
    boolean applyNumberFormat = false;
    if (numFmt == null) {
      out.print(" numFmtId=\"0\"");
    } else {
      out.print(" numFmtId=\"" + numFmt.getId() + "\"");
      applyNumberFormat = true;
    }
    
    int fontIndex = styles.fontIndex(font);
    out.print(" fontId=\"" + fontIndex + "\"");
    
    int bordersIndex = styles.bordersIndex(borders);
    out.print(" borderId=\"" + bordersIndex + "\"");
    
    int fillIndex = styles.fillIndex(fill);
    out.print(" fillId=\"" + fillIndex + "\"");
    
    if (bordersIndex > 0) {
      out.print(" applyBorder=\"1\"");
    }
    if (fontIndex > 0) {
      out.print(" applyFont=\"1\"");
    }
    if (fillIndex > 0) {
      out.print(" applyFill=\"1\"");
    }
    if (applyNumberFormat) {
      out.print(" applyNumberFormat=\"1\"");
    }
    if (!alignment.isStandart()) {
      out.print(" applyAlignment=\"1\"");
    }
    
    out.println(" xfId=\"0\" >");
    if (!alignment.isStandart()) {
      alignment.print(out);
    }
    out.println("</xf>");
  }
  
  public Alignment alignment() {
    return alignment;
  }
  
}
