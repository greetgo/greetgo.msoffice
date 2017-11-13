package kz.greetgo.msoffice.docx;

import java.awt.Color;
import java.io.PrintStream;

import kz.greetgo.msoffice.UtilOffice;

public class Borders implements XmlWriter {
  private final String tagName;
  
  public static enum Side {
    LEFT("w:left"), TOP("w:top"), RIGHT("w:right"), BOTTOM("w:bottom"), INSIDE_H("w:insideH"),
    INSIDE_V("w:insideV");
    
    private final String code;
    
    private Side(String code) {
      this.code = code;
    }
    
    public String getCode() {
      return code;
    }
  }
  
  public Borders(String tagName) {
    this.tagName = tagName;
  }
  
  public static class Border implements XmlWriter {
    private final Side side;
    
    private Border(Side side) {
      this.side = side;
    }
    
    private LineStyle lineStyle = LineStyle.NONE;
    private int size = 4, space = 0;
    private Color color = null;
    
    public Side getSide() {
      return side;
    }
    
    public void setSize(int size) {
      this.size = size;
    }
    
    public int getSize() {
      return size;
    }
    
    public void setSpace(int space) {
      this.space = space;
    }
    
    public int getSpace() {
      return space;
    }
    
    public void setColor(Color color) {
      this.color = color;
    }
    
    public Color getColor() {
      return color;
    }
    
    public void setLineStyle(LineStyle lineStyle) {
      this.lineStyle = lineStyle;
    }
    
    public LineStyle getLineStyle() {
      return lineStyle;
    }
    
    public boolean isEmpty() {
      if (getLineStyle() == null) return true;
      if (getLineStyle() == LineStyle.NONE) return true;
      return false;
    }
    
    @Override
    public void write(PrintStream out) {
      if (isEmpty()) return;
      String sColor = "auto";
      if (getColor() != null) {
        sColor = UtilOffice.toHEX(getColor());
      }
      out.print("<" + getSide().getCode() + " w:val=\"" + getLineStyle().getCode() + "\" w:sz=\""
          + getSize() + "\" w:space=\"" + getSpace() + "\" w:color=\"" + sColor + "\" />");
    }
  }
  
  private final Border left = new Border(Side.LEFT);
  private final Border right = new Border(Side.RIGHT);
  private final Border top = new Border(Side.TOP);
  private final Border bottom = new Border(Side.BOTTOM);
  private final Border insideH = new Border(Side.INSIDE_H);
  private final Border insideV = new Border(Side.INSIDE_V);
  
  @Override
  public void write(PrintStream out) {
    if (isEmpty()) return;
    out.print("<" + getTagName() + ">");
    {
      getLeft().write(out);
      getTop().write(out);
      getRight().write(out);
      getBottom().write(out);
      getInsideH().write(out);
      getInsideV().write(out);
    }
    out.print("</" + getTagName() + ">");
  }
  
  public boolean isEmpty() {
    return getLeft().isEmpty() && getTop().isEmpty() && getRight().isEmpty()
        && getBottom().isEmpty() && getInsideH().isEmpty() && getInsideV().isEmpty();
  }
  
  public void setLineStyle(LineStyle lineStyle) {
    getLeft().setLineStyle(lineStyle);
    getTop().setLineStyle(lineStyle);
    getRight().setLineStyle(lineStyle);
    getBottom().setLineStyle(lineStyle);
    getInsideH().setLineStyle(lineStyle);
    getInsideV().setLineStyle(lineStyle);
  }
  
  public void setSize(int size) {
    getLeft().setSize(size);
    getTop().setSize(size);
    getRight().setSize(size);
    getBottom().setSize(size);
    getInsideH().setSize(size);
    getInsideV().setSize(size);
  }
  
  public Border getLeft() {
    return left;
  }
  
  public Border getRight() {
    return right;
  }
  
  public Border getTop() {
    return top;
  }
  
  public Border getBottom() {
    return bottom;
  }
  
  public Border getInsideH() {
    return insideH;
  }
  
  public Border getInsideV() {
    return insideV;
  }
  
  public String getTagName() {
    return tagName;
  }
}
