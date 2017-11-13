package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

public class Margins implements XmlWriter {
  private final String tagName;
  
  Margins(String tagName) {
    this.tagName = tagName;
  }
  
  public static class Margin implements XmlWriter {
    private final Side side;
    
    public static enum Side {
      LEFT("w:left"), TOP("w:top"), RIGHT("w:right"), BOTTOM("w:bottom");
      
      private final String code;
      
      private Side(String code) {
        this.code = code;
      }
      
      public String getCode() {
        return code;
      }
    }
    
    private int width = 0;
    
    public boolean isEmpty() {
      return width == 0;
    }
    
    public Margin(Side side) {
      this.side = side;
      assert side != null;
    }
    
    public void setWidth(int width) {
      this.width = width;
    }
    
    public int getWidth() {
      return width;
    }
    
    @Override
    public void write(PrintStream out) {
      out.print("<" + getSide().getCode() + " w:w=\"" + getWidth() + "\" w:type=\"dxa\" />");
    }
    
    public Side getSide() {
      return side;
    }
  }
  
  private final Margin left = new Margin(Margin.Side.LEFT);
  private final Margin top = new Margin(Margin.Side.TOP);
  private final Margin right = new Margin(Margin.Side.RIGHT);
  private final Margin bottom = new Margin(Margin.Side.BOTTOM);
  
  @Override
  public void write(PrintStream out) {
    if (isEmpty()) return;
    out.print("<" + tagName + ">");
    getLeft().write(out);
    getTop().write(out);
    getRight().write(out);
    getBottom().write(out);
    out.print("</" + tagName + ">");
  }
  
  public Margin getLeft() {
    return left;
  }
  
  public Margin getTop() {
    return top;
  }
  
  public Margin getRight() {
    return right;
  }
  
  public Margin getBottom() {
    return bottom;
  }
  
  public void setWidth(int width) {
    getLeft().setWidth(width);
    getTop().setWidth(width);
    getRight().setWidth(width);
    getBottom().setWidth(width);
  }
  
  public boolean isEmpty() {
    return getLeft().isEmpty() && getTop().isEmpty() && getRight().isEmpty()
        && getBottom().isEmpty();
  }
}
