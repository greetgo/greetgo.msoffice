package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

public class PageSetup {
  
  private Integer fitToWidth;
  private Integer fitToHeight;
  private Orientation orientation = Orientation.PORTRAIT;
  private PaperSize paperSize;
  
  public static enum Orientation {
    PORTRAIT, //
    LANDSCAPE;
  }
  
  public static enum PaperSize {
    
    A4(9), //
    A5(11), //
    A6(70), //
    Legal(5), //
    Executive(7), //
    B5_JIS(13), //
    Letter(-1);
    
    private int code;
    
    private PaperSize(int code) {
      this.code = code;
    }
  }
  
  public void setFitToWidth(Integer fitToWidth) {
    this.fitToWidth = fitToWidth;
  }
  
  public void setFitToHeight(Integer fitToHeight) {
    this.fitToHeight = fitToHeight;
  }
  
  public void setOrientation(Orientation orientation) {
    this.orientation = orientation;
  }
  
  public void setPaperSize(PaperSize paperSize) {
    
    if (paperSize == PaperSize.Letter) {
      this.paperSize = null;
      return;
    }
    
    this.paperSize = paperSize;
  }
  
  boolean printHeader(PrintStream out) {
    if (fitToWidth == null && fitToHeight == null) return false;
    
    out.append("<pageSetUpPr fitToPage=\"1\"/>");
    return true;
  }
  
  void print(PrintStream out) {
    
    StringBuffer sb = new StringBuffer();
    
    sb.append("<pageSetup scale=\"74\" ");
    
    if (paperSize != null) {
      sb.append(" paperSize=\"");
      sb.append(paperSize.code);
      sb.append("\"");
    }
    
    if (fitToWidth != null) {
      sb.append(" fitToWidth=\"");
      sb.append(fitToWidth);
      sb.append("\"");
    }
    
    if (fitToHeight != null) {
      sb.append(" fitToHeight=\"");
      sb.append(fitToHeight);
      sb.append("\"");
    }
    
    sb.append(" orientation=\"");
    sb.append(orientation.name().toLowerCase());
    sb.append("\"");
    
    sb.append("/>");
    
    out.println(sb.toString());
  }
}