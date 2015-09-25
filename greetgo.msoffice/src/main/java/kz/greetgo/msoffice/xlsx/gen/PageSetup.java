package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

public class PageSetup {
  
  private Integer fitToWidth;
  private Integer fitToHeight;
  private Orientation orientation = Orientation.PORTRAIT;
  
  public static enum Orientation {
    PORTRAIT, //
    LANDSCAPE;
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
  
  boolean printHeader(PrintStream out) {
    if (fitToWidth == null && fitToHeight == null) return false;
    
    out.append("<pageSetUpPr fitToPage=\"1\"/>");
    return true;
  }
  
  void print(PrintStream out) {
    
    StringBuffer sb = new StringBuffer();
    
    sb.append("<pageSetup scale=\"74\" ");
    
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