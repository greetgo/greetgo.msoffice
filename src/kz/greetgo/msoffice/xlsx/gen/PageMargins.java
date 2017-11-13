package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

public class PageMargins {
  private double left = 0.7, top = 0.7, right = 0.7, bottom = 0.7;
  private double header = 0.3, footer = 0.3;
  
  public double getLeft() {
    return left;
  }
  
  public void setLeft(double left) {
    this.left = left;
  }
  
  public PageMargins left(double left) {
    setLeft(left);
    return this;
  }
  
  public double getTop() {
    return top;
  }
  
  public void setTop(double top) {
    this.top = top;
  }
  
  public PageMargins top(double top) {
    setTop(top);
    return this;
  }
  
  public double getRight() {
    return right;
  }
  
  public void setRight(double right) {
    this.right = right;
  }
  
  public PageMargins right(double right) {
    setRight(right);
    return this;
  }
  
  public double getBottom() {
    return bottom;
  }
  
  public void setBottom(double bottom) {
    this.bottom = bottom;
  }
  
  public PageMargins bottom(double bottom) {
    setBottom(bottom);
    return this;
  }
  
  public double getFooter() {
    return footer;
  }
  
  public void setFooter(double footer) {
    this.footer = footer;
  }
  
  public PageMargins footer(double footer) {
    setFooter(footer);
    return this;
  }
  
  public double getHeader() {
    return header;
  }
  
  public void setHeader(double header) {
    this.header = header;
  }
  
  public PageMargins header(double header) {
    setHeader(header);
    return this;
  }
  
  public static void main(String[] args) {
    PageMargins x = new PageMargins();
    x.print(System.out);
  }
  
  void print(PrintStream out) {
    out.println("<pageMargins left=\"" + left + "\" right=\"" + right + "\" top=\"" + top + "\" "
        + "bottom=\"" + bottom + "\" header=\"" + header + "\" footer=\"" + footer + "\" />");
  }
}
