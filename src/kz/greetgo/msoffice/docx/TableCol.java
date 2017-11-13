package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

public class TableCol implements XmlWriter {
  private int width = 0;
  private final int colIndex;
  
  public TableCol(int colIndex) {
    this.colIndex = colIndex;
  }
  
  public void setWidth(int width) {
    this.width = width;
  }
  
  public int getWidth() {
    return width;
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<w:gridCol w:w=\"" + getWidth() + "\" />");
  }
  
  public int getColIndex() {
    return colIndex;
  }
}
