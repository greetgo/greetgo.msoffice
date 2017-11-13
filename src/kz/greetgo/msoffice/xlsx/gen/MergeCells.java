package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

class MergeCells {
  MergeCells() {}
  
  private final List<MergeCell> cells = new ArrayList<MergeCell>();
  
  public void print(PrintStream out) {
    if (cells.size() == 0) return;
    out.println("<mergeCells count=\"" + cells.size() + "\">");
    for (MergeCell cell : cells) {
      cell.print(out);
    }
    out.println("</mergeCells>");
  }
  
  public void addMerge(int rowFrom, int colFrom, int rowTo, int colTo) {
    MergeCell x = new MergeCell();
    x.colFrom = colFrom - 1;
    x.colTo = colTo - 1;
    x.rowFrom = rowFrom;
    x.rowTo = rowTo;
    cells.add(x);
  }
}
