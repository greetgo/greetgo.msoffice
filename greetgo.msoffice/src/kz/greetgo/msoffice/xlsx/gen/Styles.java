package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

class Styles {
  private final List<Fill> fills = new ArrayList<Fill>();
  
  int fillIndex(Fill fill) {
    if (fill == null) fill = new PatternFill(PatternFillType.none);
    int index = 0;
    for (Fill f : fills) {
      if (f.equals(fill)) return index;
      index++;
    }
    fills.add(fill.copy());
    return index;
  }
  
  private final List<Borders> bordersList = new ArrayList<Borders>();
  
  int bordersIndex(Borders borders) {
    if (borders == null) borders = new Borders();
    int index = 0;
    for (Borders b : bordersList) {
      if (b.equals(borders)) return index;
      index++;
    }
    bordersList.add(new Borders(borders));
    return index;
  }
  
  private final List<Font> fonts = new ArrayList<Font>();
  
  int fontIndex(Font font) {
    if (font == null) font = new Font();
    int index = 0;
    for (Font f : fonts) {
      if (f.equals(font)) return index;
      index++;
    }
    fonts.add(new Font(font));
    return index;
  }
  
  private final List<Style> styles = new ArrayList<Style>();
  
  int styleIndex(Style style) {
    int index = 0;
    for (Style s : styles) {
      if (s.equals(style)) return index;
      index++;
    }
    styles.add(style.copy());
    return index;
  }
  
  public Styles() {
    fills.add(new PatternFill(PatternFillType.none));
    bordersList.add(new Borders());
  }
  
  void print(PrintStream out) {
    indexAll();
    
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.println("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
    NumFmt.printAll(out);
    printFonts(out);
    printFills(out);
    printBorders(out);
    printCellStyleXfs(out);
    printCellXfs(out);
    out.println("</styleSheet>");
  }
  
  private void printCellStyleXfs(PrintStream out) {
    out.println("<cellStyleXfs count=\"1\">");
    out.println("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" />");
    out.println("</cellStyleXfs>");
  }
  
  private void indexAll() {
    for (Style style : styles) {
      fontIndex(style.font);
      fillIndex(style.fill);
      bordersIndex(style.borders);
    }
  }
  
  private void printFonts(PrintStream out) {
    out.println("<fonts count=\"" + fonts.size() + "\">");
    for (Font font : fonts) {
      font.print(out);
    }
    out.println("</fonts>");
  }
  
  private void printFills(PrintStream out) {
    out.println("<fills count=\"" + fills.size() + "\">");
    for (Fill fill : fills) {
      fill.print(out);
    }
    out.println("</fills>");
  }
  
  private void printBorders(PrintStream out) {
    out.println("<borders count=\"" + bordersList.size() + "\">");
    for (Borders borders : bordersList) {
      borders.print(out);
    }
    out.println("</borders>");
  }
  
  private void printCellXfs(PrintStream out) {
    out.println("<cellXfs count=\"" + bordersList.size() + "\">");
    for (Style style : styles) {
      style.print(out);
    }
    out.println("</cellXfs>");
  }
}
