package kz.greetgo.msoffice.docx;

import java.awt.Color;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

public class Para implements FlowElement {
  
  private List<Run> runs = new ArrayList<Run>();
  
  private Align align = Align.LEFT;
  private Integer spacingBefore;
  private Integer spacingAfter;
  private Integer firstLine, left, right;
  private Integer spacingLine;
  
  private final Borders borders = new Borders("w:pBdr");
  
  private final String ownerPartName;
  
  private final MSHelper msHelper;
  
  // for run
  private boolean bold = false;
  private boolean italic = false;
  private Underline underline = Underline.NONE;
  private TextVertAlign textVertAlign = TextVertAlign.NORMAL;
  private Highlight highlight = Highlight.NONE;
  private Color color = null;
  private Integer textSize = null;
  private Font font = null;
  
  public void setFontName(String name) {
    if (name == null) {
      font = null;
      return;
    }
    font = msHelper.getFont(name);
  }
  
  public void applyFor(Run run) {
    run.setBold(isBold());
    run.setItalic(isItalic());
    run.setUnderline(getUnderline());
    run.setTextVertAlign(getTextVertAlign());
    run.setHighlight(getHighlight());
    run.setColor(getColor());
    run.setTextSize(getTextSize());
    run.font = font;
  }
  
  public void copyDecorationFrom(Para p) {
    align = p.align;
    spacingBefore = p.spacingBefore;
    spacingAfter = p.spacingAfter;
    spacingLine = p.spacingLine;
    firstLine = p.firstLine;
    left = p.left;
    right = p.right;
    
    bold = p.bold;
    italic = p.italic;
    underline = p.underline;
    textVertAlign = p.textVertAlign;
    highlight = p.highlight;
    color = p.color;
    font = p.font;
    textSize = p.textSize;
  }
  
  public Para(String ownerPartName, MSHelper msHelper) {
    this.ownerPartName = ownerPartName;
    this.msHelper = msHelper;
  }
  
  @Override
  public void write(PrintStream out) {
    List<String> options = new ArrayList<String>();
    {
      if (getAlign() != null && getAlign() != Align.LEFT) {
        options.add("<w:jc w:val=\"" + getAlign().getCode() + "\" />");
      }
      if (getSpacingBefore() != null || getSpacingAfter() != null || getSpacingLine() != null) {
        StringBuilder sb = new StringBuilder();
        sb.append("<w:spacing ");
        if (getSpacingBefore() != null) {
          sb.append("w:before=\"");
          sb.append(getSpacingBefore());
          sb.append("\" ");
        }
        if (getSpacingAfter() != null) {
          sb.append("w:after=\"");
          sb.append(getSpacingAfter());
          sb.append("\" ");
        }
        if (getSpacingLine() != null) {
          sb.append("w:line=\"");
          sb.append(getSpacingLine());
          sb.append("\" ");
        }
        sb.append("/>");
        options.add(sb.toString());
      }
      if (getFirstLine() != null || getLeft() != null || getRight() != null) {
        StringBuilder sb = new StringBuilder();
        sb.append("<w:ind");
        if (getFirstLine() != null) {
          sb.append(" w:firstLine=\"");
          sb.append(getFirstLine());
          sb.append("\"");
        }
        if (getLeft() != null) {
          sb.append(" w:left=\"");
          sb.append(getLeft());
          sb.append("\"");
        }
        if (getRight() != null) {
          sb.append(" w:right=\"");
          sb.append(getRight());
          sb.append("\"");
        }
        sb.append(" />");
        options.add(sb.toString());
      }
      
    }
    {
      out.print("<w:p>");
      {
        out.print("<w:pPr>");
        getBorders().write(out);
        for (String s : options) {
          out.print(s);
        }
        out.print("</w:pPr>");
      }
      for (Run run : runs) {
        run.write(out);
      }
      out.print("</w:p>");
    }
  }
  
  public void setAlign(Align textAlign) {
    this.align = textAlign;
  }
  
  public Align getAlign() {
    return align;
  }
  
  public void setSpacingBefore(Integer spacingBefore) {
    this.spacingBefore = spacingBefore;
  }
  
  public Integer getSpacingBefore() {
    return spacingBefore;
  }
  
  public void setSpacingAfter(Integer spacingAfter) {
    this.spacingAfter = spacingAfter;
  }
  
  public Integer getSpacingAfter() {
    return spacingAfter;
  }
  
  public void setFirstLine(Integer firstLine) {
    this.firstLine = firstLine;
  }
  
  public Integer getFirstLine() {
    return firstLine;
  }
  
  public Run createRun() {
    Run ret = new Run(ownerPartName, msHelper);
    applyFor(ret);
    runs.add(ret);
    return ret;
  }
  
  public void addPageBreak() {
    createRun().addPageBreak();
  }
  
  public void setBold(boolean bold) {
    this.bold = bold;
  }
  
  public boolean isBold() {
    return bold;
  }
  
  public void setItalic(boolean italic) {
    this.italic = italic;
  }
  
  public boolean isItalic() {
    return italic;
  }
  
  public void setUnderline(Underline underline) {
    this.underline = underline;
  }
  
  public Underline getUnderline() {
    return underline;
  }
  
  public void setTextVertAlign(TextVertAlign textVertAlign) {
    this.textVertAlign = textVertAlign;
  }
  
  public TextVertAlign getTextVertAlign() {
    return textVertAlign;
  }
  
  public void setHighlight(Highlight highlight) {
    this.highlight = highlight;
  }
  
  public Highlight getHighlight() {
    return highlight;
  }
  
  public void setColor(Color color) {
    this.color = color;
  }
  
  public Color getColor() {
    return color;
  }
  
  public void setTextSize(Integer textSize) {
    this.textSize = textSize;
  }
  
  public Integer getTextSize() {
    return textSize;
  }
  
  public Font getFont() {
    return font;
  }
  
  public void setSpacingLine(Integer spacingLine) {
    this.spacingLine = spacingLine;
  }
  
  public Integer getSpacingLine() {
    return spacingLine;
  }
  
  public Borders getBorders() {
    return borders;
  }
  
  public void setRight(Integer right) {
    this.right = right;
  }
  
  public Integer getRight() {
    return right;
  }
  
  public void setLeft(Integer left) {
    this.left = left;
  }
  
  public Integer getLeft() {
    return left;
  }
}
