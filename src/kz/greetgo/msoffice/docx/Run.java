package kz.greetgo.msoffice.docx;

import java.awt.Color;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

import kz.greetgo.msoffice.UtilOffice;

public class Run implements XmlWriter {
  private List<RunElement> elements = new ArrayList<RunElement>();
  
  private boolean bold = false;
  private boolean italic = false;
  private Underline underline = Underline.NONE;
  private TextVertAlign textVertAlign = TextVertAlign.NORMAL;
  private Highlight highlight = Highlight.NONE;
  private Color color = null;
  private Integer textSize;
  
  private final String ownerPartName;
  
  private final MSHelper msHelper;
  
  Font font;
  
  Run(String ownerPartName, MSHelper msHelper) {
    this.ownerPartName = ownerPartName;
    this.msHelper = msHelper;
  }
  
  public ImageElement createImage(InputSource inputSource) {
    Relationship r = msHelper.createRelationshipForImage(ownerPartName, inputSource);
    ImageElement ret = new ImageElement(r.getId());
    elements.add(ret);
    return ret;
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<w:r>");
    {
      List<String> options = new ArrayList<String>();
      {
        if (isBold()) {
          options.add("<w:b />");
        }
        if (isItalic()) {
          options.add("<w:i />");
        }
        if (getUnderline() != null && getUnderline() != Underline.NONE) {
          options.add("<w:u w:val=\"" + getUnderline().getCode() + "\" />");
        }
        if (getTextVertAlign() != null && getTextVertAlign() != TextVertAlign.NORMAL) {
          options.add("<w:vertAlign w:val=\"" + getTextVertAlign().getCode() + "\" />");
        }
        if (getHighlight() != null && getHighlight() != Highlight.NONE) {
          options.add("<w:highlight w:val=\"" + getHighlight().getCode() + "\" />");
        }
        if (getColor() != null) {
          options.add("<w:color w:val=\"" + UtilOffice.toHEX(getColor()) + "\" />");
        }
        if (getTextSize() != null) {
          options.add("<w:sz w:val=\"" + getTextSize() + "\" />");
          options.add("<w:szCs w:val=\"" + getTextSize() + "\" />");
        }
      }
      if (options.size() > 0 || getFont() != null) {
        out.print("<w:rPr>");
        for (String s : options) {
          out.print(s);
        }
        if (getFont() != null) {
          getFont().write(out);
        }
        out.print("</w:rPr>");
      }
    }
    for (RunElement e : elements) {
      e.write(out);
    }
    out.print("</w:r>");
  }
  
  public void addText(String text) {
    elements.add(new RunText(text));
  }
  
  public void addTab() {
    elements.add(new RunTab());
  }
  
  public void addBr() {
    elements.add(new RunBr());
  }
  
  public void addPageBreak() {
    elements.add(new PageBreak());
  }
  
  public void setBold(boolean bold) {
    this.bold = bold;
  }
  
  public boolean isBold() {
    return bold;
  }
  
  public void setUnderline(Underline underline) {
    this.underline = underline;
  }
  
  public Underline getUnderline() {
    return underline;
  }
  
  public void setItalic(boolean italic) {
    this.italic = italic;
  }
  
  public boolean isItalic() {
    return italic;
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
  
  public void setFontName(String name) {
    if (name == null) {
      font = null;
      return;
    }
    font = msHelper.getFont(name);
  }
  
  public Font getFont() {
    return font;
  }
  
  private FieldAppender fieldAppender = null;
  
  public FieldAppender field() {
    if (fieldAppender == null) fieldAppender = new FieldAppender(elements);
    return fieldAppender;
  }
}
