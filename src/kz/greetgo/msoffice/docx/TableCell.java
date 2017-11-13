package kz.greetgo.msoffice.docx;

import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

public class TableCell implements XmlWriter {
  
  private final TableCol tableCol;
  private final String ownerPartName;
  private final MSHelper msHelper;
  
  TableCell(TableCol tableCol, String ownerPartName, MSHelper msHelper) {
    this.tableCol = tableCol;
    this.ownerPartName = ownerPartName;
    this.msHelper = msHelper;
    assert tableCol != null;
  }
  
  private VertAlign vertAlign = VertAlign.TOP;
  private Integer width = null;
  private Integer gridSpan;
  
  private ColMergeStatus colMergeStatus = ColMergeStatus.NONE;
  
  private List<FlowElement> elements = new ArrayList<FlowElement>();
  private final Borders borders = new Borders("w:tcBorders");
  private final Margins margins = new Margins("w:tcMar");
  private final Shd shd = new Shd();
  
  @Override
  public void write(PrintStream out) {
    out.print("<w:tc>");
    {
      List<String> options = new ArrayList<String>();
      if (getVertAlign() != null && getVertAlign() != VertAlign.TOP) {
        options.add("<w:vAlign w:val=\"" + getVertAlign().getCode() + "\" />");
      }
      {
        int w = tableCol.getWidth();
        if (getWidth() != null) {
          w = getWidth();
        }
        if (w > 0) {
          options.add("<w:tcW w:w=\"" + w + "\" w:type=\"dxa\" />");
        }
      }
      
      if (getGridSpan() != null) {
        options.add("<w:gridSpan w:val=\"" + getGridSpan() + "\" />");
      }
      
      boolean needTcPr = false;
      {
        if (options.size() > 0) needTcPr = true;
        if (!getBorders().isEmpty()) needTcPr = true;
        if (!getMargins().isEmpty()) needTcPr = true;
        if (colMergeStatus != ColMergeStatus.NONE) needTcPr = true;
      }
      
      if (needTcPr) out.print("<w:tcPr>");
      {
        for (String s : options) {
          out.print(s);
        }
        getBorders().write(out);
        getMargins().write(out);
        getShd().write(out);
        out.print("<w:hideMark />");
        
        {
          switch (colMergeStatus) {
          case RESTART:
            out.print("<w:vMerge w:val=\"restart\" />");
            break;
          
          case CONTINUE:
            out.print("<w:vMerge />");
            break;
          }
        }
      }
      if (needTcPr) out.print("</w:tcPr>");
      for (FlowElement e : elements) {
        e.write(out);
      }
      if (elements.size() == 0) {
        out.print("<w:p />");
      }
    }
    out.print("</w:tc>");
  }
  
  public Para createPara() {
    return createPara(null);
  }
  
  public Para createPara(Para example) {
    Para ret = new Para(ownerPartName, msHelper);
    if (example != null) {
      ret.copyDecorationFrom(example);
    } else {
      Para defaultPara = msHelper.getDefaultPara();
      if (defaultPara != null) {
        ret.copyDecorationFrom(defaultPara);
      }
    }
    elements.add(ret);
    return ret;
  }
  
  public Borders getBorders() {
    return borders;
  }
  
  public void setWidth(Integer width) {
    this.width = width;
  }
  
  public Integer getWidth() {
    return width;
  }
  
  public void setVertAlign(VertAlign vertAlign) {
    this.vertAlign = vertAlign;
  }
  
  public VertAlign getVertAlign() {
    return vertAlign;
  }
  
  public Margins getMargins() {
    return margins;
  }
  
  public Integer getGridSpan() {
    return gridSpan;
  }
  
  public void setGridSpan(Integer gridSpan) {
    this.gridSpan = gridSpan;
  }
  
  public Shd getShd() {
    return shd;
  }
  
  public void mergeRestart() {
    colMergeStatus = ColMergeStatus.RESTART;
  }
  
  public void mergeCotinue() {
    colMergeStatus = ColMergeStatus.CONTINUE;
  }
  
}
