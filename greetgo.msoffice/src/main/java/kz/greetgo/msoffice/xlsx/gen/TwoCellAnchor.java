package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

/**
 * Графический объект листа, связанный с двумя ячейками и перемещяющийся при их перемещении
 * (например при изменении ширины столбцов)
 */
class TwoCellAnchor {
  
  // верхний левый угол
  private final int col1; // колонка
  private final int row1; // строка
  private int col1off; // отступ от начала колонки
  private int row1off; // отступ от начала строки
  
  // правый нижний угол
  private final int col2; // колонка
  private final int row2; // строка
  private int col2off; // отступ от начала колонки
  private int row2off; // отступ от начала строки
  
  private final Chart chart; // график
  
  public TwoCellAnchor(Chart chart, int columnFrom, int rowFrom, int columnTo, int rowTo) {
    
    this.chart = chart;
    
    col1 = columnFrom;
    row1 = rowFrom;
    col2 = columnTo;
    row2 = rowTo;
  }
  
  public TwoCellAnchor(Chart chart, int columnFrom, int rowFrom, int columnTo, int rowTo,
      int columnFromOff, int rowFromOff, int columnToOff, int rowToOff) {
    
    this(chart, columnFrom, rowFrom, columnTo, rowTo);
    
    col1off = columnFromOff;
    row1off = rowFromOff;
    col2off = columnToOff;
    row2off = rowToOff;
  }
  
  int getChartId() {
    return chart.getId();
  }
  
  void print(PrintStream os) {
    
    StringBuffer buf = new StringBuffer();
    
    buf.append("<xdr:twoCellAnchor><xdr:from><xdr:col>");
    buf.append(col1);
    buf.append("</xdr:col><xdr:colOff>");
    buf.append(col1off);
    buf.append("</xdr:colOff><xdr:row>");
    buf.append(row1);
    buf.append("</xdr:row><xdr:rowOff>");
    buf.append(row1off);
    buf.append("</xdr:rowOff></xdr:from><xdr:to><xdr:col>");
    buf.append(col2 + 1);
    buf.append("</xdr:col><xdr:colOff>");
    buf.append(col2off);
    buf.append("</xdr:colOff><xdr:row>");
    buf.append(row2);
    buf.append("</xdr:row><xdr:rowOff>");
    buf.append(row2off);
    buf.append("</xdr:rowOff></xdr:to>");
    buf.append("<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\""
        + chart.getId() * 2 + "\" name=\"anchor" + chart.getId() + "\" />");
    buf.append("<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic>");
    buf.append("<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId");
    buf.append(chart.getId());
    buf.append("\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>");
    
    os.print(buf.toString());
  }
}