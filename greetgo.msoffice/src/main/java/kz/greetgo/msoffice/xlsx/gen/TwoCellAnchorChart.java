package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

/**
 * Графический объект листа для диаграмм
 */
class TwoCellAnchorChart extends TwoCellAnchor {
  
  private final Chart chart; // диаграмма
  
  public TwoCellAnchorChart(Chart chart, SheetCoord coordFrom, SheetCoord coordTo) {
    
    super(coordFrom, coordTo);
    this.chart = chart;
  }
  
  @Override
  int getRelId() {
    return chart.getRelId();
  }
  
  @Override
  String getType() {
    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart"
        + chart.getFileId() + ".xml";
  }
  
  @Override
  void print(PrintStream os) {
    
    StringBuffer buf = new StringBuffer();
    
    buf.append("<xdr:twoCellAnchor><xdr:from><xdr:col>");
    buf.append(col1 - 1);
    buf.append("</xdr:col><xdr:colOff>");
    buf.append(col1off);
    buf.append("</xdr:colOff><xdr:row>");
    buf.append(row1 - 1);
    buf.append("</xdr:row><xdr:rowOff>");
    buf.append(row1off);
    buf.append("</xdr:rowOff></xdr:from><xdr:to><xdr:col>");
    buf.append(col2);
    buf.append("</xdr:col><xdr:colOff>");
    buf.append(col2off);
    buf.append("</xdr:colOff><xdr:row>");
    buf.append(row2);
    buf.append("</xdr:row><xdr:rowOff>");
    buf.append(row2off);
    buf.append("</xdr:rowOff></xdr:to>");
    buf.append("<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\""
        + chart.getRelId() * 2 + "\" name=\"anchor" + chart.getRelId() + "\"/>");
    buf.append("<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic>");
    buf.append("<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId");
    buf.append(chart.getRelId());
    buf.append("\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>");
    
    os.print(buf.toString());
  }
}