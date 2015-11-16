package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

/**
 * Графический объект листа с изображением
 */
class TwoCellAnchorImage extends TwoCellAnchor {
  
  private final int relid; // id уникальной ссылки на графический объект внутри листа
  private final Image image; // графический объект -- изображение
  
  public TwoCellAnchorImage(int relid, Image image, SheetCoord coordFrom, SheetCoord coordTo) {
    
    super(coordFrom, coordTo);
    this.relid = relid;
    this.image = image;
  }
  
  @Override
  int getRelId() {
    return relid;
  }
  
  @Override
  String getType() {
    return image.getType();
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
    buf.append("<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" + getRelId() * 2 + "\" name=\"Рисунок "
        + getRelId()
        + "\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>");
    buf.append("<xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"rId");
    buf.append(getRelId());
    buf.append("\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>");
    buf.append("<xdr:spPr><a:xfrm></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>");
    buf.append("</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>");
    
    os.print(buf.toString());
  }
}