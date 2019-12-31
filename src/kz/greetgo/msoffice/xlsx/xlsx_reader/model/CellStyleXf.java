package kz.greetgo.msoffice.xlsx.xlsx_reader.model;

public class CellStyleXf {
  public Integer numFmtId;
  public Integer fontId;
  public Integer fillId;
  public Integer borderId;
  public Integer xfId;

  public VerticalAlign verticalAlign = VerticalAlign.bottom;
  public HorizontalAlign horizontalAlign = HorizontalAlign.left;
}
