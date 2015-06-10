package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

public class Chart {
  
  private final ChartType type;
  private final int id;
  
  private String title = "";
  private String data;
  private String titles;
  private String data2;
  private String titles2;
  private int expl;
  private Align alignl = Align.bottom;
  private Align alignOY = Align.left;
  private int rotate = 30;
  private int wline = 18440;
  private int wline2 = 18440;
  private Color cink = Color.black();
  private Color cink2 = Color.black();
  private Color cpaper = Color.white();
  
  Chart(ChartType type, int id) {
    
    this.type = type;
    this.id = id;
  }
  
  /** Название диаграммы. */
  public void setTitle(String title) {
    
    if (title == null) title = "";
    this.title = title;
  }
  
  /** Зазор между сегментами диаграммы (круговой). */
  public void setExplosion(int explosion) {
    if (explosion > 0) expl = explosion;
  }
  
  /** Положение списка подписей к данным (легенды). */
  public void setAlignmentLegend(Align align) {
    alignl = align;
  }
  
  /** Положение подписей к оси Y (график): слева / справа. */
  public void setAlignmentOY(Align align) {
    alignOY = align;
  }
  
  /** Поворот графика по горизонтали. */
  public void setRotation(int rotationAngle) {
    rotate = rotationAngle;
  }
  
  /** Цвет линий диаграммы. */
  public void setColorLine(Color lineColor) {
    cink = lineColor;
  }
  
  /** Цвет дополнительных линий диаграммы (для графика с двумя функциями). */
  public void setColorLine2(Color lineColor) {
    cink2 = lineColor;
  }
  
  /** Ширина линий диаграммы (1-30). */
  public void setWidthLine(int widthLine) {
    wline = widthLine * 1000;
  }
  
  /** Ширина дополнительных линий диаграммы (1-30) (для графика с двумя функциями). */
  public void setWidthLine2(int widthLine) {
    wline2 = widthLine * 1000;
  }
  
  /** Цвет фона диаграммы. */
  public void setColorBackground(Color bgColor) {
    cpaper = bgColor;
  }
  
  /**
   * Определение строки или столбца с числовыми данными.
   * 
   * @param data
   *          формула строки в формате
   *          <b>ИмяЛиста</b>!$<b>БукваСтолбца1</b>$<b>НомерСтроки</b>:$<b>БукваСтолбца2
   *          </b>$<b>НомерСтроки</b>
   * 
   *          или формула столбца в формате <b>ИмяЛиста</b>!$<b>БукваСтолбца</b>$<b>
   *          НомерСтроки1</b>:$<b>БукваСтолбца</b>$<b>НомерСтроки2</b>
   */
  public void setData(String data) {
    this.data = data;
  }
  
  public void setData2(String data) {
    this.data2 = data;
  }
  
  /**
   * Определение столбца с числовыми данными, удобна для динамического определения.
   * 
   * @param dataSheet
   *          лист, содержащий столбец с числовыми данными
   * @param dataCol
   *          буква столбца (A, B, C)
   * @param dataRow1
   *          номер первой строки в столбце (1, 2, 3)
   * @param dataRow2
   *          номер последней строки в столбце (4, 5, 6)
   */
  public void setData(Sheet dataSheet, String dataCol, int dataRow1, int dataRow2) {
    data = setData_(dataSheet, dataCol, dataRow1, dataRow2);
  }
  
  public void setData2(Sheet dataSheet, String dataCol, int dataRow1, int dataRow2) {
    data2 = setData_(dataSheet, dataCol, dataRow1, dataRow2);
  }
  
  private String setData_(Sheet dataSheet, String dataCol, int dataRow1, int dataRow2) {
    
    StringBuffer buf = new StringBuffer();
    
    buf.append(dataSheet.getDisplayName());
    buf.append("!$");
    buf.append(dataCol);
    buf.append("$");
    buf.append(dataRow1);
    buf.append(":$");
    buf.append(dataCol);
    buf.append("$");
    buf.append(dataRow2);
    
    return buf.toString();
  }
  
  /**
   * Определение строки с числовыми данными, удобна для динамического определения.
   * 
   * @param dataSheet
   *          лист, содержащий строку с числовыми данными
   * @param dataCol1
   *          буква первого столбца в строке (A, B, C)
   * @param dataCol2
   *          буква первого столбца в строке (A, B, C)
   * @param dataRow
   *          номер строки (1, 2, 3)
   */
  public void setData(Sheet dataSheet, String dataCol1, String dataCol2, int dataRow) {
    data = setData_(dataSheet, dataCol1, dataCol2, dataRow);
  }
  
  public void setData2(Sheet dataSheet, String dataCol1, String dataCol2, int dataRow) {
    data2 = setData_(dataSheet, dataCol1, dataCol2, dataRow);
  }
  
  public String setData_(Sheet dataSheet, String dataCol1, String dataCol2, int dataRow) {
    
    StringBuffer buf = new StringBuffer();
    
    buf.append(dataSheet.getDisplayName());
    buf.append("!$");
    buf.append(dataCol1);
    buf.append("$");
    buf.append(dataRow);
    buf.append(":$");
    buf.append(dataCol2);
    buf.append("$");
    buf.append(dataRow);
    
    return buf.toString();
  }
  
  /**
   * Определение строки или столбца с подписями.
   * 
   * @param titles
   *          формула строки в формате
   *          <b>ИмяЛиста</b>!$<b>БукваСтолбца1</b>$<b>НомерСтроки</b>:$<b>БукваСтолбца2
   *          </b>$<b>НомерСтроки</b>
   * 
   *          или формула столбца в формате <b>ИмяЛиста</b>!$<b>БукваСтолбца</b>$<b>
   *          НомерСтроки1</b>:$<b>БукваСтолбца</b>$<b>НомерСтроки2</b>
   */
  public void setDataTitles(String titles) {
    this.titles = titles;
  }
  
  public void setDataTitles2(String titles) {
    this.titles2 = titles;
  }
  
  /**
   * Определение столбца с подписями, удобна для динамического определения.
   * 
   * @param dataSheet
   *          лист, содержащий столбец с числовыми данными
   * @param dataCol
   *          буква столбца (A, B, C)
   * @param dataRow1
   *          номер первой строки в столбце (1, 2, 3)
   * @param dataRow2
   *          номер последней строки в столбце (4, 5, 6)
   */
  public void setDataTitles(Sheet dataSheet, String dataCol, int dataRow1, int dataRow2) {
    titles = setDataTitles_(dataSheet, dataCol, dataRow1, dataRow2);
  }
  
  public void setDataTitles2(Sheet dataSheet, String dataCol, int dataRow1, int dataRow2) {
    titles2 = setDataTitles_(dataSheet, dataCol, dataRow1, dataRow2);
  }
  
  public String setDataTitles_(Sheet dataSheet, String dataCol, int dataRow1, int dataRow2) {
    
    StringBuffer buf = new StringBuffer();
    
    buf.append(dataSheet.getDisplayName());
    buf.append("!$");
    buf.append(dataCol);
    buf.append("$");
    buf.append(dataRow1);
    buf.append(":$");
    buf.append(dataCol);
    buf.append("$");
    buf.append(dataRow2);
    
    return buf.toString();
  }
  
  /**
   * Определение строки с подписями, удобна для динамического определения.
   * 
   * @param dataSheet
   *          лист, содержащий строку с числовыми данными
   * @param dataCol1
   *          буква первого столбца в строке (A, B, C)
   * @param dataCol2
   *          буква первого столбца в строке (A, B, C)
   * @param dataRow
   *          номер строки (1, 2, 3)
   */
  public void setDataTitles(Sheet dataSheet, String dataCol1, String dataCol2, int dataRow) {
    data = setDataTitles_(dataSheet, dataCol1, dataCol2, dataRow);
  }
  
  public void setDataTitles2(Sheet dataSheet, String dataCol1, String dataCol2, int dataRow) {
    data2 = setDataTitles_(dataSheet, dataCol1, dataCol2, dataRow);
  }
  
  public String setDataTitles_(Sheet dataSheet, String dataCol1, String dataCol2, int dataRow) {
    
    StringBuffer buf = new StringBuffer();
    
    buf.append(dataSheet.getDisplayName());
    buf.append("!$");
    buf.append(dataCol1);
    buf.append("$");
    buf.append(dataRow);
    buf.append(":$");
    buf.append(dataCol2);
    buf.append("$");
    buf.append(dataRow);
    
    return buf.toString();
  }
  
  int getId() {
    return id;
  }
  
  void print(PrintStream os) {
    
    StringBuffer buf = new StringBuffer();
    
    buf.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    buf.append("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
    buf.append("<c:lang val=\"ru-RU\"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr sz=\"1300\"><a:latin typeface=\"Arial\"/></a:rPr><a:t>");
    buf.append(title);
    buf.append("</a:t></a:r></a:p></c:rich></c:tx><c:layout/></c:title><c:view3D><c:rotX val=\"");
    buf.append(rotate);
    buf.append("\"/><c:perspective val=\"30\"/></c:view3D><c:plotArea><c:layout/><c:");
    buf.append(type.tag);
    buf.append(">");
    
    if (type == ChartType.CIRCLE_DIAGRAM) {
      buf.append("<c:varyColors val=\"1\"/>");
    }
    
    appendSer(buf, data, titles, cink, wline);
    
    if (data2 != null && !data2.trim().equals("")) appendSer(buf, data2, titles2, cink2, wline2);
    
    if (type == ChartType.LINE_CHART) {
      buf.append("<c:hiLowLines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:hiLowLines><c:marker val=\"0\"/><c:axId val=\"12344193\"/><c:axId val=\"63668120\"/>");
    }
    
    buf.append("</c:");
    buf.append(type.tag);
    buf.append(">");
    
    if (type == ChartType.LINE_CHART) {
      buf.append("<c:catAx><c:axId val=\"12344193\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:spPr><a:ln w=\"9360\"><a:solidFill><a:srgbClr val=\"878787\"/></a:solidFill><a:round/></a:ln></c:spPr><c:crossAx val=\"63668120\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx>");
      buf.append("<c:valAx><c:axId val=\"63668120\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"r\"/><c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"");
      buf.append(alignOY.getTagOY());
      buf.append("\"/><c:spPr><a:ln><a:solidFill><a:srgbClr val=\"b3b3b3\"/></a:solidFill></a:ln></c:spPr><c:crossAx val=\"12344193\"/><c:crosses val=\"max\"/></c:valAx>");
      buf.append("<c:spPr><a:solidFill><a:srgbClr val=\"");
      buf.append(cpaper.strRGB());
      buf.append("\"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>");
    }
    
    buf.append("</c:plotArea>");
    
    if (type == ChartType.CIRCLE_DIAGRAM) {
      buf.append("<c:legend><c:legendPos val=\"");
      buf.append(alignl.getTag());
      buf.append("\"/><c:layout/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr rtl=\"0\"><a:defRPr/></a:pPr><a:endParaRPr lang=\"ru-RU\"/></a:p></c:txPr></c:legend>");
    }
    
    buf.append("</c:chart>");
    buf.append("<c:printSettings><c:headerFooter/><c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\"/><c:pageSetup/></c:printSettings>");
    buf.append("</c:chartSpace>");
    
    os.print(buf.toString());
  }
  
  private void appendSer(StringBuffer buf, String data, String titles, Color cink, int wline) {
    
    buf.append("<c:ser><c:explosion val=\"");
    buf.append(expl);
    buf.append("\"/>");
    
    if (type == ChartType.LINE_CHART) {
      buf.append("<c:spPr><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill><a:ln w=\"");
      buf.append(wline);
      buf.append("\"><a:solidFill><a:srgbClr val=\"");
      buf.append(cink.strRGB());
      buf.append("\"/></a:solidFill><a:round/></a:ln></c:spPr>");
      //  buf.append("<c:dLbls><c:dLbl><c:idx val=\"0\"/><c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/></c:dLbl><c:dLbl><c:idx val=\"1\"/><c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/></c:dLbl><c:dLbl><c:idx val=\"2\"/><c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/></c:dLbl><c:dLbl><c:idx val=\"3\"/><c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/></c:dLbl><c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/></c:dLbls>");
    }
    
    if (type == ChartType.CIRCLE_DIAGRAM) {
      buf.append("<c:dLbls><c:showPercent val=\"1\"/><c:showLeaderLines val=\"1\"/></c:dLbls>");
    }
    
    buf.append("<c:cat><c:strRef><c:f>");
    buf.append(titles);
    buf.append("</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>");
    buf.append(data);
    buf.append("</c:f></c:numRef></c:val></c:ser>");
  }
}