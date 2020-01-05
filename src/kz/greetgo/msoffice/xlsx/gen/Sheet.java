package kz.greetgo.msoffice.xlsx.gen;

import kz.greetgo.msoffice.ImageType;
import kz.greetgo.msoffice.util.UtilOffice;
import kz.greetgo.msoffice.xlsx.parse.SharedStrings;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.PrintStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

public class Sheet {

  private final Xlsx parent;

  private final Style style;
  private final String name;
  private final String workDir;
  private final SharedStrings strs;
  private double defaultRowHeight = 15;
  private boolean summaryBelow = true;
  private boolean summaryRight = true;

  private int curRow = 0;
  private boolean inRow = false;

  private final Map<Integer, Double> widths = new HashMap<Integer, Double>();
  private final Set<Integer> colhiddens = new HashSet<Integer>();
  private final Map<Integer, Double> heights = new HashMap<Integer, Double>();
  final boolean selected;
  private boolean started = false;
  private final boolean finished = false;
  private PrintStream out;
  private final PageMargins pageMargins = new PageMargins();
  private final PageSetup pageSetup = new PageSetup();
  private String displayName;
  private final MergeCells mergeCells = new MergeCells();

  private int lastColumn;

  private final List<TwoCellAnchor> drawing = new ArrayList<>(); // графические объекты листа
  private Integer drawingFileId; // номер файла графики для листа
  private int drawingRelIdLast = 0; // счётчик ссылок графических объектов внутри файла графики для листа
  private Map<Image, Integer> imagesrels = new HashMap<>(); // картинка -> ссылка графического объекта в листе   

  public PageMargins pageMargins() {
    return pageMargins;
  }

  public PageSetup pageSetup() {
    return pageSetup;
  }

  /**
   * Установка положения страницы
   *
   * @param orientation положение страницы: горизонтальное Landscape, вертикальное Portrait
   */
  public void setPageOrientation(PageSetup.Orientation orientation) {
    pageSetup.setOrientation(orientation);
  }

  /**
   * Установка формата бумаги
   *
   * @param paperSize формат бумаги: A4, A5, Letter, etc
   */
  public void setPageSize(PageSetup.PaperSize paperSize) {
    pageSetup.setPaperSize(paperSize);
  }

  /**
   * При печати подгонять масштаб под ширину страницы
   */
  public void setScaleByWidth() {
    pageSetup.setFitToHeight(0);
    pageSetup.setFitToWidth(1);
  }

  public String getDisplayName() {
    return displayName;
  }

  public void setDisplayName(String displayName) {
    this.displayName = displayName;
  }

  public double getDefaultRowHeight() {
    return defaultRowHeight;
  }

  public void setDefaultRowHeight(double defaultRowHeight) {
    this.defaultRowHeight = defaultRowHeight;
  }

  public Sheet cellFormula(int col, String formulaValue, NumFmt numFmt) {
    init();
    setLastColumn(col);
    style.numFmt = numFmt;

    if (formulaValue == null) {
      int styleIndex = style.index();
      String pos = UtilOffice.toTablePosition(curRow, col - 1);
      out.printf("<c r=\"%s\" s=\"%d\" t=\"e\"><f></f></c>", pos, styleIndex);
      out.println();
    } else {
      int styleIndex = style.index();
      String pos = UtilOffice.toTablePosition(curRow, col - 1);
      out.printf("<c r=\"%s\" s=\"%d\" t=\"e\"><f>%s</f></c>", pos, styleIndex, formulaValue);
      out.println();
    }
    return this;
  }

  public Sheet cellFormula(int col, String formulaValue) {
    init();
    setLastColumn(col);
    style.clearNumFmt();

    if (formulaValue == null) {
      int styleIndex = style.index();
      String pos = UtilOffice.toTablePosition(curRow, col - 1);
      out.printf("<c r=\"%s\" s=\"%d\" t=\"e\"><f></f></c>", pos, styleIndex);
      out.println();
    } else {
      int styleIndex = style.index();
      String pos = UtilOffice.toTablePosition(curRow, col - 1);
      out.printf("<c r=\"%s\" s=\"%d\" t=\"e\"><f>%s</f></c>", pos, styleIndex, formulaValue);
      out.println();
    }
    return this;
  }

  public Sheet cellInlineStr(int col, String strValue) {
    init();
    setLastColumn(col);
    style.clearNumFmt();

    if (strValue == null) {
      int styleIndex = style.index();
      String pos = UtilOffice.toTablePosition(curRow, col - 1);
      out.printf("<c r=\"%s\" s=\"%d\" t=\"inlineStr\"></c>", pos, styleIndex);
      out.println();
    } else {
      int styleIndex = style.index();
      String pos = UtilOffice.toTablePosition(curRow, col - 1);
      out.printf("<c r=\"%s\" s=\"%d\" t=\"inlineStr\"><is><t>" + strValue + "</t></is></c>", pos,
        styleIndex);
      out.println();
    }
    return this;
  }

  public Sheet cellStr(int col, String strValue) {
    init();
    setLastColumn(col);
    style.clearNumFmt();

    if (strValue == null) {
      int styleIndex = style.index();
      String pos = UtilOffice.toTablePosition(curRow, col - 1);
      out.printf("<c r=\"%s\" s=\"%d\" t=\"s\"></c>", pos, styleIndex);
      out.println();
    } else {
      int styleIndex = style.index();
      int strIndex = strs.index(strValue);
      String pos = UtilOffice.toTablePosition(curRow, col - 1);
      out.printf("<c r=\"%s\" s=\"%d\" t=\"s\"><v>%d</v></c>", pos, styleIndex, strIndex);
      out.println();
    }
    return this;
  }

  private void setLastColumn(int col) {
    if (col <= 0) throw new IllegalArgumentException("must be col > 0, but col = " + col);
    if (col <= lastColumn) throw new IllegalArgumentException("must be col > " + lastColumn
      + ", but col = " + col);
    lastColumn = col;
  }

  public Sheet cellInt(int col, int intValue) {
    init();
    setLastColumn(col);
    style.clearNumFmt();

    int styleIndex = style.index();
    String pos = UtilOffice.toTablePosition(curRow, col - 1);
    out.printf("<c r=\"%s\" s=\"%d\"><v>%d</v></c>", pos, styleIndex, intValue);
    out.println();
    return this;
  }

  public Sheet cellDouble(int col, double doubleValue, NumFmt numFmt) {
    init();
    setLastColumn(col);
    style.numFmt = numFmt;

    int styleIndex = style.index();
    String pos = UtilOffice.toTablePosition(curRow, col - 1);
    out.printf("<c r=\"%s\" s=\"%d\"><v>" + doubleValue + "</v></c>", pos, styleIndex);
    out.println();
    return this;
  }

  public Sheet cellDouble(int col, double doubleValue) {
    init();
    setLastColumn(col);
    style.clearNumFmt();

    int styleIndex = style.index();
    String pos = UtilOffice.toTablePosition(curRow, col - 1);
    out.printf("<c r=\"%s\" s=\"%d\"><v>" + doubleValue + "</v></c>", pos, styleIndex);
    out.println();
    return this;
  }

  private Sheet cellDate(int col, Date date, NumFmt numFmt) {
    init();
    setLastColumn(col);
    style.numFmt = numFmt;

    int styleIndex = style.index();
    String pos = UtilOffice.toTablePosition(curRow, col - 1);
    out.printf("<c r=\"%s\" s=\"%d\"><v>%s</v></c>", pos, styleIndex,
      UtilOffice.toExcelDateTime(date));
    out.println();
    return this;
  }

  public Sheet cellYMD(int col, Date date) {
    return cellDate(col, date, NumFmt.YMD);
  }

  public Sheet cellYMD_HMS(int col, Date date) {
    return cellDate(col, date, NumFmt.YMD_HMS);
  }

  public Sheet cellDMY(int col, Date date) {
    return cellDate(col, date, NumFmt.DMY);
  }

  public void skipRows(int rowCount) {
    if (rowCount <= 0) {
      throw new IllegalArgumentException("rowCount must be > 0");
    }
    init();
    innerFinishRow();
    curRow += rowCount;
  }

  public void innerFinishRow() {
    if (!inRow) return;
    out.println("</row>");
    inRow = false;
  }

  public void skipRow() {
    init();
    skipRows(1);
  }

  public class Row {
    private Double height = null;
    private boolean collapsed = false;
    private boolean hidden = false;
    private int outlineLevel = 0;

    private Row() {}

    public Row height(double height) {
      this.height = height;
      return this;
    }

    public Row collapsed() {
      collapsed = true;
      return this;
    }

    public Row hidden() {
      hidden = true;
      return this;
    }

    public Row hidden(boolean hidden) {
      this.hidden = hidden;
      return this;
    }

    public Row collapsed(boolean collapsed) {
      this.collapsed = collapsed;
      return this;
    }

    public Row outline() {
      outlineLevel = 1;
      return this;
    }

    public Row outline(boolean hidden) {
      outlineLevel = 1;
      this.hidden = hidden;
      return this;
    }

    public Row hiddenOutline() {
      outlineLevel = 1;
      hidden = true;
      return this;
    }

    public Row outlineLevel(int outlineLevel) {
      if (outlineLevel < 1) throw new IllegalArgumentException("outlineLevel must be > 0");
      this.outlineLevel = outlineLevel;
      return this;
    }

    public Row outlineLevel(int outlineLevel, boolean hidden) {
      if (outlineLevel < 1) throw new IllegalArgumentException("outlineLevel must be > 0");
      this.hidden = hidden;
      this.outlineLevel = outlineLevel;
      return this;
    }

    public Row start() {
      init();
      innerFinishRow();

      lastColumn = 0;

      StringBuilder after = new StringBuilder();
      out.printf("<row r=\"%d\"", ++curRow);
      if (height != null) {
        after.append(" customHeight=\"1\"");
        out.printf(" ht=\"%s\"", height.toString());
        heights.put(Integer.valueOf(curRow), height);
      }
      if (hidden) {
        out.print(" hidden=\"1\"");
      }
      if (outlineLevel > 0) {
        out.printf(" outlineLevel=\"%d\"", outlineLevel);
      }
      if (collapsed) {
        out.print(" collapsed=\"1\"");
      }
      out.print(after.toString());
      out.println('>');

      inRow = true;

      return this;
    }

    public void finish() {
      innerFinishRow();
    }
  }

  public Row row() {
    return new Row();
  }

  public void cleanWidths() {
    checkNotStarted();
    widths.clear();
  }

  public void cleanHiddens() {
    checkNotStarted();
    colhiddens.clear();
  }

  private void checkNotStarted() {
    if (started) {
      throw new IllegalStateException("Cannot change widths when data filling has begun");
    }
  }

  public void setWidth(int col, double width) {
    if (col <= 0) throw new IllegalArgumentException("must be: col > 0, but col = " + col);
    checkNotStarted();
    widths.put(col, width);
  }

  /*
   *
   */

  /**
   * Скрыть колонку
   *
   * @param col индекс скрываемой колонки
   */
  public void setHidden(int col) {
    if (col <= 0) throw new IllegalArgumentException("must be: col > 0, but col = " + col);
    checkNotStarted();
    colhiddens.add(Integer.valueOf(col));
  }

  public Style style() {
    return style;
  }

  Sheet(Xlsx owner, Styles styles, int sheetNo, String workDir, SharedStrings strs, boolean selected) {

    parent = owner;
    this.name = "sheet" + sheetNo;
    this.displayName = "Лист" + sheetNo;
    this.workDir = workDir;
    this.strs = strs;
    this.selected = selected;

    style = new Style(styles);
  }

  public String name() {
    return name;
  }

  private final void init() {
    if (started) return;
    try {
      initEx();
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException) e;
      }
      throw new RuntimeException(e);
    }
  }

  private void initEx() throws Exception {
    if (started) {
      throw new IllegalArgumentException("Cannot start already started sheet");
    }
    if (finished) {
      throw new IllegalArgumentException("Cannot start already finished sheet");
    }
    {
      File dir = new File(workDir + "/xl/worksheets");
      dir.mkdirs();
      out = new PrintStream(new FileOutputStream(dir.getPath() + "/" + name + ".xml"), true,
        "UTF-8");
      printHeader();
    }
    started = true;
  }

  private void printHeader() {
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.println("<worksheet xmlns=\"http://schemas.openxmlformats.org/"
      + "spreadsheetml/2006/main\""
      + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
    printSheetPr();
    printSelected();
    out.println("<sheetFormatPr defaultRowHeight=\"" + defaultRowHeight + "\" />");
    printWidths();
    out.println("<sheetData>");
  }

  private void printSheetPr() {

    out.println("<sheetPr>");

    if (!pageSetup.printHeader(out)) {
      out.println("<outlinePr summaryBelow=\"" + (summaryBelow ? 1 : 0) + "\" summaryRight=\""
        + (summaryRight ? 1 : 0) + "\" />");
    }

    out.println("</sheetPr>");
  }

  private void printSelected() {
    if (selected) {
      out.println("<sheetViews>");
      out.println("<sheetView tabSelected=\"1\" workbookViewId=\"0\">");
      out.println("<selection activeCell=\"A1\" sqref=\"A1\" />");
      out.println("</sheetView>");
      out.println("</sheetViews>");
    } else {
      out.println("<sheetViews>");
      out.println("<sheetView workbookViewId=\"0\" />");
      out.println("</sheetViews>");
    }
  }

  private void printWidths() {

    if (widths.size() == 0 && colhiddens.size() == 0) return;

    Set<Integer> keys = new TreeSet<>();
    keys.addAll(widths.keySet());
    keys.addAll(colhiddens);

    out.println("<cols>");
    for (Integer key : keys) {
      out.printf("<col min=\"%d\" max=\"%d\" ", key, key);

      if (widths.get(key) != null) out.printf("width=\"%s\" ", widths.get(key).toString());

      if (colhiddens.contains(key)) out.printf("hidden=\"1\" ");

      out.printf("customWidth=\"1\" />");
    }
    out.println("</cols>");
  }

  void finish() {
    out.println("</sheetData>");
    mergeCells.print(out);
    pageMargins.print(out);
    pageSetup.print(out);
    printDrawings();
    out.println("</worksheet>");
    out.close();
    out = null;
  }

  public void addMergeInRow(int colFrom, int colTo) {
    mergeCells.addMerge(curRow, colFrom, curRow, colTo);
  }

  public void addMerge(int rowFrom, int colFrom, int rowTo, int colTo) {
    mergeCells.addMerge(rowFrom, colFrom, rowTo, colTo);
  }

  public void addMergeRel(int rowFrom, int colFrom, int rowTo, int colTo) {
    mergeCells.addMerge(curRow + rowFrom, colFrom, curRow + rowTo, colTo);
  }

  public void addMergeCount(int colFrom, int colTo, int rowsCount) {
    if (rowsCount > 0) {
      mergeCells.addMerge(curRow, colFrom, curRow + rowsCount, colTo);
    } else {
      mergeCells.addMerge(curRow + rowsCount, colFrom, curRow, colTo);
    }
  }

  public boolean isSummaryBelow() {
    return summaryBelow;
  }

  public void setSummaryBelow(boolean summaryBelow) {
    this.summaryBelow = summaryBelow;
  }

  public boolean isSummaryRight() {
    return summaryRight;
  }

  public void setSummaryRight(boolean summaryRight) {
    this.summaryRight = summaryRight;
  }

  public Chart addChart(ChartType type, SheetCoord coordFrom, SheetCoord coordTo) {

    Chart chart = parent.newChart(type, ++drawingRelIdLast);
    addChart(chart, coordFrom, coordTo);

    return chart;
  }

  public Chart addChart(ChartType type, String cellFrom, String cellTo) {

    Chart chart = parent.newChart(type, ++drawingRelIdLast);
    addChart(chart, new SheetCoord(cellFrom), new SheetCoord(cellTo));

    return chart;
  }

  @Deprecated
  public Chart addChart(ChartType type, int colFrom, int rowFrom, int colTo, int rowTo) {

    Chart chart = parent.newChart(type, ++drawingRelIdLast);
    addChart(chart, new SheetCoord(colFrom, rowFrom), new SheetCoord(colTo, rowTo));

    return chart;
  }

  @Deprecated
  public Chart addChart(ChartType type, String colFrom, int rowFrom, String colTo, int rowTo) {
    return addChart(type, new SheetCoord(colFrom, rowFrom), new SheetCoord(colTo, rowTo));
  }

  public void addChart(Chart chart, SheetCoord coordFrom, SheetCoord coordTo) {

    setDrawingId();
    TwoCellAnchor anchor = new TwoCellAnchorChart(chart, coordFrom, coordTo);
    drawing.add(anchor);
  }

  public void addChart(Chart chart, String cellFrom, String cellTo) {

    setDrawingId();
    TwoCellAnchor anchor = new TwoCellAnchorChart(chart, new SheetCoord(cellFrom), new SheetCoord(
      cellTo));
    drawing.add(anchor);
  }

  @Deprecated
  public void addChart(Chart chart, int colFrom, int rowFrom, int colTo, int rowTo) {

    setDrawingId();
    TwoCellAnchor anchor = new TwoCellAnchorChart(chart, new SheetCoord(colFrom, rowFrom),
      new SheetCoord(colTo, rowTo));
    drawing.add(anchor);
  }

  @Deprecated
  public void addChart(Chart chart, String colFrom, int rowFrom, String colTo, int rowTo) {
    addChart(chart, new SheetCoord(colFrom, rowFrom), new SheetCoord(colTo, rowTo));
  }

  public Image addImage(byte[] img, ImageType type, SheetCoord coordFrom, SheetCoord coordTo) {
    InputStream is = getISFromBytes(img, type);
    return addImage(is, type, coordFrom, coordTo);
  }

  public Image addImage(byte[] img, ImageType type, SheetCoord coordFrom, int percentOfSize) {
    InputStream is = getISFromBytes(img, type);
    return addImage(is, type, coordFrom, percentOfSize);
  }

  public Image addImage(File file, SheetCoord coordFrom, SheetCoord coordTo) {

    InputStream is = getISFromFile(file);
    ImageType type = getImageTypeFromFile(file);
    return addImage(is, type, coordFrom, coordTo);
  }

  public Image addImage(File file, SheetCoord coordFrom, int percentOfSize) {

    InputStream is = getISFromFile(file);
    ImageType type = getImageTypeFromFile(file);
    return addImage(is, type, coordFrom, percentOfSize);
  }

  public Image addImage(InputStream is, ImageType type, SheetCoord coordFrom, SheetCoord coordTo) {

    Image image = addImageFile(is, type);
    addImage(image, coordFrom, coordTo);
    return image;
  }

  public Image addImage(InputStream is, ImageType type, SheetCoord coordFrom, int percentOfSize) {

    Image image = addImageFile(is, type);
    addImage(image, coordFrom, percentOfSize);
    return image;
  }

  /**
   * Добавление объекта с изображением, прописываются внутренние ссылки и связи.
   *
   * @param image     изображение
   * @param coordFrom координата, куда ставиться изображение
   * @param coordTo   координата, куда ставиться изображение
   */
  public void addImage(Image image, SheetCoord coordFrom, SheetCoord coordTo) {

    if (image == null) throw new IllegalArgumentException("Не задано изображение");
    if (image.getFilename() == null) throw new IllegalArgumentException("Не задано изображение");

    setDrawingId();

    Integer imagerel = imagesrels.get(image);
    if (imagerel == null) imagesrels.put(image, Integer.valueOf(++drawingRelIdLast));

    TwoCellAnchor anchor;
    anchor = new TwoCellAnchorImage(imagesrels.get(image), image, coordFrom, coordTo);

    drawing.add(anchor);
    parent.imageexts.add(image.getImageType().getExt());
  }

  private final static double K_IMAGEX = .1285; // коэффициент для пикселов изображения по ширине
  private final static double K_IMAGEY = 0.74; // по высоте

  public void addImage(Image image, SheetCoord coordFrom, int percentOfSize) {

    SheetCoord coordTo = new SheetCoord();
    coordTo.col = coordFrom.col;
    coordTo.row = coordFrom.row;

    // размеры изображения в excel

    double imagewidthexp = image.width * K_IMAGEX * percentOfSize / 100;
    double imageheightexp = image.height * K_IMAGEY * percentOfSize / 100;

    // вычисление колонки правой нижней ячейки 

    double imagewidth = getColWidth(coordFrom.col) - coordFrom.coloff / K_COLOFF;

    while (imagewidth < imagewidthexp) {
      coordTo.col++;
      imagewidth += getColWidth(coordFrom.col);
    }

    coordTo.coloff = (int) ((imagewidthexp - imagewidth) * K_COLOFF);

    // вычисление строки правой нижней ячейки

    double imageheight = getRowHeight(coordFrom.row) - coordFrom.rowoff / K_ROWOFF;

    while (imageheight < imageheightexp) {
      coordTo.row++;
      imageheight += getRowHeight(coordFrom.row);
    }

    coordTo.rowoff = (int) ((imageheightexp - imageheight) * K_ROWOFF);

    addImage(image, coordFrom, coordTo);
  }

  private final static long K_COLOFF = 65500; // для отступов по горизонтали
  private final static long K_ROWOFF = 11500; // для отступов по вертикали

  private double getColWidth(int col) {

    double width = 8.43; // ширина колонки по-умолчанию
    Double mapwidth = widths.get(Integer.valueOf(col));
    if (mapwidth != null && mapwidth > 0) width = mapwidth.doubleValue();

    return width;
  }

  private double getRowHeight(int row) {

    double height = 15.0; // ширина колонки по-умолчанию
    Double mapheight = heights.get(Integer.valueOf(row));
    if (mapheight != null && mapheight > 0) height = mapheight.doubleValue();

    return height;
  }

  /**
   * Добавление файла с изображением.
   *
   * @param is           Поток с изображением
   * @param type         Формат файла
   * @param getImageSize нужно ли определить размеры изображения
   * @return Объект с картинкой
   */
  private Image addImageFile(InputStream is, ImageType type) {

    if (is == null) throw new IllegalArgumentException("Не задан поток с изображением");
    if (type == null) throw new IllegalArgumentException("Не задан тип изображения");

    int fileid = parent.newImageFileId();
    String filename = "image" + fileid + "." + type.getExt();
    Image image = new Image(filename, type);

    try {
      String dir = workDir + "/xl/media";
      new File(dir).mkdirs();

      File os = new File(dir + "/" + filename);
      Files.copy(is, os.toPath());

      is.close();

      BufferedImage img = ImageIO.read(os);
      image.height = img.getHeight();
      image.width = img.getWidth();
    } catch (Exception ex) {}

    return image;
  }

  private InputStream getISFromBytes(byte[] img, ImageType type) {

    if (img == null) throw new IllegalArgumentException("Не задано изображение");
    if (img.length == 0) throw new IllegalArgumentException("Задано пустое изображение");
    if (type == null) throw new IllegalArgumentException("Не задан тип изображения");

    InputStream is = new ByteArrayInputStream(img);

    return is;
  }

  private InputStream getISFromFile(File file) {

    if (file == null) throw new IllegalArgumentException("Не задан файл с изображением");
    if (!file.exists()) throw new IllegalArgumentException(
      "Указанный файл с изображением не существует");

    InputStream is = null;
    try {
      is = new FileInputStream(file);
    } catch (FileNotFoundException ex) {
      throw new IllegalArgumentException("Указанный файл с изображением не найден", ex);
    }

    return is;
  }

  private ImageType getImageTypeFromFile(File file) {

    ImageType type = ImageType.getByExt(file.getName().substring(
      file.getName().lastIndexOf(".") + 1));
    if (type == null) throw new IllegalArgumentException("Неизвестный тип изображения");

    return type;
  }

  private void setDrawingId() {
    if (drawingFileId != null) return;
    drawingFileId = parent.getDrawingIdNext();
  }

  Integer getDrawingId() {
    return drawingFileId;
  }

  List<TwoCellAnchor> getDrawing() {
    return drawing;
  }

  Map<Image, Integer> getImagesRels() {
    return imagesrels;
  }

  private void printDrawings() {

    if (drawingFileId == null) return;

    out.println("<drawing r:id=\"rId" + getDrawingId() + "\"/>");

    try {
      String dir = workDir + "/xl/worksheets/_rels";
      new File(dir).mkdirs();

      PrintStream os = new PrintStream(dir + "/" + name + ".xml.rels", "UTF-8");

      os.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
      os.print("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId");
      os.print(getDrawingId());
      os.print("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing");
      os.print(getDrawingId());
      os.println(".xml\"/>");
      os.print("</Relationships>");

      os.close();
    } catch (Exception ex) {}
  }
}
