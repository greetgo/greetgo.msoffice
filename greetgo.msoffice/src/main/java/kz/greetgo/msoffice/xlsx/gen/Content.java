package kz.greetgo.msoffice.xlsx.gen;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

import kz.greetgo.msoffice.UtilOffice;

public class Content {
  
  private final Xlsx xlsx;
  
  private final List<Sheet> sheets = new ArrayList<Sheet>();
  final CoreProperties coreProperties = new CoreProperties();
  private String workDir;
  
  Content(Xlsx xlsx) {
    this.xlsx = xlsx;
  }
  
  public void addSheet(Sheet sheet) {
    sheets.add(sheet);
  }
  
  public void setWorkDir(String workDir) {
    this.workDir = workDir;
  }
  
  public void finish() throws Exception {
    printContentTypes();
    printWorkbook();
    printThemes();
    printDrawings();
    printCharts();
    printXl_rels();
    print_rels();
    printAppProperties();
    printCoreProperties();
  }
  
  private void printContentTypes() throws Exception {
    PrintStream out = new PrintStream(workDir + "/[Content_Types].xml", "UTF-8");
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    
    out.println("<Types xmlns=\"http://schemas.openxmlformats.org"
        + "/package/2006/content-types\">");
    out.println("<Override PartName=\"/xl/theme/theme1.xml\" "
        + "ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\" />");
    out.println("<Override PartName=\"/xl/styles.xml\" "
        + "ContentType=\"application/vnd.openxmlformats-officedocument"
        + ".spreadsheetml.styles+xml\" />");
    out.println("<Default Extension=\"rels\" "
        + "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />");
    out.println("<Default Extension=\"xml\" " + "ContentType=\"application/xml\" />");
    
    for (String ext : xlsx.imageexts)
      out.println("<Default Extension=\"" + ext + "\" " + "ContentType=\"image/" + ext + "\"/>");
    
    out.println("<Override PartName=\"/xl/workbook.xml\" "
        + "ContentType=\"application/vnd.openxmlformats-officedocument"
        + ".spreadsheetml.sheet.main+xml\" />");
    out.println("<Override PartName=\"/docProps/app.xml\" "
        + "ContentType=\"application/vnd.openxmlformats-officedocument"
        + ".extended-properties+xml\" />");
    
    for (Chart chart : xlsx.getCharts()) {
      out.print("<Override PartName=\"/xl/charts/chart");
      out.print(chart.getFileId());
      out.println(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>");
    }
    
    for (Sheet sheet : sheets) {
      out.println("<Override PartName=\"/xl/worksheets/" + sheet.name() + ".xml\" "
          + "ContentType=\"application/"
          + "vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />");
      
      if (sheet.getDrawing().size() < 1) continue;
      
      out.print("<Override PartName=\"/xl/drawings/drawing");
      out.print(sheet.getDrawingId());
      out.println(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
    }
    
    out.println("<Override PartName=\"/xl/sharedStrings.xml\" "
        + "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml"
        + ".sharedStrings+xml\" />");
    
    out.println("<Override PartName=\"/docProps/core.xml\" "
        + "ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\" />");
    out.println("</Types>");
    
    out.close();
  }
  
  private void printWorkbook() throws Exception {
    String dir = workDir + "/xl";
    new File(dir).mkdirs();
    PrintStream out = new PrintStream(dir + "/workbook.xml", "UTF-8");
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    
    out.println("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
        + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
    out.println("<fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\""
        + " rupBuild=\"4505\" />");
    out.println("<workbookPr defaultThemeVersion=\"124226\" />");
    out.println("<bookViews>");
    
    int activeIndex = 0;
    for (Sheet sheet : sheets) {
      if (sheet.selected) break;
      activeIndex++;
    }
    
    out.println("<workbookView xWindow=\"480\" yWindow=\"120\" windowWidth=\"23775\""
        + " windowHeight=\"8640\" activeTab=\"" + activeIndex + "\" />");
    out.println("</bookViews>");
    out.println("<sheets>");
    
    int id = 1;
    for (Sheet sheet : sheets) {
      out.println("<sheet name=\"" + sheet.getDisplayName() + "\" sheetId=\"" + id
          + "\" r:id=\"rId" + id + "\" />");
      id++;
    }
    
    out.println("</sheets>");
    out.println("<calcPr calcId=\"124519\" />");
    out.println("</workbook>");
    
    out.close();
  }
  
  private void printThemes() throws Exception {
    String dir = workDir + "/xl/theme";
    new File(dir).mkdirs();
    
    OutputStream out = new FileOutputStream(dir + "/theme1.xml");
    InputStream in = getClass().getResourceAsStream("theme1.xml");
    
    UtilOffice.copyStreams(in, out);
    
    in.close();
    out.close();
  }
  
  private void printXl_rels() throws Exception {
    String dir = workDir + "/xl/_rels";
    new File(dir).mkdirs();
    PrintStream out = new PrintStream(dir + "/workbook.xml.rels", "UTF-8");
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    
    out.println("<Relationships "
        + "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    
    int id = 1;
    for (Sheet sheet : sheets) {
      out.println("<Relationship Id=\"rId" + id + "\" "
          + "Type=\"http://schemas.openxmlformats.org/officeDocument/"
          + "2006/relationships/worksheet\"" + " Target=\"worksheets/" + sheet.name() + ".xml\" />");
      id++;
    }
    
    out.println("<Relationship Id=\"rId" + id++ + "\" "
        + "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/"
        + "relationships/sharedStrings\"" + " Target=\"sharedStrings.xml\" />");
    
    out.println("<Relationship Id=\"rId" + id++ + "\" "
        + "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\""
        + " Target=\"styles.xml\" />");
    
    out.println("<Relationship Id=\"rId" + id++ + "\" "
        + "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\""
        + " Target=\"theme/theme1.xml\" />");
    
    out.println("</Relationships>");
    
    out.close();
  }
  
  private void print_rels() throws Exception {
    String dir = workDir + "/_rels";
    new File(dir).mkdirs();
    PrintStream out = new PrintStream(dir + "/.rels", "UTF-8");
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    
    out.println("<Relationships "
        + "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    
    out.println("<Relationship" + " Id=\"rId3\""
        + " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/"
        + "relationships/extended-properties\"" + " Target=\"docProps/app.xml\" />");
    
    out.println("<Relationship" + " Id=\"rId2\""
        + " Type=\"http://schemas.openxmlformats.org/package/2006/"
        + "relationships/metadata/core-properties\"" + " Target=\"docProps/core.xml\" />");
    
    out.println("<Relationship" + " Id=\"rId1\""
        + " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/"
        + "relationships/officeDocument\"" + " Target=\"xl/workbook.xml\" />");
    
    out.println("</Relationships>");
    
    out.close();
  }
  
  private void printAppProperties() throws Exception {
    String dir = workDir + "/docProps";
    new File(dir).mkdirs();
    PrintStream out = new PrintStream(dir + "/app.xml", "UTF-8");
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    
    out.println("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/"
        + "2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/"
        + "officeDocument/2006/docPropsVTypes\">");
    out.println("<Application>Microsoft Excel</Application>");
    out.println("<DocSecurity>0</DocSecurity>");
    out.println("<ScaleCrop>false</ScaleCrop>");
    
    out.println("<HeadingPairs>");
    out.println("<vt:vector size=\"2\" baseType=\"variant\">");
    out.println("<vt:variant>");
    out.println("<vt:lpstr>Листы</vt:lpstr>");
    out.println("</vt:variant>");
    out.println("<vt:variant>");
    out.println("<vt:i4>" + sheets.size() + "</vt:i4>");
    out.println("</vt:variant>");
    out.println("</vt:vector>");
    out.println("</HeadingPairs>");
    
    out.println("<TitlesOfParts>");
    out.println("<vt:vector size=\"" + sheets.size() + "\" baseType=\"lpstr\">");
    for (Sheet sheet : sheets) {
      out.println("<vt:lpstr>" + sheet.getDisplayName() + "</vt:lpstr>");
    }
    out.println("</vt:vector>");
    out.println("</TitlesOfParts>");
    
    out.println("<Company>Microsoft</Company>");
    out.println("<LinksUpToDate>false</LinksUpToDate>");
    out.println("<SharedDoc>false</SharedDoc>");
    out.println("<HyperlinksChanged>false</HyperlinksChanged>");
    out.println("<AppVersion>12.0000</AppVersion>");
    
    out.println("</Properties>");
    
    out.close();
  }
  
  private void printCoreProperties() throws Exception {
    String dir = workDir + "/docProps";
    new File(dir).mkdirs();
    PrintStream out = new PrintStream(dir + "/core.xml", "UTF-8");
    coreProperties.print(out);
    out.close();
  }
  
  private void printCharts() throws Exception {
    
    if (xlsx.getCharts().size() < 1) return;
    
    String dir = workDir + "/xl/charts";
    new File(dir).mkdirs();
    
    for (Chart chart : xlsx.getCharts()) {
      PrintStream os = new PrintStream(dir + "/chart" + chart.getFileId() + ".xml", "UTF-8");
      chart.print(os);
      os.close();
    }
  }
  
  private void printDrawings() throws Exception {
    
    for (Sheet sheet : sheets) {
      if (sheet.getDrawing().size() < 1) continue;
      
      String dir = workDir + "/xl/drawings";
      new File(dir).mkdirs();
      String dirrel = workDir + "/xl/drawings/_rels";
      new File(dirrel).mkdirs();
      
      PrintStream os = new PrintStream(dir + "/drawing" + sheet.getDrawingId() + ".xml", "UTF-8");
      os.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
      os.println("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");
      
      PrintStream osrel = new PrintStream(dirrel + "/drawing" + sheet.getDrawingId() + ".xml.rels",
          "UTF-8");
      osrel.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
      osrel
          .println("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
      
      for (TwoCellAnchor anch : sheet.getDrawing()) {
        anch.print(os);
        
        osrel.print("<Relationship Id=\"rId");
        osrel.print(anch.getRelId());
        osrel.print("\" Type=\"");
        osrel.print(anch.getType());
        osrel.print("\"/>");
      }
      
      os.print("</xdr:wsDr>");
      os.close();
      
      osrel.print("</Relationships>");
      osrel.close();
    }
  }
}