package kz.greetgo.msoffice.xlsx.fastgen.simple_inline;

import static kz.greetgo.msoffice.util.UtilOffice.toTablePosition;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.UUID;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import kz.greetgo.msoffice.xlsx.fastgen.simple.SimpleRowStyle;

public class SimpleInlineFastXlsxFile {
  
  private final String workingDir;
  private int sheetNumber = 0;
  private double[] colWidths;
  
  private PrintStream outCurrentSheet = null;
  private PrintStream outSharedStrings = null;
  
  public SimpleInlineFastXlsxFile(String tmpDir) {
    workingDir = tmpDir + "/" + UUID.randomUUID().toString();
    new File(workingDir).mkdirs();
    
    try {
      
      prepareFiles();
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }
  
  private void prepareFiles() throws Exception {
    ZipInputStream zin = new ZipInputStream(getClass().getResourceAsStream(
        "SimpleInlineFastXlsxFile.template.xlsx"));
    try {
      
      byte buf[] = new byte[1024 * 4];
      
      while (true) {
        ZipEntry entry = zin.getNextEntry();
        if (entry == null) break;
        if (entry.getName().endsWith("/")) continue;
        
        File entryFile = new File(workingDir + "/zip/" + entry.getName());
        
        entryFile.getParentFile().mkdirs();
        
        OutputStream fout = new FileOutputStream(entryFile);
        try {
          
          IN: while (true) {
            int count = zin.read(buf);
            if (count < 0) break IN;
            fout.write(buf, 0, count);
          }
          
        } finally {
          fout.close();
        }
        
        zin.closeEntry();
      }
    } finally {
      zin.close();
    }
    
    openSharedStrings();
  }
  
  public void newSheet(double[] colWidths) {
    this.colWidths = colWidths;
    
    closeCurrentSheet();
    sheetNumber++;
    try {
      outCurrentSheet = new PrintStream(workingDir + "/zip/xl/worksheets/sheet" + sheetNumber
          + ".xml", "UTF-8");
      outCurrentSheet.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
      outCurrentSheet
          .println("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
              + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
              + "  <dimension ref=\"A1:J4\" />\n" + "  <sheetViews>\n"
              + "    <sheetView tabSelected=\"1\" workbookViewId=\"0\">\n"
              + "      <selection activeCell=\"A1\" sqref=\"A1\" />\n" + "    </sheetView>\n"
              + "  </sheetViews>\n" + "  <sheetFormatPr defaultRowHeight=\"15\" />");
      outCurrentSheet.println("<cols>");
      for (int i = 1; i <= colWidths.length; i++) {
        outCurrentSheet.println("<col min=\"" + i + "\" max=\"" + i + "\" width=\""
            + colWidths[i - 1] + "\" customWidth=\"1\" />");
      }
      outCurrentSheet.println("</cols>");
      outCurrentSheet.println("<sheetData>");
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
    
    currentRow = 1;
  }
  
  private int currentRow;
  
  public void appendRow(SimpleRowStyle rowStyle, String[] rowData) {
    if (rowData == null) throw new NullPointerException("rowData == null");
    if (rowData.length != colWidths.length) throw new NullPointerException("rowData.length == "
        + rowData.length + " but must be == " + colWidths.length);
    if (sheetNumber <= 0) throw new RuntimeException("No sheet");
    outCurrentSheet.print("<row r=\"" + currentRow + "\">");
    for (int i = 0, C = rowData.length; i < C; i++) {
      String r = " r = \"" + toTablePosition(currentRow, i) + '"';
      String s = getRowStyle(rowStyle);
      
      String value = rowData[i];
      
      if (value == null) {
        outCurrentSheet.println("<c " + r + s + " />");
      } else {
        outCurrentSheet.println("<c " + r + s + " t=\"inlineStr\"><is><t>" + quote(rowData[i])
            + "</t></is></c>");
      }
    }
    outCurrentSheet.println("</row>");
    currentRow++;
  }
  
  private String quote(String s) {
    if (s == null) s = "";
    s = s.replaceAll("&", "&amp;");
    s = s.replaceAll("<", "&lt;");
    s = s.replaceAll(">", "&gt;");
    return s;
  }
  
  private String getRowStyle(SimpleRowStyle rowStyle) {
    switch (rowStyle) {
    case NORMAL:
      return "";
      
    case GREEN:
      return " s=\"1\"";
      
    default:
      throw new IllegalArgumentException("Unknown rowStyle = " + rowStyle);
    }
    
  }
  
  public void complete(OutputStream output) {
    try {
      
      closeCurrentSheet();
      
      ZipOutputStream zout = new ZipOutputStream(output);
      
      byte[] buffer = new byte[1024 * 4];
      
      new File(workingDir + "/zip/xl/workbook.xml").delete();
      new File(workingDir + "/zip/[Content_Types].xml").delete();
      new File(workingDir + "/zip/xl/_rels/workbook.xml.rels").delete();
      new File(workingDir + "/zip/docProps/app.xml").delete();
      
      new File(workingDir + "/zip/xl/sharedStrings.xml").delete();
      
      copyAll(zout, null, workingDir + "/zip", buffer);
      
      printWorkbook(zout);
      printContentTypes(zout);
      printWorkbookXmlRels(zout);
      printApp(zout);
      
      printSharedStrings(zout);
      
      zout.close();
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
    
  }
  
  private void copyAll(ZipOutputStream zout, String localDir, String realDir, byte[] buffer)
      throws Exception {
    for (String name : new File(realDir).list(null)) {
      File file = new File(realDir + "/" + name);
      String localName = localDir == null ? name :localDir + "/" + name;
      if (file.isDirectory()) {
        copyAll(zout, localName, realDir + "/" + name, buffer);
      } else if (file.isFile()) {
        zout.putNextEntry(new ZipEntry(localName));
        FileInputStream fin = new FileInputStream(file);
        try {
          while (true) {
            int count = fin.read(buffer);
            if (count < 0) break;
            zout.write(buffer, 0, count);
          }
        } finally {
          fin.close();
        }
        zout.closeEntry();
      }
    }
  }
  
  private void closeCurrentSheet() {
    if (outCurrentSheet == null) return;
    
    outCurrentSheet.print("</sheetData><pageMargins left=\"0.7\" right=\"0.7\""
        + " top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\" /></worksheet>");
    outCurrentSheet.close();
    outCurrentSheet = null;
  }
  
  private int newStringIndex = 0;
  
  private void openSharedStrings() throws Exception {
    outSharedStrings = new PrintStream(workingDir + "/sharedStrings.xml", "UTF-8");
  }
  
  @SuppressWarnings("unused")
  private int str(String s) {
    s = s.replaceAll("<", "&lt;");
    outSharedStrings.print("<si><t>" + s + "</t></si>");
    return newStringIndex++;
  }
  
  private void printWorkbook(ZipOutputStream zout) throws Exception {
    zout.putNextEntry(new ZipEntry("xl/workbook.xml"));
    
    PrintStream out = new PrintStream(zout, false, "UTF-8");
    
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.println("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
        + "  <fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\" />\n"
        + "  <workbookPr defaultThemeVersion=\"124226\" />\n"
        + "  <bookViews>\n"
        + "    <workbookView xWindow=\"240\" yWindow=\"15\" windowWidth=\"25680\" windowHeight=\"10305\" />\n"
        + "  </bookViews>");
    
    out.println("<sheets>");
    for (int i = 0; i < sheetNumber; i++) {
      out.println("<sheet name=\"Лист" + (i + 1) + "\" sheetId=\"" + (i + 1) + "\" r:id=\"rId"
          + (i + 4) + "\" />");
    }
    out.println("</sheets>");
    
    out.println("<calcPr calcId=\"124519\" />\n" + "</workbook>");
    
    out.flush();
    
    zout.closeEntry();
    
  }
  
  private void printContentTypes(ZipOutputStream zout) throws Exception {
    zout.putNextEntry(new ZipEntry("[Content_Types].xml"));
    
    PrintStream out = new PrintStream(zout, false, "UTF-8");
    
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.println("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
        + "  <Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\" />\n"
        + "  <Override PartName=\"/xl/styles.xml\"\n"
        + "    ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" />\n"
        + "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />\n"
        + "  <Default Extension=\"xml\" ContentType=\"application/xml\" />\n"
        + "  <Override PartName=\"/xl/workbook.xml\"\n"
        + "    ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />\n"
        + "  <Override PartName=\"/docProps/app.xml\"\n"
        + "    ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\" />\n"
        + "  <Override PartName=\"/xl/sharedStrings.xml\"\n"
        + "    ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\" />\n"
        + "  <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\" />");
    
    for (int i = 0; i < sheetNumber; i++) {
      out.println("<Override PartName=\"/xl/worksheets/sheet" + (i + 1) + ".xml\"\n"
          + "    ContentType=\"application/vnd.openxmlformats-officedocument."
          + "spreadsheetml.worksheet+xml\" />");
    }
    out.println("</Types>");
    
    out.flush();
    
    zout.closeEntry();
  }
  
  private void printWorkbookXmlRels(ZipOutputStream zout) throws Exception {
    
    zout.putNextEntry(new ZipEntry("xl/_rels/workbook.xml.rels"));
    
    PrintStream out = new PrintStream(zout, false, "UTF-8");
    
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.println("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
        + "  <Relationship Id=\"rId3\"\n"
        + "    Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\" />\n"
        + "  <Relationship Id=\"rId1\"\n"
        + "    Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\" />\n"
        + "  <Relationship Id=\"rId2\"\n"
        + "    Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\" />");
    
    for (int i = 0; i < sheetNumber; i++) {
      out.println("<Relationship Id=\"rId" + (i + 4) + "\"\n"
          + "    Type=\"http://schemas.openxmlformats.org/officeDocument/2006/"
          + "relationships/worksheet\" Target=\"worksheets/sheet" + (i + 1) + ".xml\" />");
    }
    out.println("</Relationships>");
    
    out.flush();
    
    zout.closeEntry();
  }
  
  private void printApp(ZipOutputStream zout) throws Exception {
    
    zout.putNextEntry(new ZipEntry("docProps/app.xml"));
    
    PrintStream out = new PrintStream(zout, false, "UTF-8");
    
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.println("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"\n"
        + "  xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n"
        + "  <Application>Microsoft Excel</Application>\n"
        + "  <DocSecurity>0</DocSecurity>\n"
        + "  <ScaleCrop>false</ScaleCrop>\n"
        + "  <HeadingPairs>\n"
        + "    <vt:vector size=\"2\" baseType=\"variant\">\n"
        + "      <vt:variant>\n"
        + "        <vt:lpstr>Листы</vt:lpstr>\n"
        + "      </vt:variant>\n"
        + "      <vt:variant>\n"
        + "        <vt:i4>3</vt:i4>\n"
        + "      </vt:variant>\n"
        + "    </vt:vector>\n"
        + "  </HeadingPairs>\n" + "  <TitlesOfParts>");
    
    out.println("<vt:vector size=\"" + sheetNumber + "\" baseType=\"lpstr\">");
    
    for (int i = 0; i < sheetNumber; i++) {
      out.println("<vt:lpstr>Лист" + (i + 1) + "</vt:lpstr>");
    }
    
    out.println("</vt:vector>");
    
    out.println("</TitlesOfParts>\n" + "  <Company>Microsoft</Company>\n"
        + "  <LinksUpToDate>false</LinksUpToDate>\n" + "  <SharedDoc>false</SharedDoc>\n"
        + "  <HyperlinksChanged>false</HyperlinksChanged>\n"
        + "  <AppVersion>12.0000</AppVersion>\n" + "</Properties>");
    
    out.flush();
    
    zout.closeEntry();
  }
  
  private void printSharedStrings(ZipOutputStream zout) throws Exception {
    outSharedStrings.close();
    outSharedStrings = null;
    
    zout.putNextEntry(new ZipEntry("xl/sharedStrings.xml"));
    
    PrintStream out = new PrintStream(zout, false, "UTF-8");
    
    out.println("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.println("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\""
        + newStringIndex + "\" uniqueCount=\"" + newStringIndex + "\">");
    out.flush();
    
    {
      InputStream in = new FileInputStream(workingDir + "/sharedStrings.xml");
      try {
        byte buf[] = new byte[1024 * 4];
        while (true) {
          int count = in.read(buf);
          if (count < 0) break;
          zout.write(buf, 0, count);
        }
      } finally {
        in.close();
      }
    }
    
    out.println("</sst>");
    out.flush();
    
    zout.closeEntry();
  }
}
