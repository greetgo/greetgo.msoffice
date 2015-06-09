package kz.greetgo.msoffice.xlsx.gen;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Random;
import java.util.Set;

import kz.greetgo.msoffice.UtilOffice;
import kz.greetgo.msoffice.xlsx.parse.SharedStrings;

public class Xlsx {
  private String tmpDirBase = System.getProperty("java.io.tmpdir", ".");
  
  private String workDir = null;
  private SharedStrings strs = null;
  
  private final Styles styles = new Styles();
  private final Content content;
  
  private final Map<String, Sheet> sheets = new HashMap<String, Sheet>();
  private int nextSheetIndex = 1;
  private boolean wasSelected = false;
  
  private final Set<Chart> charts = new HashSet<Chart>();
  private int drawingIdLast = 0;
  private int chartIdLast = 0;
  
  public Xlsx() {
    
    content = new Content(this);
    styles.styleIndex(new Style(styles));//чтобы хотябы один всегда печатался
    styles.bordersIndex(new Borders());//чтобы хотябы один всегда печатался
    
    //Убираем макрасятовский касяк
    {
      Style s = new Style(styles);
      s.patternFill().setType(PatternFillType.gray125);
      styles.styleIndex(s);
    }
  }
  
  public Sheet newSheet(boolean selected) {
    if (selected) {
      if (wasSelected) {
        throw new IllegalArgumentException("Selected sheet may be only one");
      }
      wasSelected = true;
    }
    Sheet ret = new Sheet(this, styles, nextSheetIndex++, workDir(), strs(), selected);
    sheets.put(ret.name(), ret);
    content.addSheet(ret);
    
    return ret;
  }
  
  public void setWorkDir(String workDir) {
    if (strs != null) {
      throw new IllegalStateException("Cannot change workDir after initialization");
    }
    this.workDir = workDir;
  }
  
  private String workDir() {
    if (workDir == null) {
      String newName = "xlsxGen-" + System.currentTimeMillis() + "-" + new Random().nextLong();
      workDir = tmpDirBase + "/" + newName;
      new File(workDir).mkdirs();
    }
    return workDir;
  }
  
  private SharedStrings strs() {
    try {
      if (strs == null) strs = new SharedStrings(workDir());
      return strs;
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException)e;
      }
      throw new RuntimeException(e);
    }
  }
  
  private void printStyles() throws Exception {
    String dir = workDir() + "/xl";
    new File(dir).mkdirs();
    PrintStream out = new PrintStream(new FileOutputStream(dir + "/styles.xml"), false, "UTF-8");
    styles.print(out);
    out.flush();
    out.close();
  }
  
  void finish() {
    try {
      finishInner();
    } catch (Exception e) {
      if (e instanceof RuntimeException) {
        throw (RuntimeException)e;
      }
      throw new RuntimeException(e);
    }
  }
  
  private void finishInner() throws Exception {
    for (Sheet sheet : sheets.values()) {
      sheet.finish();
    }
    strs().close();
    printStyles();
    {
      content.setWorkDir(workDir());
      content.finish();
    }
  }
  
  void print(OutputStream out) {
    UtilOffice.zipDir(workDir(), out);
  }
  
  void close() {
    UtilOffice.removeDir(workDir());
  }
  
  public void complete(OutputStream out) {
    finish();
    print(out);
    close();
  }
  
  public Chart newChart(ChartType type) {
    
    Chart chart = new Chart(type, ++chartIdLast);
    charts.add(chart);
    
    return chart;
  }
  
  public void add(Chart chart) {
    charts.add(chart);
  }
  
  Set<Chart> getCharts() {
    return charts;
  }
  
  int getDrawingIdNext() {
    return ++drawingIdLast;
  }
  
  public static void main(String[] args) {
    
    System.out.println(new File(System.getProperty("java.io.tmpDir", ".")).getAbsolutePath());
  }
}