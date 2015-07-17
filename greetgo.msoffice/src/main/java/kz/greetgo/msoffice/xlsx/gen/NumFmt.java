package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

public enum NumFmt {
  
  PERCENT("10", null), // процент с двумя знаками после запятой
  FRACTION("12", null), // простые дроби
  DMY("165", "dd\\/mm\\/yyyy"), //
  HMS("166", "hh:mm:ss"), //
  DMY_HMS("167", "dd\\/mm\\/yyyy\\ hh:mm:ss"), //
  YMD("168", "yyyy\\-mm\\-dd"), //
  YMD_HMS("169", "yyyy\\-mm\\-dd\\ hh:mm:ss"), //
  NUM_SIMPLE0("170", "0_ ;\\-0\\ "), //
  NUM_SPACE0("171", "#,##0_ ;\\-#,##0\\ "), //
  NUM_SIMPLE2("172", "0.00_ ;\\-0.00\\ "), //
  NUM_SPACE2("173", "#,##0.00_ ;\\-#,##0.00\\ ");
  
  private final String id;
  private final String formatCode;
  
  public String getId() {
    return id;
  }
  
  private NumFmt(String id, String formatCode) {
    this.id = id;
    this.formatCode = formatCode;
  }
  
  private static int extfmtn = 0;
  private static String extfmts = "";
  static {
    for (NumFmt fmt : values()) {
      if (fmt.formatCode != null) {
        extfmtn++;
        extfmts += "<numFmt numFmtId=\"" + fmt.id + "\" formatCode=\"" + fmt.formatCode + "\" />";
      }
    }
  }
  
  public static void printAll(PrintStream out) {
    
    out.println("<numFmts count=\"" + extfmtn + "\">");
    out.println(extfmts);
    out.println("</numFmts>");
  }
}