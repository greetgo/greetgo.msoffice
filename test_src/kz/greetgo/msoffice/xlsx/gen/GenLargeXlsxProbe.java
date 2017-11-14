package kz.greetgo.msoffice.xlsx.gen;

import java.io.FileOutputStream;

public class GenLargeXlsxProbe {
  public static void main(String[] args) throws Exception {
    int rowCount = 9000;
    
    String tmpDir = System.getProperty("user.home") + "/trans/tmp";
    String workDir = tmpDir + "/generated_xlsx_" + rowCount;
    
    try {
      Runtime.getRuntime().exec(/*wmd*/"rm -rvf " + workDir).waitFor();
    } catch (Exception e) {
      e.printStackTrace();
    }
    
    long time1 = System.currentTimeMillis();
    
    Xlsx f = new Xlsx();
    f.setWorkDir(workDir);
    
    Sheet sheet = f.newSheet(true);
    sheet.skipRows(3);
    
    for (int row = 0; row < rowCount; row++) {
      sheet.row().start();
      sheet.cellInt(1, 123);
      
      sheet.style().alignment().horizontalCenter();
      for (int i = 0; i < 20; i++) {
        if (i == 10) sheet.style().alignment().clean();
        sheet.cellStr(2 + i, "Ячейка Ы" + i + "_" + row + " of " + rowCount);
      }
      
      sheet.row().finish();
    }
    
    f.finish();
    
    long time2 = System.currentTimeMillis();
    
    System.out.println("Time to form of " + rowCount + " is " + (time2 - time1) / 1000.0 + " c");
    
    FileOutputStream fOut = new FileOutputStream(tmpDir + "/generated_" + rowCount + ".xlsx");
    
    f.print(fOut);
    
    fOut.close();
    
    long time3 = System.currentTimeMillis();
    
    System.out.println("Time to print of " + rowCount + " is " + (time3 - time2) / 1000.0 + " c");
    System.out.println("Total time of " + rowCount + " is " + (time3 - time1) / 1000.0 + " c");
    System.out.println("COMPLETE");
  }
}
