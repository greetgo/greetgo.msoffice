package kz.greetgo.msoffice.xlsx.gen;

import java.io.FileOutputStream;
import java.util.Date;

public class SerachXlsxError {
  public static void main(String[] args) throws Exception {
    String tmpDir = System.getProperty("user.home") + "/tmp";
    String workDir = tmpDir + "/asdss";
    workDir = "tmp/Search";
    
    try {
      Runtime.getRuntime().exec("rm -rvf " + workDir).waitFor();
    } catch (Exception e) {
      e.printStackTrace();
    }
    
    Xlsx f = new Xlsx();
    f.setWorkDir(workDir);
    
    sheet1(f);
    sheet2(f);
    sheet3(f);
    
    f.finish();
    
    FileOutputStream fout = new FileOutputStream("tmp/asd_Search.xlsx");
    
    f.print(fout);
    
    fout.close();
    
    System.out.println("COMPLETE");
  }
  
  private static void sheet1(Xlsx f) {
    Sheet sheet = f.newSheet(true);
    {
      sheet.row().start();
      sheet.cellInt(1, 1000);
      sheet.row().finish();
    }
  }
  
  private static void sheet2(Xlsx f) {
    Sheet s = f.newSheet(false);
    
    s.setDisplayName("Перестановки кубика Рубика");
    
    s.setWidth(2, 50);
    
    s.skipRows(3);
    {
      s.style().clean();
      s.style().font().bold();
      s.style().alignment().horizontalCenter();
      s.style().alignment().verticalCenter();
      s.style().borders().bottom().setStyle(BorderStyle.slantDashDot);
      
      s.row().start();
      s.cellStr(1, "ddd");
      s.cellInt(2, 200);
      s.cellInt(3, 333);
      s.row().finish();
    }
    {
      s.row().start();
      s.cellYMD_HMS(2, new Date());
      s.row().finish();
    }
  }
  
  private static void sheet3(Xlsx f) {
    Sheet sheet = f.newSheet(false);
    {
      sheet.row().start();
      sheet.cellInt(1, 3000);
      sheet.row().finish();
    }
    sheet.skipRows(2);
    int U = 10;
    {
      sheet.row().hiddenOutline().start();
      for (int i = 1; i <= 10; i++) {
        sheet.cellInt(i, 3000 + U * 10 + i);
      }
      sheet.row().finish();
    }
    {
      sheet.row().hiddenOutline().start();
      for (int i = 1; i <= 10; i++) {
        sheet.cellInt(i, 3000 + U * 10 + i);
      }
      sheet.row().finish();
    }
    {
      sheet.row().hiddenOutline().start();
      for (int i = 1; i <= 10; i++) {
        sheet.cellInt(i, 3000 + U * 10 + i);
      }
      sheet.row().finish();
    }
    sheet.row().collapsed().start().finish();
  }
}
