package kz.greetgo.msoffice.xlsx.gen;

import static org.fest.assertions.api.Assertions.assertThat;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;

import org.testng.annotations.Test;

public class BrokenPipeError {
  
  @Test
  public void brokenPipe() throws Exception {
    String workDir = "build/TestBrokenPipe";
    final String msg = "Broken pipe";
    
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
    
    OutputStream fout = new OutputStream() {
      private int c = 0;
      
      @Override
      public void write(int b) throws IOException {
        if (++c > 24) {
          throw new IOException(msg);
        }
      }
    };
    
    RuntimeException error = null;
    
    try {
      f.complete(fout);
    } catch (RuntimeException ex) {
      error = ex;
    }
    
    assertThat(error).isNotNull();
    assertThat(error.getMessage()).isEqualTo("java.io.IOException: " + msg);
    
    assertThat(new File(workDir).list()).isNullOrEmpty();
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
