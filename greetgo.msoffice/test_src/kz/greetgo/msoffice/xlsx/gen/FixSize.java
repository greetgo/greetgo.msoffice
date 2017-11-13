package kz.greetgo.msoffice.xlsx.gen;

import java.io.File;
import java.io.PrintStream;

public class FixSize {
  public static void main(String[] args) throws Exception {
    int rowCount = 900000;
    
    String tmpDir = System.getProperty("user.home") + "/trans/tmp";
    String workDir = tmpDir + "/asdss_" + rowCount;
    
    File file = new File(workDir + "/xl/sharedStrings.xml");
    
    PrintStream out = new PrintStream(tmpDir + "/size_info.txt", "UTF-8");
    
    long timeStart = 0;
    long lastSize = 0;
    
    System.out.println("See go...");
    
    CI: while (true) {
      if (file.exists()) {
        long now = System.currentTimeMillis();
        long size = file.length();
        
        if (timeStart == 0) timeStart = now;
        if (lastSize > 0 && lastSize == size) break CI;
        lastSize = size;
        
        String s = "sharedStrings.xml " + (now - timeStart) + " " + size;
        out.println(s);
        System.out.println(s);
      } else if (timeStart > 0) {
        break CI;
      }
      
      Thread.sleep(3000);
    }
    
    out.close();
  }
}
