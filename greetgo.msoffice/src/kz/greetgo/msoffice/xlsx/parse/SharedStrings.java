package kz.greetgo.msoffice.xlsx.parse;

import java.io.File;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.io.RandomAccessFile;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;

public class SharedStrings {
  private static final int SPACE_FOR_HEADER = 200;
  
  private final File file;
  
  public SharedStrings(String workDir) throws Exception {
    
    file = new File(workDir + "/xl/sharedStrings.xml");
    prepareFile();
    
  }
  
  private PrintStream printStream;
  
  private final void prepareFile() throws Exception {
    file.getParentFile().mkdirs();
    
    printStream = new PrintStream(new FileOutputStream(file), false, "UTF-8");
    
    StringBuilder sb = new StringBuilder();
    for (int i = 0; i < SPACE_FOR_HEADER; i++) {
      sb.append(' ');
    }
    sb.append("\r\n");
    
    printStream.print(sb.toString());
  }
  
  private int index = 0;
  
  public int index(String str) {
    StringBuilder sb = new StringBuilder();
    sb.append("<si><t>");
    if (str == null) str = "";
    str = str.replaceAll("&", "&amp;");
    str = str.replaceAll("<", "&lt;");
    str = str.replaceAll(">", "&gt;");
    sb.append(str);
    sb.append("</t></si>\r\n");
    
    printStream.print(sb.toString());
    
    return index++;
  }
  
  public void close() throws Exception {
    printStream.print("</sst>");
    printStream.flush();
    printStream.close();
    printStream = null;
    
    RandomAccessFile raf = new RandomAccessFile(file, "rw");
    FileChannel fc = raf.getChannel();
    
    {
      fc.position(0);
      StringBuilder sb = new StringBuilder();
      sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
      sb.append("<sst xmlns=\"http://schemas.openxmlformats.org/"
          + "spreadsheetml/2006/main\" count=\"" + index + "\" uniqueCount=\"" + index + "\">");
      
      {
        ByteBuffer buf = ByteBuffer.wrap(sb.toString().getBytes("UTF-8"));
        while (buf.hasRemaining()) {
          fc.write(buf);
        }
      }
    }
    
    fc.close();
    raf.close();
  }
}
