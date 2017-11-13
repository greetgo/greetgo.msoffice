package kz.greetgo.msoffice.docx;

import java.io.InputStream;

public interface BinaryFilePart {
  public String getPartName();
  
  public InputStream openInputStream() throws Exception;
}
