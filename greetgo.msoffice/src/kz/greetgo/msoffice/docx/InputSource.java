package kz.greetgo.msoffice.docx;

import java.io.InputStream;

public interface InputSource {
  public InputStream openInputStream() throws Exception;
}
