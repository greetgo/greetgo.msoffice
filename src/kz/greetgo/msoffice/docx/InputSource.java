package kz.greetgo.msoffice.docx;

import java.io.InputStream;

public interface InputSource {
  InputStream openInputStream() throws Exception;
}
