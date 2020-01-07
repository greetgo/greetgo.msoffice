package kz.greetgo.msoffice.docx;

import java.io.InputStream;

public interface BinaryFilePart {
  String getPartName();

  InputStream openInputStream() throws Exception;
}
