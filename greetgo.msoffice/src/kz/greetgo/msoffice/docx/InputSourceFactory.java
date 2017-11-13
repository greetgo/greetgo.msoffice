package kz.greetgo.msoffice.docx;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

public class InputSourceFactory {
  private InputSourceFactory() {}
  
  public static InputSource createByFile(final File file) {
    return new InputSource() {
      @Override
      public InputStream openInputStream() throws Exception {
        return new FileInputStream(file);
      }
    };
  }
  
  public static InputSource createByResource(final Class<?> classs, final String name) {
    return new InputSource() {
      @Override
      public InputStream openInputStream() throws Exception {
        return classs.getResourceAsStream(name);
      }
    };
  }
  
  public static InputSource createByResource(final String name) {
    return new InputSource() {
      @Override
      public InputStream openInputStream() throws Exception {
        return Object.class.getResourceAsStream(name);
      }
    };
  }
}
