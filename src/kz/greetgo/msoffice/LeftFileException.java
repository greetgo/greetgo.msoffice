package kz.greetgo.msoffice;

public class LeftFileException extends RuntimeException {
  private static final long serialVersionUID = 4609950215490132611L;
  
  public LeftFileException() {}
  
  public LeftFileException(String message, Throwable cause) {
    super(message, cause);
  }
  
  public LeftFileException(String message) {
    super(message);
  }
  
  public LeftFileException(Throwable cause) {
    super(cause);
  }
}
