package kz.greetgo.sqlmanager.gen;

public class TooLongIdException extends RuntimeException {
  
  public TooLongIdException() {}
  
  public TooLongIdException(String message, Throwable cause) {
    super(message, cause);
  }
  
  public TooLongIdException(String message) {
    super(message);
  }
  
  public TooLongIdException(Throwable cause) {
    super(cause);
  }
  
}
