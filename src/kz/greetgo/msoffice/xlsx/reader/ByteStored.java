package kz.greetgo.msoffice.xlsx.reader;

public interface ByteStored {
  int byteSize();

  void writeBytes(byte[] buffer, int offset);


}
