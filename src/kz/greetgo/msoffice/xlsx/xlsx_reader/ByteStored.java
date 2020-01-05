package kz.greetgo.msoffice.xlsx.xlsx_reader;

public interface ByteStored {
  int byteSize();

  void writeBytes(byte[] buffer, int offset);


}
