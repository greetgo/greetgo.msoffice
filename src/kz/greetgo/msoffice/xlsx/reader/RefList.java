package kz.greetgo.msoffice.xlsx.reader;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.AbstractList;

public class RefList extends AbstractList<Ref> implements AutoCloseable {

  private final Path filePath;
  private final RandomAccessFile file;
  private final byte[] buffer = new byte[Ref.SIZE];

  public RefList(Path filePath) {
    this.filePath = filePath;
    try {
      file = new RandomAccessFile(filePath.toFile(), "rw");
    } catch (FileNotFoundException e) {
      throw new RuntimeException(e);
    }
  }

  @Override
  public Ref get(int index) {

    try {

      file.seek((long) Ref.SIZE * (long) index);
      int count = file.read(buffer);
      if (count < Ref.SIZE) {
        throw new IndexOutOfBoundsException("index = " + index + " but must be < " + size());
      }

      return Ref.readFrom(buffer, 0);

    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  @Override
  public int size() {
    try {
      return (int) (file.length() / Ref.SIZE);
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  @Override
  public void close() {
    try {
      file.close();
      Files.delete(filePath);
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  @Override
  public void add(int index, Ref ref) {
    try {

      file.seek((long) Ref.SIZE * (long) index);
      ref.writeTo(buffer, 0);
      file.write(buffer);

    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }
}
