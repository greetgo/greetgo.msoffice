package kz.greetgo.msoffice.xlsx.parse2;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.Objects;

public class StoredStrings implements AutoCloseable {

  private final Path refPath;
  private final Path contentPath;
  private final int cutSize;
  private final int maxSize;
  private final RandomAccessFile ref, content;

  public StoredStrings(Path refPath, Path contentPath, int cutSize, int maxSize) {
    this.refPath = refPath;
    this.contentPath = contentPath;
    this.cutSize = cutSize;
    this.maxSize = maxSize;
    try {
      ref = new RandomAccessFile(refPath.toFile(), "rw");
      content = new RandomAccessFile(contentPath.toFile(), "rw");
    } catch (FileNotFoundException e) {
      throw new RuntimeException(e);
    }
  }

  @Override
  public void close() {
    try {
      ref.close();
      content.close();
      refPath.toFile().delete();
      contentPath.toFile().delete();
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  private long strCount = 0;

  public long strCount() {
    return strCount;
  }

  public long append(String str) {
    Objects.requireNonNull(str);
    try {

      byte[] bytes = str.getBytes(StandardCharsets.UTF_8);
      long offset = content.length();

      content.seek(offset);
      content.write(bytes, 0, bytes.length);

      Ref refData = Ref.of(offset, bytes.length);

      byte[] refBytes = new byte[Ref.SIZE];
      refData.writeTo(refBytes, 0);

      long refOffset = strCount++ * Ref.SIZE;
      ref.seek(refOffset);
      ref.write(refBytes, 0, refBytes.length);

      return refOffset / Ref.SIZE;

    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  public String get(long i) {
    try {

      long refOffset = i * Ref.SIZE;
      ref.seek(refOffset);
      byte[] refBytes = new byte[Ref.SIZE];
      ref.read(refBytes);

      Ref ref = Ref.readFrom(refBytes, 0);
      byte[] buffer = new byte[ref.length];
      content.seek(ref.offset);
      content.read(buffer);
      return new String(buffer, StandardCharsets.UTF_8);

    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }
}
