package kz.greetgo.msoffice.xlsx.reader.model;

import kz.greetgo.msoffice.xlsx.reader.Ref;
import kz.greetgo.msoffice.xlsx.reader.RefList;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

public class SheetData implements AutoCloseable {

  public final List<ColumnInfo> columnInfoList = new ArrayList<>();

  public final int id;
  private final RefList rowRefList;
  private final RefList mergeCellRefList;
  private final RandomAccessFile data;

  public boolean tabSelected = false;

  public SheetData(int id, Function<String, Path> createTmpFile) {
    this.id = id;
    rowRefList = new RefList(createTmpFile.apply("sheet" + id + "-row-ref-list"));
    mergeCellRefList = new RefList(createTmpFile.apply("sheet" + id + "-merge-cell-ref-list"));
    try {
      data = new RandomAccessFile(createTmpFile.apply("sheet" + id + "-data").toFile(), "rw");
    } catch (FileNotFoundException e) {
      throw new RuntimeException(e);
    }
  }

  public void addRow(RowData rowData) {
    if (rowData.index < rowSize) {
      throw new IllegalArgumentException("index must be >= rowSize:"
        + " index = " + rowData.index
        + ", rowSize = " + rowSize);
    }

    while (rowSize < rowData.index) {
      append(RowData.empty(rowSize));
    }
    append(rowData);
  }

  private void append(RowData rowData) {
    try {

      Ref ref = Ref.of(data.length(), rowData.byteSize());
      rowRefList.add(ref);
      rowSize = rowRefList.size();

      byte[] buffer = new byte[ref.length];
      rowData.writeTo(buffer, 0);

      data.seek(data.length());
      data.write(buffer);

    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  private int rowSize = 0;

  public int rowSize() {
    return rowSize;
  }

  public RowData getRowData(int index) {
    if (index < 0) throw new IllegalArgumentException("index = " + index + " bug must be >= 0");
    if (index >= rowSize) return RowData.empty(index);

    try {
      Ref ref = rowRefList.get(index);

      byte[] buffer = new byte[ref.length];
      data.seek(ref.offset);
      assertReadOk(data.read(buffer), buffer);

      return RowData.readFrom(buffer, 0);

    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  private void assertReadOk(int readCount, byte[] buffer) {
    if (readCount != buffer.length) {
      throw new RuntimeException("readCount != buffer.length, BUT must be equals:" +
        " readCount =" + readCount + ", buffer.length = " + buffer.length);
    }
  }

  @Override
  public void close() {
    try {
      data.close();
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  private int mergeCellSize = 0;

  public void addMergeCell(MergeCell mergeCell) {
    try {

      Ref ref = Ref.of(data.length(), mergeCell.byteSize());
      mergeCellRefList.add(ref);
      mergeCellSize = mergeCellRefList.size();

      byte[] buffer = new byte[ref.length];
      mergeCell.writeBytes(buffer, 0);

      data.seek(data.length());
      data.write(buffer);

    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  public int mergeCellSize() {
    return mergeCellSize;
  }

  public MergeCell getMergeCell(int index) {
    if (index < 0 || index >= mergeCellSize) {
      throw new IndexOutOfBoundsException("index = " + index + ", mergeCellSize = " + mergeCellSize);
    }

    try {
      Ref ref = mergeCellRefList.get(index);

      byte[] buffer = new byte[ref.length];
      data.seek(ref.offset);
      assertReadOk(data.read(buffer), buffer);

      return MergeCell.readFrom(buffer, 0);

    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

}
