package kz.greetgo.msoffice.xlsx.xlsx_reader.model;

import kz.greetgo.msoffice.xlsx.xlsx_reader.Ref;
import kz.greetgo.msoffice.xlsx.xlsx_reader.RefList;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

public class SheetData implements AutoCloseable {
  public final String name;

  public final List<ColumnInfo> columnInfoList = new ArrayList<>();

  private final RefList rowRefList;
  private final RefList mergeCellRefList;
  private final RandomAccessFile data;

  public SheetData(String name, Function<String, Path> createTmpFile) {
    this.name = name;
    rowRefList = new RefList(createTmpFile.apply(name + "-row-ref-list"));
    mergeCellRefList = new RefList(createTmpFile.apply(name + "-merge-cell-ref-list"));
    try {
      data = new RandomAccessFile(createTmpFile.apply(name + "-data").toFile(), "rw");
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

  public void addMergeCell(MergeCell mergeCell) {}

  private int rowSize = 0;

  public int rowSize() {
    return rowSize;
  }

  public RowData getRowData(int index) {
    if (index < 0 || index >= rowSize) {
      throw new IndexOutOfBoundsException("index = " + index + ", rowSize = " + rowSize);
    }


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
}
