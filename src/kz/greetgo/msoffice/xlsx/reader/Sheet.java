package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.xlsx.reader.model.MergeCell;

public interface Sheet {
  String name();

  int index();

  int rowCount();

  Row row(int rowIndex);

  default Cell cell(int rowIndex, int colIndex) {
    return row(rowIndex).cell(colIndex);
  }

  boolean tabSelected();

  int frozenRowCount();

  int cellMergeCount();

  MergeCell cellMerge(int i);
}
