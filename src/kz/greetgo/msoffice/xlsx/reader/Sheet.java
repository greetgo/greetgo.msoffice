package kz.greetgo.msoffice.xlsx.reader;

public interface Sheet {
  String name();

  int index();

  int rowCount();

  Row row(int rowIndex);

  default Cell cell(int rowIndex, int colIndex) {
    return row(rowIndex).cell(colIndex);
  }
}
