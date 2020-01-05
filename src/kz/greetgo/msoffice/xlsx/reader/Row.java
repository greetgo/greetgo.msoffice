package kz.greetgo.msoffice.xlsx.reader;

public interface Row {
  double height();

  Cell cell(int i);

  int count();
}
