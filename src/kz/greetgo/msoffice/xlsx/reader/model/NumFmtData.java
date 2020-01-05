package kz.greetgo.msoffice.xlsx.reader.model;

public class NumFmtData {
  public final int id;
  public String formatCode;

  public NumFmtData(String id) {
    this.id = Integer.parseInt(id);
  }
}
