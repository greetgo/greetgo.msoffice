package kz.greetgo.msoffice.xlsx.reader;

public class SheetRef {
  public final int index;
  public final int id;
  public final String name;

  public SheetRef(int index, int id, String name) {
    this.index = index;
    this.id = id;
    this.name = name;
  }
}
