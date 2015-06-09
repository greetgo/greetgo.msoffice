package kz.greetgo.msoffice.xlsx.gen;

public enum Align {
	
  left("l"), //
  top("t"), //
  right("r"), //
  bottom("b"), //
  center("c");
  
  private String tag;
  
  private Align(String tag) {
	this.tag = tag;
  }
  
  String getTag() { return tag; }
  
  public String str() {
    return name();
  }
}