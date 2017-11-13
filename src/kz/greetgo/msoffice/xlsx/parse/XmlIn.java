package kz.greetgo.msoffice.xlsx.parse;

import java.util.LinkedList;

public class XmlIn {
  private LinkedList<String> inPath = new LinkedList<String>();
  
  public void stepIn(String name) {
    inPath.add(name);
    chachedCurrent = null;
  }
  
  public void stepOut() {
    inPath.removeLast();
    chachedCurrent = null;
  }
  
  private String chachedCurrent = null;
  
  public String current() {
    if (chachedCurrent != null) return chachedCurrent;
    
    StringBuilder sb = new StringBuilder();
    for (String name : inPath) {
      if (sb.length() > 0) sb.append('/');
      sb.append(name);
    }
    
    return chachedCurrent = sb.toString();
  }
}
