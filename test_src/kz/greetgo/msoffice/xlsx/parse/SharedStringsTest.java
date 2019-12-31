package kz.greetgo.msoffice.xlsx.parse;

import org.testng.annotations.Test;

public class SharedStringsTest {
  @Test
  public void someTest() throws Exception {
    SharedStrings ss = new SharedStrings("tmp/asf");
    ss.index("fds1f dsa fds2a fds3af ");
    ss.index("Начинается новый день - 新的            一天");
    ss.index("Начинается новый день - 新的 fds6afd7sf 一天");
    ss.index("Начинается новый vc f gfd8s gf7dg день - 新的一天");
    
    ss.close();
    System.out.println("COMPLETE");
  }
}
