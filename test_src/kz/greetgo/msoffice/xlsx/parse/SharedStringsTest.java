package kz.greetgo.msoffice.xlsx.parse;

import org.testng.annotations.Test;

public class SharedStringsTest {
  @Test
  public void asd() throws Exception {
    SharedStrings ss = new SharedStrings("tmp/asdf");
    ss.index("fdsf dsa fdsa fdsaf ");
    ss.index("Начинается новый день - 新的一天");
    ss.index("Начинается новый день - 新的fdsafdsf一天");
    ss.index("Начинается новый vc f gfds gfdg день - 新的fdsafdsf一天");
    
    ss.close();
    System.out.println("COMPLETE");
  }
}
