package kz.greetgo.msoffice.xlsx.probes;

import java.util.Date;
import java.util.List;

import kz.greetgo.msoffice.xlsx.parse.Cell;
import kz.greetgo.msoffice.xlsx.parse.RowHandler;
import kz.greetgo.msoffice.xlsx.parse.XlsxParserMemory;

public class ParseBook3Memory {
  public static void main(String[] args) {
    XlsxParserMemory p = new XlsxParserMemory();
    p.removeWorkDirOnClose = false;
    
    System.out.println("start=" + new Date());
    p.load(ParseBook3Memory.class.getResourceAsStream("reportClientAttributet.xlsx"));
    System.out.println("end=" + new Date());
    int r = 0;
    List<Cell> row = p.activeSheet().loadRow(r);
    System.out.println("row=" + r + " size=" + row.size() + " ~ " + row);
    
    System.out.println("start=" + new Date());
    p.activeSheet().scanRows(10, new RowHandler() {
      @Override
      public void handle(List<Cell> row, int rowIndex) throws Exception {
        //System.out.println("[" + rowIndex + "] " + row);
      }
    });
    System.out.println("end=" + new Date());
  }
}
