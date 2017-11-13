package kz.greetgo.msoffice.xlsx.probes;

import java.util.List;

import kz.greetgo.msoffice.xlsx.parse.Cell;
import kz.greetgo.msoffice.xlsx.parse.RowHandler;
import kz.greetgo.msoffice.xlsx.parse.XlsxParserMemory;

public class ParseBook1Memory {
  public static void main(String[] args) {
    XlsxParserMemory p = new XlsxParserMemory();
    p.removeWorkDirOnClose = false;
    
    p.load(ParseBook1Memory.class.getResourceAsStream("Book1.xlsx"));
    
    List<Cell> row = p.activeSheet().loadRow(2);
    
    System.out.println(row.size() + " ~ " + row);
    
    p.activeSheet().scanRows(10, new RowHandler() {
      @Override
      public void handle(List<Cell> row, int rowIndex) throws Exception {
        System.out.println("[" + rowIndex + "] " + row);
      }
    });
  }
}
