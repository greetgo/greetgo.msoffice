package kz.greetgo.msoffice.xlsx.probes;

import java.util.List;

import kz.greetgo.msoffice.xlsx.parse.Cell;
import kz.greetgo.msoffice.xlsx.parse.RowHandler;
import kz.greetgo.msoffice.xlsx.parse.XlsxParserMemory;

public class ParseBook2Memory {
  public static void main(String[] args) {
    XlsxParserMemory p = new XlsxParserMemory();
    p.removeWorkDirOnClose = false;
    
    p.load(ParseBook2Memory.class.getResourceAsStream("Book2.xlsx"));
    
    List<Cell> row = p.activeSheet().loadRow(0);
    System.out.println(row.size() + " ~ " + row);
    
    row = p.activeSheet().loadRow(1);
    System.out.println(row.size() + " ~ " + row);
    
    row = p.activeSheet().loadRow(2);
    System.out.println(row.size() + " ~ " + row);
    
    row = p.activeSheet().loadRow(3);
    System.out.println(row.size() + " ~ " + row);
    
    row = p.activeSheet().loadRow(4);
    System.out.println(row.size() + " ~ " + row);
    
    row = p.activeSheet().loadRow(5);
    System.out.println(row.size() + " ~ " + row);
    
    row = p.activeSheet().loadRow(6);
    System.out.println(row.size() + " ~ " + row);
    
    p.activeSheet().scanRows(10, new RowHandler() {
      @Override
      public void handle(List<Cell> row, int rowIndex) throws Exception {
        System.out.println("[ri=" + rowIndex + "] " + row);
      }
    });
    
    List<Cell> row5 = p.activeSheet().loadRow(3);
    System.out.println(row5.size() + " ~ " + row5);
  }
}
