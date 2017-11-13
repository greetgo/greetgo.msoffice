package kz.greetgo.msoffice.xlsx.probes;

import java.util.List;

import kz.greetgo.msoffice.xlsx.parse.Cell;
import kz.greetgo.msoffice.xlsx.parse.RowHandler;
import kz.greetgo.msoffice.xlsx.parse.XlsxParser;

public class ParseBook1 {
  public static void main(String[] args) {
    XlsxParser p = new XlsxParser();
    p.removeWorkDirOnClose = false;
    
    p.load(ParseBook1.class.getResourceAsStream("Book1.xlsx"));
    
    List<Cell> row = p.activeSheet().loadRow(2);
    
    System.out.println(row.size() + " ~ " + row);
    
    p.activeSheet().scanRows(10, new RowHandler() {
      @Override
      public void handle(List<Cell> row, int rowIndex) throws Exception {
        System.out.println("[" + rowIndex + "] " + row);
      }
    });
    
    p.close();
  }
}
