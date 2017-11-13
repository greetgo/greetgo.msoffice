package kz.greetgo.msoffice.xlsx.probes;

import java.io.FileOutputStream;

import kz.greetgo.msoffice.xlsx.gen.BorderStyle;
import kz.greetgo.msoffice.xlsx.gen.Sheet;
import kz.greetgo.msoffice.xlsx.gen.Xlsx;

public class GenXlsxProbe2 {
  public static void main(String[] args) throws Exception {
    Xlsx f = new Xlsx();
    
    createSheet1(f);
    createSheet2(f);
    
    FileOutputStream fout = new FileOutputStream("build/asd3.xlsx");
    
    f.complete(fout);
    
    fout.close();
    
    System.out.println("COMPLETE");
  }
  
  private static void createSheet1(Xlsx f) {
    Sheet s = f.newSheet(true);
    
    s.setDisplayName("Перестановки кубика Рубика");
    
    s.setWidth(2, 50);
    
    s.skipRows(3);
    {
      s.style().clean();
      s.style().font().bold();
      s.style().alignment().horizontalCenter();
      s.style().alignment().verticalCenter();
      s.style().borders().bottom().setStyle(BorderStyle.slantDashDot);
      
      s.row().height(30).start();
      s.cellStr(1, "No");
      s.cellStr(2, "ФИО");
      s.cellStr(3, "История");
      s.cellInlineStr(4, "Пример inline");
      s.row().finish();
    }
    {
      s.row().start();
      s.cellDouble(1, 100);
      s.cellDouble(2, 100);
      s.cellFormula(3, "A5*100/B5");
      s.cellInlineStr(4, "A234*100/B5");
      s.row().finish();
    }
  }
  
  private static void createSheet2(Xlsx f) {
    Sheet s = f.newSheet(false);
    
    s.setDisplayName("Саламандра");
    
    s.setWidth(2, 50);
    
    s.skipRows(3);
    {
      s.style().clean();
      s.style().font().bold();
      s.style().alignment().horizontalCenter();
      s.style().alignment().verticalCenter();
      
      s.row().height(30).start();
      s.cellStr(1, "No555");
      s.cellStr(2, "ФИО667");
      s.cellStr(3, "Исторgfdhgfdjия");
      s.cellInlineStr(4, "Пример inline");
      s.row().finish();
    }
  }
}
