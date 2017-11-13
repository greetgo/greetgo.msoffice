package kz.greetgo.msoffice.xlsx.probes;

import java.io.FileOutputStream;

import kz.greetgo.msoffice.xlsx.gen.Color;
import kz.greetgo.msoffice.xlsx.gen.PatternFillType;
import kz.greetgo.msoffice.xlsx.gen.Sheet;
import kz.greetgo.msoffice.xlsx.gen.Xlsx;

public class GenXlsxProbe3 {
  public static void main(String[] args) throws Exception {
    Xlsx f = new Xlsx();
    
    Sheet s = f.newSheet(true);
    
    s.style().clean();
    s.style().patternFill().setType(PatternFillType.solid);
    s.style().patternFill().setFgColor(Color.green());
    
    s.row().start();
    s.cellStr(2, "Hello world!!!!");
    s.cellStr(3, "By & world <beeee>!!!!");
    s.row().finish();
    
    f.complete(new FileOutputStream("/tmp/asd-fill.xlsx"));
    System.out.println("COMPLETE");
  }
}
