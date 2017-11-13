package kz.greetgo.msoffice.xlsx.gen;

import java.io.File;
import java.io.FileOutputStream;

public class GenXlsxProbe {
  public static void main(String[] args) throws Exception {
    //    String outDir = "/home/pompei/trans/msoffice.test";
    String outDir = "build";
    String workDir = outDir + "/workDir";
    
    new File(outDir).mkdirs();
    new File(workDir).mkdirs();
    
    try {
      Runtime.getRuntime().exec("rm -rvf " + workDir).waitFor();
    } catch (Exception e) {
      e.printStackTrace();
    }
    
    Xlsx f = new Xlsx();
    f.setWorkDir(workDir);
    
    sheet1(f, false);
    sheet2(f, false);
    sheet3_outline(f, false);
    sheet4_numFmt(f, false);
    sheet5_outline(f, true);
    
    f.finish();
    
    FileOutputStream fout = new FileOutputStream(outDir + "/gen.xlsx");
    
    f.print(fout);
    
    fout.close();
    
    System.out.println("COMPLETE");
  }
  
  private static void sheet1(Xlsx f, boolean active) {
    Sheet sheet = f.newSheet(active);
    {
      sheet.style().patternFill().setType(PatternFillType.solid);
      sheet.style().patternFill().setFgColor(Color.blue());
      
      sheet.skipRows(3);
      
      sheet.row().height(100).start();
      sheet.cellInt(1, 123);
      
      sheet.style().noFill();
      sheet.cellStr(2, "fds fd fsd");
      
      sheet.style().patternFill().setType(PatternFillType.darkDown);
      sheet.style().patternFill().setFgColor(Color.red());
      sheet.style().patternFill().getFgColor().setAlfa(128);
      sheet.style().patternFill().setBgColor(Color.yellow());
      sheet.style().alignment().horizontalCenter();
      
      sheet.style().borders().top().clean();
      
      sheet.cellStr(3, "dfasd gfd sgfds gfds g");
      
      sheet.style().borders().top().medium();
      sheet.style().borders().top().setColor(Color.cyan());
      
      sheet.cellStr(4, "DFASD GFD SGFDS GFDS G");
      
      sheet.row().finish();
    }
    {
      sheet.row().start();
      sheet.cellInt(3, 321);
      sheet.cellStr(4, " FSDF SF SF SF ");
      sheet.row().finish();
    }
    {
      sheet.skipRow();
      sheet.row().start();
      sheet.style().clean();
      sheet.style().borders().bottom().medium();
      sheet.cellInt(1, 321);
      sheet.cellStr(2, "Марионетка");
      sheet.cellStr(3, null);
      sheet.row().finish();
    }
    {
      sheet.skipRow();
      sheet.row().start();
      sheet.style().clean();
      sheet.style().alignment().wrap();
      sheet.cellStr(2, "Марионетка vgfd sgdf sgdfs thytrhjuyr ur kuk");
      sheet.row().finish();
    }
    {
      sheet.skipRow();
      sheet.row().start();
      sheet.style().clean();
      sheet.style().borders().top().setColor(Color.cyan());
      sheet.style().borders().top().medium();
      sheet.cellStr(2, "Марионетка vgfd sgdf sgdfs thytrhjuyr ur kuk");
      sheet.cellStr(3, null);
      sheet.cellStr(4, null);
      //sheet.cellStr(5, null);
      //sheet.cellStr(6, null);
      sheet.addMergeInRow(2, 6);
      sheet.row().finish();
    }
  }
  
  private static void sheet2(Xlsx f, boolean active) {
    Sheet sheet = f.newSheet(active);
    {
      sheet.style().patternFill().setType(PatternFillType.solid);
      sheet.style().patternFill().setFgColor(Color.blue());
      
      sheet.row().start();
      sheet.cellInt(1, 123);
      
      sheet.style().noFill();
      sheet.cellStr(2, "fds fd fsd");
      
      sheet.style().patternFill().setType(PatternFillType.darkDown);
      sheet.style().patternFill().setFgColor(Color.red());
      sheet.style().patternFill().getFgColor().setAlfa(128);
      sheet.style().patternFill().setBgColor(Color.yellow());
      sheet.style().alignment().horizontalCenter();
      
      sheet.style().borders().top().clean();
      
      sheet.cellStr(3, "dfasd gfd sgfds gfds g");
      
      sheet.style().borders().top().medium();
      sheet.style().borders().top().setColor(Color.cyan());
      
      sheet.cellStr(4, "DFASD GFD SGFDS GFDS G");
      
      sheet.row().finish();
    }
    {
      sheet.row().start();
      sheet.cellInt(3, 321);
      sheet.cellStr(4, " FSDF SF SF SF ");
      sheet.row().finish();
    }
    {
      sheet.skipRows(2);
      sheet.row().start();
      sheet.cellInt(3, 321);
      sheet.cellStr(4, " FSDF SF SF SF fdsf dfdsf");
      sheet.addMergeInRow(4, 6);
      sheet.row().finish();
    }
  }
  
  private static void sheet3_outline(Xlsx f, boolean active) {
    Sheet sheet = f.newSheet(active);
    
    sheet.setSummaryBelow(false);
    sheet.skipRows(2);
    
    for (int i = 0; i < 5; i++) {
      sheet.row().outline().hidden().start();
      
      sheet.cellInt(1, 123 + i);
      sheet.cellStr(2, "asd" + i);
      sheet.cellStr(3, "dfasd gfd" + i);
      sheet.cellStr(4, "DFASD GFD" + i);
      
      sheet.row().finish();
    }
    {
      sheet.row().collapsed().start();
      sheet.cellStr(3, "The END");
      sheet.row().finish();
    }
  }
  
  private static void sheet4_numFmt(Xlsx f, boolean active) {
    Sheet sheet = f.newSheet(active);
    
    sheet.skipRows(2);
    sheet.setSummaryBelow(false);
    
    {
      sheet.row().collapsed().start();
      sheet.cellDouble(2, 1234235.56);
      sheet.row().finish();
    }
    {
      sheet.row().collapsed().start();
      sheet.cellDouble(2, 1234235.56, NumFmt.NUM_SPACE0);
      sheet.row().finish();
    }
    {
      sheet.row().collapsed().start();
      sheet.cellDouble(2, 435453465.56, NumFmt.NUM_SPACE2);
      sheet.row().finish();
    }
    {
      sheet.row().collapsed().start();
      sheet.cellDouble(2, 4.3532434354665476e70, NumFmt.NUM_SPACE2);
      sheet.row().finish();
    }
  }
  
  private static void sheet5_outline(Xlsx f, boolean active) {
    Sheet sheet = f.newSheet(active);
    sheet.setSummaryBelow(false);
    
    sheet.setWidth(1, 40);
    sheet.setWidth(2, 20);
    sheet.setWidth(3, 20);
    
    sheet.row().start();
    sheet.cellStr(1, "Пример группировки");
    sheet.row().finish();
    
    sheet.skipRows(2);
    
    {
      sheet.row().start();
      sheet.cellStr(1, "Это заголовок группы");
      sheet.cellStr(2, "Данные в заголовке");
      sheet.cellStr(3, "Бла бла бла");
      sheet.cellInt(4, 34555);
      sheet.row().finish();
      
      for (int i = 0; i < 5; i++) {
        sheet.row().outline().hidden().start();
        sheet.cellStr(1, "Это свёрнутый элемент");
        sheet.cellStr(2, "Данные в элементе");
        sheet.cellStr(3, "Фу фу фу " + i);
        sheet.cellInt(4, 222 + i);
        sheet.row().finish();
      }
      
    }
    {
      sheet.row().start();
      sheet.cellStr(1, "Это заголовок ещё одной группы группы");
      sheet.cellStr(2, "Данные в заголовке");
      sheet.cellStr(3, "Бла бла бла ЫЫЫ");
      sheet.cellInt(4, 765);
      sheet.row().finish();
      
      for (int i = 0; i < 10; i++) {
        sheet.row().outline().hidden().start();
        sheet.cellStr(1, "Это свёрнутый элемент");
        sheet.cellStr(2, "Данные в элементе");
        sheet.cellStr(3, "МААМАМА " + i);
        sheet.cellInt(4, 765 + i);
        sheet.row().finish();
      }
      
    }
    
    {
      sheet.row().start();
      sheet.cellStr(1, "Это заголовок группы");
      sheet.cellStr(2, "Данные в заголовке");
      sheet.cellStr(3, "Бла бла бла");
      sheet.cellInt(4, 34555);
      sheet.row().finish();
      
      for (int i = 0; i < 5; i++) {
        sheet.row().outline().hidden().start();
        sheet.cellStr(1, "Это свёрнутый элемент");
        sheet.cellStr(2, "Данные в элементе");
        sheet.cellStr(3, "Фу фу фу " + i);
        sheet.cellInt(4, 222 + i);
        sheet.row().finish();
      }
    }
  }
  
}
