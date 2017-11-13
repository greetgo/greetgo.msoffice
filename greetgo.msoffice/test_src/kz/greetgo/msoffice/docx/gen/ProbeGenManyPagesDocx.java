package kz.greetgo.msoffice.docx.gen;

import java.io.File;

import kz.greetgo.msoffice.docx.Align;
import kz.greetgo.msoffice.docx.Document;
import kz.greetgo.msoffice.docx.DocumentFooter;
import kz.greetgo.msoffice.docx.DocumentHeader;
import kz.greetgo.msoffice.docx.Docx;
import kz.greetgo.msoffice.docx.Para;
import kz.greetgo.msoffice.docx.Run;

public class ProbeGenManyPagesDocx {
  public static void main(String[] args) {
    Docx docx = new Docx();
    
    Document doc = docx.getDocument();
    for (int i = 1; i <= 1000; i++) {
      Para para = doc.createPara();
      {
        Run частьАбзаца = para.createRun();
        частьАбзаца.addText("Какой-то текст.");
      }
      {
        Run частьАбзаца = para.createRun();
        частьАбзаца.addText(" 在中國這個文本.");
      }
      {
        Run частьАбзаца = para.createRun();
        частьАбзаца.addText(" Это текст под номером " + i + ".");
      }
    }
    
    {
      DocumentFooter footer = docx.getOrCreateFooter();
      {
        Para para = footer.createPara();
        para.setAlign(Align.RIGHT);//Делаем чтобы страница отображалась справа
        {
          Run частьАбзаца = para.createRun();
          частьАбзаца.addText("Это страница номер ");
          частьАбзаца.field().addPageNomer();
          частьАбзаца.addText(". Всего страниц - ");
          частьАбзаца.field().addPageCount();
          частьАбзаца.addText(".");
        }
      }
    }
    
    DocumentHeader header = docx.getOrCreateHeader();
    {
      Para para = header.createPara();
      para.setAlign(Align.CENTER);//Пусть отображается по середине
      {
        Run частьАбзаца = para.createRun();
        частьАбзаца.addText("Hello from header - ");
        частьАбзаца.field().addPageNomer();//И в заголовке тоже можно отображать стрницы
        частьАбзаца.addText("/");
        частьАбзаца.field().addPageCount();
      }
    }
    
    new File("build").mkdirs();
    docx.write("build/ManyPages.docx");
    
    System.out.println("OK");
  }
}
