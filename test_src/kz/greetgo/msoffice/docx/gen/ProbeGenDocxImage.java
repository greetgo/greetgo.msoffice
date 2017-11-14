package kz.greetgo.msoffice.docx.gen;

import kz.greetgo.msoffice.docx.Align;
import kz.greetgo.msoffice.docx.Document;
import kz.greetgo.msoffice.docx.Docx;
import kz.greetgo.msoffice.docx.ImageElement;
import kz.greetgo.msoffice.docx.ImageElementPosition;
import kz.greetgo.msoffice.docx.InputSource;
import kz.greetgo.msoffice.docx.Para;
import kz.greetgo.msoffice.docx.Run;

import java.io.File;
import java.io.InputStream;

public class ProbeGenDocxImage {

  public static void main(String[] args) {

    Docx docx = new Docx();

    docx.getDocument().setLeft(2000);
    docx.getDocument().setRight(2000);

    Document doc = docx.getDocument();

    {
      Para para = doc.createPara();
      para.setAlign(Align.CENTER);

      {
        Run runImage = para.createRun();
        ImageElement image = runImage.createImage(new InputSource() {

          @Override
          public InputStream openInputStream() throws Exception {
            return ProbeGenDocxImage.class.getResourceAsStream("a_horse.jpg");
          }
        });

        image.setCx(2_000_000);
        image.setCy(1_500_000);

        image.setPosition(ImageElementPosition.ANCHOR);
        image.setXOffset(5_200_000);
        image.setYOffset(200_000);

        runImage
          .addText("Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст Текст ");
      }

      {
        Run частьАбзаца = para.createRun();
        частьАбзаца.setBold(true);
        частьАбзаца.setTextSize(70);
        частьАбзаца.addText("Текст");
      }
    }

    {
      Para para = doc.createPara();
      para.setAlign(Align.CENTER);
      {
        Run частьАбзаца = para.createRun();
        частьАбзаца.setBold(true);
        частьАбзаца.setTextSize(70);
      }
    }

    new File("build").mkdirs();
    docx.write("build/example-gen-image.docx");

    System.out.println("OK");
  }
}