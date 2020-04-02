package kz.greetgo.msoffice.docx.gen;

import kz.greetgo.msoffice.PageOrientation;
import kz.greetgo.msoffice.docx.Align;
import kz.greetgo.msoffice.docx.Document;
import kz.greetgo.msoffice.docx.Docx;
import kz.greetgo.msoffice.docx.LineStyle;
import kz.greetgo.msoffice.docx.Para;
import kz.greetgo.msoffice.docx.Run;
import kz.greetgo.msoffice.docx.Table;
import kz.greetgo.msoffice.docx.TableCell;
import kz.greetgo.msoffice.docx.TableCol;
import kz.greetgo.msoffice.docx.TableRow;
import kz.greetgo.msoffice.docx.Underline;

public class ProbeGenDocx {
  public static void main(String[] args) {
    Docx docx = new Docx();
    
    docx.getDocument().setPageOrientation(PageOrientation.LANDSCAPE);
    docx.getDocument().setLeft(2000);
    docx.getDocument().setRight(2000);
    
    Document doc = docx.getDocument();
    {
      Para para = doc.createPara();
      
      {
        Run частьАбзаца = para.createRun();
        частьАбзаца.setFontName("Courier New");
        частьАбзаца.setUnderline(Underline.DOUBLE);
        частьАбзаца.addText("Вот часть абзаца дважды подчёркнутая,");
      }
      {
        Run частьАбзаца = para.createRun();
        частьАбзаца.setFontName("Courier New");
        частьАбзаца.addText("а вот не подчёркнутая, ");
      }
      {
        Run частьАбзаца = para.createRun();
        частьАбзаца.setBold(true);
        частьАбзаца.addText("а вот жирная.");
      }
    }
    
    {
      Para para = doc.createPara();
      para.setAlign(Align.RIGHT);
      {
        Run частьАбзаца = para.createRun();
        частьАбзаца.addText("Вот ещё один абзац");
      }
      {
        Run частьАбзаца = para.createRun();
        частьАбзаца.setItalic(true);
        частьАбзаца.setBold(true);
        частьАбзаца.addText(" выровненный справа.");
      }
    }
    
    doc.createPara().createRun().addText("");
    doc.createPara().createRun().addText("");
    
    {
      Para para = doc.createPara();
      para.setAlign(Align.CENTER);
      {
        Run частьАбзаца = para.createRun();
        частьАбзаца.setBold(true);
        частьАбзаца.setTextSize(70);
        частьАбзаца.addText("Вот так!");
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
    {
      Table table = doc.createTable();
      table.setTableWidth(9540);
      table.setTblCellSpacing(0);
      table.getMargins().setWidth(105);
      table.setTableLookVal("04A0");
      {
        TableRow row = table.createRow();
        row.setTblCellSpacing(0);
        {
          TableCol col1 = table.createCol();
          col1.setWidth(4601);
          
          TableCell cell1 = row.createCell();
          cell1.setWidth(9300);
          cell1.setGridSpan(3);
          cellStyle(cell1);
          cell1.getShd().setVal("clear");
          cell1.getShd().setFill("C0C0C0");
          Para para = cell1.createPara();
          para.setAlign(Align.BOTH);
          Run title = para.createRun();
          title.setTextSize(13);
          title.setFontName("Arial");
          for (int i = 0; i < 10; i++) {
            title.addText("Краткая информация по Заемщику/Гаранту и проекту " + i);
          }
          
        }
      }
    }
    {
      Table table = doc.createTable();
      table.setTableWidth(9540);
      table.setTblCellSpacing(0);
      table.getMargins().setWidth(105);
      table.setTableLookVal("04A0");
      
      TableCol col1 = table.createCol();
      col1.setWidth(4601 / 3);
      TableCol col2 = table.createCol();
      col2.setWidth(4601 / 3);
      TableCol col3 = table.createCol();
      col3.setWidth(4601 / 3);
      
      {
        TableRow row = table.createRow();
        row.setTblCellSpacing(0);
        
        {
          TableCell cell = row.createCell();
          cellStyle(cell);
          setCellText(cell, "Текст уау 1");
        }
        {
          TableCell cell = row.createCell();
          cellStyle(cell);
          setCellText(cell, "Текст уау 2");
        }
        {
          TableCell cell = row.createCell();
          cellStyle(cell);
          setCellText(cell, "Текст уау 3 fdsf");
        }
      }
      {
        TableRow row = table.createRow();
        row.setTblCellSpacing(0);
        
        {
          TableCell cell = row.createCell();
          cellStyle(cell);
          setCellText(cell, "Текст SSS 1");
        }
        {
          TableCell cell = row.createCell();
          cellStyle(cell);
          setCellText(cell, "Текст SSS 2");
        }
        {
          TableCell cell = row.createCell();
          cellStyle(cell);
          cell.mergeRestart();
          setCellText(cell, "Текст SSS 3 fdsf");
        }
      }
      {
        TableRow row = table.createRow();
        row.setTblCellSpacing(0);
        
        {
          TableCell cell = row.createCell();
          cellStyle(cell);
          setCellText(cell, "Текст SDF 1");
        }
        {
          TableCell cell = row.createCell();
          cellStyle(cell);
          setCellText(cell, "Текст SDF 2");
        }
        {
          TableCell cell = row.createCell();
          cellStyle(cell);
          cell.mergeCotinue();
        }
      }
    }
    
    docx.write("build/example-gen2.docx");
    
    System.out.println("OK");
  }
  
  private static void setCellText(TableCell cell, String text) {
    Para para = cell.createPara();
    para.setAlign(Align.BOTH);
    Run title = para.createRun();
    title.setTextSize(13);
    title.setFontName("Arial");
    
    title.addText(text);
  }
  
  private static void cellStyle(TableCell cell) {
    cell.getBorders().setLineStyle(LineStyle.SINGLE);
    cell.getBorders().setSize(4);
    cell.getBorders().getInsideH().setSize(0);
    cell.getBorders().getInsideV().setSize(0);
    cell.getMargins().setWidth(108);
    cell.getMargins().getTop().setWidth(0);
    cell.getMargins().getBottom().setWidth(0);
  }
}
