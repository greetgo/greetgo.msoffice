package kz.greetgo.msoffice.xlsx.probes;

import java.io.FileOutputStream;

import kz.greetgo.msoffice.xlsx.gen.Align;
import kz.greetgo.msoffice.xlsx.gen.Chart;
import kz.greetgo.msoffice.xlsx.gen.ChartType;
import kz.greetgo.msoffice.xlsx.gen.Color;
import kz.greetgo.msoffice.xlsx.gen.Sheet;
import kz.greetgo.msoffice.xlsx.gen.Xlsx;

public class GenXlsxGraphProbe {
  
  public static void main(String[] args) throws Exception {
    
    Xlsx excel = new Xlsx();
    
    Sheet sh = excel.newSheet(true);
    
    sh.setWidth(1, 10);
    sh.setWidth(2, 6);
    sh.setWidth(3, 10);
    
    // данные для графика
    
    sh.skipRow();
    
    sh.row().start();
    sh.cellStr(2, "Oracle");
    sh.cellInt(3, 1000);
    sh.cellInt(4, 1400);
    sh.row().finish();
    
    sh.row().start();
    sh.cellStr(2, "PostgreSQL");
    sh.cellInt(3, 1200);
    sh.cellInt(4, 1800);
    sh.row().finish();
    
    sh.row().start();
    sh.cellStr(2, "MariaDB");
    sh.cellInt(3, 800);
    sh.cellInt(4, 600);
    sh.row().finish();
    
    sh.row().start();
    sh.cellStr(2, "SQL Server");
    sh.cellInt(3, 200);
    sh.cellInt(4, 20);
    sh.row().finish();
    
    sh.row().start();
    sh.cellStr(2, "DB2");
    sh.cellInt(3, 700);
    sh.cellInt(4, 750);
    sh.row().finish();
    
    // графики
    
    Chart chart = sh.addChart(ChartType.CIRCLE_DIAGRAM, "F", 2, "K", 10);
    chart.setTitle("Годовой баланс");
    chart.setData("Лист1!$C$2:$C$5");
    chart.setDataTitles("Лист1!$B$2:$B$5");
    chart.setExplosion(10);
    
    chart = sh.addChart(ChartType.LINE_CHART, 5, 12, 10, 20);
    chart.setTitle("График");
    
    chart.setData(sh, "C", 2, 6);
    chart.setDataTitles(sh, "B", 2, 6);
    chart.setColorLine(Color.blue());
    
    chart.setData2("Лист1!$D$2:$D$6");
    chart.setDataTitles2("Лист1!$B$2:$B$6");
    chart.setColorLine(Color.red());
    chart.setWidthLine(20);
    
    chart.setColorBackground(new Color(240, 240, 240));
    chart.setAlignmentOY(Align.left);
    
    Sheet sh2 = excel.newSheet(false);
    
    sh2.row().start();
    sh2.cellStr(2, "График");
    sh2.row().finish();
    
    chart = sh2.addChart(ChartType.CIRCLE_DIAGRAM, 2, 2, 10, 10);
    chart.setData("Лист1!$C$2:$C$5");
    chart.setDataTitles("Лист1!$B$2:$B$5");
    chart.setAlignmentLegend(Align.right);
    chart.setRotation(60);
    
    // запись
    
    FileOutputStream os = new FileOutputStream("/home/aboldyrev/tmp/vb/graph-probe.xlsx");
    excel.complete(os);
    os.flush();
    os.close();
    
    System.out.println("Файл excel создан.");
  }
}