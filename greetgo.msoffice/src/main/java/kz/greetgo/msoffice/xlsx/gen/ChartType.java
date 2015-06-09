package kz.greetgo.msoffice.xlsx.gen;

public enum ChartType {
  
  CIRCLE_DIAGRAM("pie3DChart"), // круговая диаграмма
  LINE_CHART("lineChart"); // график
  
  final String tag;
  
  private ChartType(String tag) {
	  this.tag = tag;
  }
}