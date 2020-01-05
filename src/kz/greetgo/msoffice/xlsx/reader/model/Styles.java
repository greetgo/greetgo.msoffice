package kz.greetgo.msoffice.xlsx.reader.model;

import java.util.ArrayList;
import java.util.List;

public class Styles {

  public final List<FontData> fontList = new ArrayList<>();
  public final List<Border4> border4List = new ArrayList<>();
  public final List<StyleXf> xfList = new ArrayList<>();
  public final List<CellStyleXf> cellXfList = new ArrayList<>();

  public FontData newFont() {
    FontData fontData = new FontData();
    fontList.add(fontData);
    return fontData;
  }

  public Border4 newBorder4() {
    Border4 b = new Border4();
    border4List.add(b);
    return b;
  }

  public StyleXf newStyleXf() {
    StyleXf xf = new StyleXf();
    xfList.add(xf);
    return xf;
  }

  public CellStyleXf newCellXf() {
    CellStyleXf xf = new CellStyleXf();
    cellXfList.add(xf);
    return xf;
  }
}
