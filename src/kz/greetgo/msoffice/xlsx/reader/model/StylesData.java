package kz.greetgo.msoffice.xlsx.reader.model;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class StylesData {

  public final List<FontData> fontList = new ArrayList<>();
  public final List<Border4> border4List = new ArrayList<>();
  public final List<CellXf> cellXfList = new ArrayList<>();
  public final List<CellStyleXf> cellStyleXfList = new ArrayList<>();
  public final Map<Integer, NumFmtData> numFmtDataIdMap = new HashMap<>();

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

  public CellXf newCellXf() {
    CellXf xf = new CellXf();
    cellXfList.add(xf);
    return xf;
  }

  public CellStyleXf newCellStyleXf() {
    CellStyleXf xf = new CellStyleXf();
    cellStyleXfList.add(xf);
    return xf;
  }

  public CellXf getCellXf(int style) {
    return cellXfList.get(style);
  }
}
