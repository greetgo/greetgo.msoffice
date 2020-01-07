package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.xlsx.reader.model.Border;
import kz.greetgo.msoffice.xlsx.reader.model.Border4;
import kz.greetgo.msoffice.xlsx.reader.model.BorderSide;
import kz.greetgo.msoffice.xlsx.reader.model.BorderStyle;
import kz.greetgo.msoffice.xlsx.reader.model.CellStyleXf;
import kz.greetgo.msoffice.xlsx.reader.model.CellXf;
import kz.greetgo.msoffice.xlsx.reader.model.FontData;
import kz.greetgo.msoffice.xlsx.reader.model.HorizontalAlign;
import kz.greetgo.msoffice.xlsx.reader.model.NumFmtData;
import kz.greetgo.msoffice.xlsx.reader.model.StylesData;
import kz.greetgo.msoffice.xlsx.reader.model.VerticalAlign;
import org.xml.sax.Attributes;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static kz.greetgo.msoffice.util.UtilOffice.strToInt;

public class StylesHandler extends AbstractXmlHandler {

  private final StylesData styles;

  public StylesHandler(StylesData styles) {
    this.styles = styles;
  }

  private FontData font = null;
  private Border4 border4 = null;
  private CellXf cellXf = null;

  @SuppressWarnings("FieldCanBeLocal")
  private CellStyleXf cellStyleXf = null;

  private static final Pattern PATH_BORDER = Pattern.compile(
    "/styleSheet/borders/border/(left|right|top|bottom|diagonal)");

  @Override
  protected void startTag(String tagPath, Attributes attributes) {

    if ("/styleSheet/fonts/font".equals(tagPath)) {
      font = styles.newFont();
      return;
    }

    if ("/styleSheet/fonts/font/sz".equals(tagPath)) {
      font.size = Integer.valueOf(attributes.getValue("val"));
      return;
    }

    if ("/styleSheet/fonts/font/name".equals(tagPath)) {
      font.name = attributes.getValue("val");
      return;
    }

    if ("/styleSheet/fonts/font/b".equals(tagPath)) {
      font.bold = true;
      return;
    }

    if ("/styleSheet/borders/border".equals(tagPath)) {
      border4 = styles.newBorder4();
      return;
    }

    {
      Matcher matcher = PATH_BORDER.matcher(tagPath);
      if (matcher.matches()) {
        Border border = border4.get(BorderSide.valueOf(matcher.group(1)));
        border.style = BorderStyle.from(attributes.getValue("style")).orElse(null);
        return;
      }
    }

    if ("/styleSheet/cellStyleXfs/xf".equals(tagPath)) {
      cellStyleXf = styles.newCellStyleXf();
      cellStyleXf.numFmtId = strToInt(attributes.getValue("numFmtId"));
      cellStyleXf.fontId = strToInt(attributes.getValue("fontId"));
      cellStyleXf.fillId = strToInt(attributes.getValue("fillId"));
      cellStyleXf.borderId = strToInt(attributes.getValue("borderId"));
      return;
    }

    if ("/styleSheet/cellXfs/xf".equals(tagPath)) {
      cellXf = styles.newCellXf();
      cellXf.numFmtId = "1".equals(attributes.getValue("applyNumberFormat"))
        ? strToInt(attributes.getValue("numFmtId"))
        : null;
      cellXf.fontId = strToInt(attributes.getValue("fontId"));
      cellXf.fillId = strToInt(attributes.getValue("fillId"));
      cellXf.borderId = strToInt(attributes.getValue("borderId"));
      cellXf.xfId = strToInt(attributes.getValue("xfId"));
      return;
    }

    if ("/styleSheet/cellXfs/xf/alignment".equals(tagPath)) {
      cellXf.verticalAlign = VerticalAlign.valueOrNull(attributes.getValue("vertical"));
      cellXf.horizontalAlign = HorizontalAlign.valueOrNull(attributes.getValue("horizontal"));
      return;
    }

    //noinspection SpellCheckingInspection
    if ("/styleSheet/numFmts/numFmt".equals(tagPath)) {
      NumFmtData numFmt = new NumFmtData(attributes.getValue("numFmtId"));
      numFmt.formatCode = attributes.getValue("formatCode");
      styles.numFmtDataIdMap.put(numFmt.id, numFmt);
      return;
    }

//    System.out.println("j235v6 :: styles tagPath = " + tagPath + " " + XmlUtil.toStr(attributes));

  }

  @Override
  protected void endTag(String tagPath) {}
}
