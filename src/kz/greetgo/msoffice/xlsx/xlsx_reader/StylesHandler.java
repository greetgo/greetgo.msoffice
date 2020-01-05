package kz.greetgo.msoffice.xlsx.xlsx_reader;

import kz.greetgo.msoffice.xlsx.xlsx_reader.model.Border;
import kz.greetgo.msoffice.xlsx.xlsx_reader.model.Border4;
import kz.greetgo.msoffice.xlsx.xlsx_reader.model.BorderSide;
import kz.greetgo.msoffice.xlsx.xlsx_reader.model.BorderStyle;
import kz.greetgo.msoffice.xlsx.xlsx_reader.model.CellStyleXf;
import kz.greetgo.msoffice.xlsx.xlsx_reader.model.FontData;
import kz.greetgo.msoffice.xlsx.xlsx_reader.model.HorizontalAlign;
import kz.greetgo.msoffice.xlsx.xlsx_reader.model.StyleXf;
import kz.greetgo.msoffice.xlsx.xlsx_reader.model.Styles;
import kz.greetgo.msoffice.xlsx.xlsx_reader.model.VerticalAlign;
import org.xml.sax.Attributes;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static kz.greetgo.msoffice.util.UtilOffice.strToInt;

public class StylesHandler extends AbstractXmlHandler {

  private final Styles styles = new Styles();

  private FontData font;
  private Border4 border4;
  @SuppressWarnings("FieldCanBeLocal")
  private StyleXf xf;
  private CellStyleXf cellXf;

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
      xf = styles.newStyleXf();
      xf.numFmtId = strToInt(attributes.getValue("numFmtId"));
      xf.fontId = strToInt(attributes.getValue("fontId"));
      xf.fillId = strToInt(attributes.getValue("fillId"));
      xf.borderId = strToInt(attributes.getValue("borderId"));
      return;
    }

    if ("/styleSheet/cellXfs/xf".equals(tagPath)) {
      cellXf = styles.newCellXf();
      cellXf.numFmtId = strToInt(attributes.getValue("numFmtId"));
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

  }

  @Override
  protected void endTag(String tagPath) {}
}
