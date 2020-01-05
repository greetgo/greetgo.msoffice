package kz.greetgo.msoffice.docx;

import kz.greetgo.msoffice.util.UtilOffice;

import java.io.PrintStream;

public class ImageElement implements RunElement {

  private final String id;
  private int cx, cy;

  private ImageElementPosition position = ImageElementPosition.INLINE;

  private int xOffset, yOffset;

  public ImageElement(String id) {
    this.id = id;
  }

  @Override
  public void write(PrintStream out) {

    String template = UtilOffice.streamToStr(ImageElement.class
      .getResourceAsStream("ImageElement.data.xml"));

    template = template.replace("[[CX]]", String.valueOf(getCx()));
    template = template.replace("[[CY]]", String.valueOf(getCy()));
    template = template.replace("[[ID]]", getId());

    if (ImageElementPosition.ANCHOR == position) {
      StringBuffer posExt = new StringBuffer();
      posExt.append("<wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"page\"><wp:posOffset>");
      posExt.append(xOffset);
      posExt.append("</wp:posOffset></wp:positionH><wp:positionV relativeFrom=\"page\"><wp:posOffset>");
      posExt.append(yOffset);
      posExt.append("</wp:posOffset></wp:positionV>");

      template = template.replace("[[POSITION]]", position.name().toLowerCase());
      template = template.replace("[[POSITION_PARAMS]]",
        "simplePos=\"0\" relativeHeight=\"251658240\" behindDoc=\"1\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\"");
      template = template.replace("[[POSITION_EXT_1]]", posExt.toString());
      template = template.replace("[[POSITION_EXT_2]]", "<wp:wrapNone/>");
    } else {
      template = template.replace("[[POSITION]]", ImageElementPosition.INLINE.name().toLowerCase());
      template = template.replace("[[POSITION_PARAMS]]", "");
      template = template.replace("[[POSITION_EXT_1]]", "");
      template = template.replace("[[POSITION_EXT_2]]", "");
    }

    out.print(template);
  }

  public String getId() {
    return id;
  }

  public void setCy(int cy) {
    this.cy = cy;
  }

  public int getCy() {
    return cy;
  }

  public void setCx(int cx) {
    this.cx = cx;
  }

  public int getCx() {
    return cx;
  }

  public ImageElementPosition getPosition() {
    return position;
  }

  /**
   * Установить
   *
   * @param position положение изображения по отношению к тексту
   */
  public void setPosition(ImageElementPosition position) {
    this.position = position;
  }

  public int getXOffset() {
    return xOffset;
  }

  /**
   * Установить
   *
   * @param xOffset отступ от левого края страницы (по горизонтали)
   */
  public void setXOffset(int xOffset) {
    this.xOffset = xOffset;
  }

  public int getYOffset() {
    return yOffset;
  }

  /**
   * Установить
   *
   * @param yOffset отступ от верхнего края страницы (по вертикали)
   */
  public void setYOffset(int yOffset) {
    this.yOffset = yOffset;
  }
}
