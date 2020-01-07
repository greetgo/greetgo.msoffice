package kz.greetgo.msoffice.xlsx.gen;

import kz.greetgo.msoffice.util.UtilOffice;

public class Color {
  private int alfa, red, green, blue;

  public Color(int alfa, int red, int green, int blue) {
    this.alfa = alfa;
    this.red = red;
    this.green = green;
    this.blue = blue;
    checkValues();
  }

  public static Color copy(Color x) {
    if (x == null) return null;
    return new Color(x);
  }

  private void checkValues() {
    checkValue("alfa", alfa);
    checkValue("red", red);
    checkValue("green", green);
    checkValue("blue", blue);
  }

  private void checkValue(String valueName, int value) {
    if (value < 0 || value > 255) {
      throw new IllegalArgumentException("Must be 0 <= " + valueName + " <= 255, but " + valueName
        + " = " + value);
    }
  }

  public Color(int red, int green, int blue) {
    this.alfa = 0xFF;
    this.red = red;
    this.green = green;
    this.blue = blue;
    checkValues();
  }

  public Color(String hex) {
    if (hex.length() == 6) {
      alfa = 0xFF;
      red = UtilOffice.parsePartAsHex(hex, 0, 2);
      green = UtilOffice.parsePartAsHex(hex, 2, 2);
      blue = UtilOffice.parsePartAsHex(hex, 4, 2);
      checkValues();
      return;
    }
    if (hex.length() == 8) {
      alfa = UtilOffice.parsePartAsHex(hex, 0, 2);
      red = UtilOffice.parsePartAsHex(hex, 2, 2);
      green = UtilOffice.parsePartAsHex(hex, 4, 2);
      blue = UtilOffice.parsePartAsHex(hex, 6, 2);
      checkValues();
      return;
    }
    throw new IllegalArgumentException("Cannot read color from " + hex);
  }

  public Color(Color c) {
    alfa = c.alfa;
    red = c.red;
    green = c.green;
    blue = c.blue;
  }

  public Color(java.awt.Color awtColor) {
    alfa = awtColor.getAlpha();
    red = awtColor.getRed();
    green = awtColor.getGreen();
    blue = awtColor.getBlue();
  }

  private static final char[] DIGITS = new char[]{'0', '1', '2', '3', '4', '5', '6', '7', '8',
    '9', 'A', 'B', 'C', 'D', 'E', 'F'};

  public String str() {
    StringBuilder ret = new StringBuilder();

    ret.append(DIGITS[alfa / 16]);
    ret.append(DIGITS[alfa % 16]);

    ret.append(DIGITS[red / 16]);
    ret.append(DIGITS[red % 16]);

    ret.append(DIGITS[green / 16]);
    ret.append(DIGITS[green % 16]);

    ret.append(DIGITS[blue / 16]);
    ret.append(DIGITS[blue % 16]);

    return ret.toString();
  }

  public String strRGB() {
    StringBuilder ret = new StringBuilder();

    ret.append(DIGITS[red / 16]);
    ret.append(DIGITS[red % 16]);

    ret.append(DIGITS[green / 16]);
    ret.append(DIGITS[green % 16]);

    ret.append(DIGITS[blue / 16]);
    ret.append(DIGITS[blue % 16]);

    return ret.toString();
  }

  @Override
  public int hashCode() {
    final int prime = 31;
    int result = 1;
    result = prime * result + alfa;
    result = prime * result + blue;
    result = prime * result + green;
    result = prime * result + red;
    return result;
  }

  @Override
  public boolean equals(Object obj) {
    if (this == obj) return true;
    if (obj == null) return false;
    if (getClass() != obj.getClass()) return false;
    Color other = (Color) obj;
    if (alfa != other.alfa) return false;
    if (blue != other.blue) return false;
    if (green != other.green) return false;
    return red == other.red;
  }

  public int getAlfa() {
    return alfa;
  }

  public void setAlfa(int alfa) {
    this.alfa = alfa;
  }

  public int getRed() {
    return red;
  }

  public void setRed(int red) {
    this.red = red;
  }

  public int getGreen() {
    return green;
  }

  public void setGreen(int green) {
    this.green = green;
  }

  public int getBlue() {
    return blue;
  }

  public void setBlue(int blue) {
    this.blue = blue;
  }

  //some colors ...

  public static final Color white() {
    return new Color(java.awt.Color.white);
  }

  public static final Color lightGray() {
    return new Color(java.awt.Color.lightGray);
  }

  public static final Color gray() {
    return new Color(java.awt.Color.gray);
  }

  public static final Color darkGray() {
    return new Color(java.awt.Color.darkGray);
  }

  public static final Color black() {
    return new Color(java.awt.Color.black);
  }

  public static final Color red() {
    return new Color(java.awt.Color.red);
  }

  public static final Color pink() {
    return new Color(java.awt.Color.pink);
  }

  public static final Color orange() {
    return new Color(java.awt.Color.orange);
  }

  public static final Color yellow() {
    return new Color(java.awt.Color.yellow);
  }

  public static final Color green() {
    return new Color(java.awt.Color.green);
  }

  public static final Color magenta() {
    return new Color(java.awt.Color.magenta);
  }

  public static final Color cyan() {
    return new Color(java.awt.Color.cyan);
  }

  public static final Color blue() {
    return new Color(java.awt.Color.blue);
  }
}
