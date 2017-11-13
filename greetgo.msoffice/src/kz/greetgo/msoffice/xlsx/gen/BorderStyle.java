package kz.greetgo.msoffice.xlsx.gen;

/**
 * Стил бордюра
 * 
 * @author pompei
 */
public enum BorderStyle {
  hair, dotted, dashDotDot, dashDot, dashed, thin, mediumDashDotDot, slantDashDot, mediumDashDot,
  mediumDashed, medium, thick, _double;
  
  public String str() {
    if (name().startsWith("_")) return name().substring(1);
    return name();
  }
}
