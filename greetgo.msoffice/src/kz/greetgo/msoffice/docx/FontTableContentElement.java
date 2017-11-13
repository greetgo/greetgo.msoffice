package kz.greetgo.msoffice.docx;

import java.io.PrintStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

public class FontTableContentElement implements ContentElement {
  private final String partName;
  
  private FontTableContentElement(String partName) {
    this.partName = partName;
    assert partName != null;
  }
  
  public static FontTableContentElement createDefault(String partName) {
    FontTableContentElement ret = new FontTableContentElement(partName);
    ret.initDefault();
    return ret;
  }
  
  public static class FontElement implements XmlWriter {
    private final String name;
    private List<String> xmlContent = new ArrayList<String>();
    private Font font = null;
    
    public FontElement(String name, Iterator<String> xmlContentIterator) {
      this.name = name;
      while (xmlContentIterator.hasNext()) {
        xmlContent.add(xmlContentIterator.next());
      }
    }
    
    public Font getFont() {
      if (font == null) {
        font = new Font();
        font.name = name;
      }
      return font;
    }
    
    @Override
    public void write(PrintStream out) {
      out.print("<w:font w:name=\"" + getName() + "\">");
      for (String s : xmlContent) {
        out.print(s);
      }
      out.print("</w:font>");
    }
    
    public String getName() {
      return name;
    }
  }
  
  private Map<String, FontElement> elements = new HashMap<String, FontElement>();
  
  public Font getFont(String name) {
    if (!elements.containsKey(name)) {
      throw new IllegalArgumentException("No font with name = " + name);
    }
    return elements.get(name).getFont();
  }
  
  private void initDefault() {
    List<String> c = new LinkedList<String>();
    {
      String name = "Calibri";
      c.clear();
      c.add("<w:panose1 w:val=\"020F0502020204030204\" />");
      c.add("<w:charset w:val=\"CC\" />");
      c.add("<w:family w:val=\"swiss\" />");
      c.add("<w:pitch w:val=\"variable\" />");
      c.add("<w:sig w:usb0=\"A00002EF\" w:usb1=\"4000207B\""
          + " w:usb2=\"00000000\" w:usb3=\"00000000\""
          + " w:csb0=\"0000009F\" w:csb1=\"00000000\" />");
      elements.put(name, new FontElement(name, c.iterator()));
    }
    {
      String name = "Times New Roman";
      c.clear();
      c.add("<w:panose1 w:val=\"02020603050405020304\" />");
      c.add("<w:charset w:val=\"CC\" />");
      c.add("<w:family w:val=\"roman\" />");
      c.add("<w:pitch w:val=\"variable\" />");
      c.add("<w:sig w:usb0=\"20002A87\" w:usb1=\"80000000\""
          + " w:usb2=\"00000008\" w:usb3=\"00000000\""
          + " w:csb0=\"000001FF\" w:csb1=\"00000000\" />");
      elements.put(name, new FontElement(name, c.iterator()));
    }
    {
      String name = "Tahoma";
      c.clear();
      c.add("<w:panose1 w:val=\"020B0604030504040204\" />");
      c.add("<w:charset w:val=\"CC\" />");
      c.add("<w:family w:val=\"swiss\" />");
      c.add("<w:pitch w:val=\"variable\" />");
      c.add("<w:sig w:usb0=\"61002A87\" w:usb1=\"80000000\""
          + " w:usb2=\"00000008\" w:usb3=\"00000000\""
          + " w:csb0=\"000101FF\" w:csb1=\"00000000\" />");
      elements.put(name, new FontElement(name, c.iterator()));
    }
    {
      String name = "Cambria";
      c.clear();
      c.add("<w:panose1 w:val=\"02040503050406030204\" />");
      c.add("<w:charset w:val=\"CC\" />");
      c.add("<w:family w:val=\"roman\" />");
      c.add("<w:pitch w:val=\"variable\" />");
      c.add("<w:sig w:usb0=\"A00002EF\" w:usb1=\"4000004B\""
          + " w:usb2=\"00000000\" w:usb3=\"00000000\""
          + " w:csb0=\"0000009F\" w:csb1=\"00000000\" />");
      elements.put(name, new FontElement(name, c.iterator()));
    }
    {
      String name = "Arial";
      c.clear();
      c.add(" <w:panose1 w:val=\"020B0604020202020204\" />");
      c.add("<w:charset w:val=\"CC\" />");
      c.add("<w:family w:val=\"swiss\" />");
      c.add("<w:pitch w:val=\"variable\" />");
      c.add("<w:sig w:usb0=\"20002A87\" w:usb1=\"80000000\""
          + " w:usb2=\"00000008\" w:usb3=\"00000000\" "
          + " w:csb0=\"000001FF\" w:csb1=\"00000000\" />");
      elements.put(name, new FontElement(name, c.iterator()));
    }
  }
  
  @Override
  public String getPartName() {
    return partName;
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<?xml version=\"1.0\" encoding=\"UTF-8\"" + " standalone=\"yes\"?>\n");
    out.print("<w:fonts xmlns:r=\"http://schemas.openxmlformats.org/"
        + "officeDocument/2006/relationships\"" + " xmlns:w=\"http://schemas.openxmlformats.org/"
        + "wordprocessingml/2006/main\">");
    for (FontElement e : elements.values()) {
      e.write(out);
    }
    out.print("</w:fonts>");
  }
  
  @Override
  public ContentType getContentType() {
    return ContentType.FONT_TABLE;
  }
  
}
