/**
 * 
 */
package kz.greetgo.msoffice.docx;

import java.io.PrintStream;
import java.util.HashMap;
import java.util.Map;

import kz.greetgo.msoffice.UtilOffice;

class RelationshipMap implements FilePart {
  private Map<String, Relationship> relationships = new HashMap<String, Relationship>();
  private final String partName;
  
  private RelationshipMap(String partName) {
    this.partName = partName;
  }
  
  static RelationshipMap createBySubjectPartName(String subjectPartName) {
    return new RelationshipMap("/word/_rels/" + UtilOffice.extractBaseName(subjectPartName)
        + ".rels");
  }
  
  static RelationshipMap createWithPartName(String partName) {
    return new RelationshipMap(partName);
  }
  
  @Override
  public String getPartName() {
    return partName;
  }
  
  public void put(Relationship r) {
    relationships.put(r.getTarget(), r);
  }
  
  @Override
  public void write(PrintStream out) {
    out.print("<?xml version=\"1.0\" encoding=\"UTF-8\"" + " standalone=\"yes\"?>\n");
    out.print("<Relationships" + " xmlns=\"http://schemas.openxmlformats.org/"
        + "package/2006/relationships\">");
    for (Relationship r : relationships.values()) {
      r.write(out);
    }
    out.print("</Relationships>");
  }
  
}
