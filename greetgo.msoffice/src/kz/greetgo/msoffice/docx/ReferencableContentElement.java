package kz.greetgo.msoffice.docx;

import java.io.PrintStream;

public interface ReferencableContentElement extends ContentElement {
  void writeReferences(PrintStream out);
  
  void addReference(Reference reference);
}
