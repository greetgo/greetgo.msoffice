package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.xlsx.reader.model.WorkbookData;
import org.xml.sax.Attributes;

public class WorkbookHandler extends AbstractXmlHandler {
  private final WorkbookData workbook;

  public WorkbookHandler(WorkbookData workbook) {
    this.workbook = workbook;
  }

  @Override
  protected void startTag(String tagPath, Attributes attributes) {

    if ("/workbook/sheets/sheet".equals(tagPath)) {
      String name = attributes.getValue("name");
      int index = workbook.sheetRefList.size();
      //int id = Integer.parseInt(attributes.getValue("sheetId"));
      int id = index + 1;
      workbook.sheetRefList.add(new SheetRef(index, id, name));
      return;
    }

//    System.out.println("4315v4 :: tagPath = " + tagPath + " " + XmlUtil.toStr(attributes));
  }

  @Override
  protected void endTag(String tagPath) {}
}
