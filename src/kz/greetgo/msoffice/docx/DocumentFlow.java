package kz.greetgo.msoffice.docx;

import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

public abstract class DocumentFlow implements ReferencableContentElement {
  private List<FlowElement> elements = new ArrayList<FlowElement>();
  private final MSHelper msHelper;
  private final String partName;
  
  @Override
  public String getPartName() {
    return partName;
  }
  
  public DocumentFlow(String partName, MSHelper msHelper) {
    this.partName = partName;
    this.msHelper = msHelper;
  }
  
  public Para createPara() {
    return createPara(null);
  }
  
  public Para createPara(Para example) {
    Para ret = new Para(partName, msHelper);
    if (example != null) {
      ret.copyDecorationFrom(example);
    } else {
      Para defaultPara = msHelper.getDefaultPara();
      if (defaultPara != null) {
        ret.copyDecorationFrom(defaultPara);
      }
    }
    elements.add(ret);
    return ret;
  }
  
  public Table createTable() {
    Table ret = new Table(partName, msHelper);
    elements.add(ret);
    return ret;
  }
  
  @Override
  public void write(PrintStream out) {
    writeTopXml(out);
    writeFlow(out);
    writeBottomXml(out);
  }
  
  private void writeFlow(PrintStream out) {
    for (FlowElement e : elements) {
      e.write(out);
    }
  }
  
  @Override
  public void writeReferences(PrintStream out) {
    for (Reference r : references) {
      r.write(out);
    }
  }
  
  private List<Reference> references = new ArrayList<Reference>();
  
  @Override
  public void addReference(Reference reference) {
    references.add(reference);
  }
  
  protected abstract void writeBottomXml(PrintStream out);
  
  protected abstract void writeTopXml(PrintStream out);
  
  protected MSHelper getMsHelper() {
    return msHelper;
  }
  
  /**
   * Создаёт параграф для примера. Этот параграф не используется для оттображения в документе, но
   * может служить для создания реальных параграфов как пример оформления
   * 
   * @param anotherExample
   *          другой пример параграфа, или null, для значений по умолчанию
   * @return параграф-пример
   */
  public Para createExamplePara(Para anotherExample) {
    Para ret = new Para("", null);
    if (anotherExample != null) {
      ret.copyDecorationFrom(anotherExample);
    } else if (msHelper != null) {
      Para dp = msHelper.getDefaultPara();
      if (dp != null) {
        ret.copyDecorationFrom(dp);
      }
    }
    return ret;
    
  }
}
