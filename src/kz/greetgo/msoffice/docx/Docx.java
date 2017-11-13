package kz.greetgo.msoffice.docx;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import kz.greetgo.msoffice.UtilOffice;

/**
 * Главный класс для создания файлов формата docx
 * 
 * @author pompei
 * 
 */
public class Docx {
  private Map<String, RelationshipMap> relationShipMaps = new HashMap<String, RelationshipMap>();
  private int nextRelationshipId = 1;
  
  private boolean mainRelationsInitiated = false;
  
  public Document getDocument() {
    checkInit();
    return getContent().getDocument();
  }
  
  private void checkInit() {
    if (!mainRelationsInitiated) {
      initMainRelations();
      mainRelationsInitiated = true;
    }
  }
  
  private void initMainRelations() {
    RelationshipMap mainRSet = RelationshipMap.createWithPartName("/_rels/.rels");
    relationShipMaps.put(mainRSet.getPartName(), mainRSet);
    
    {
      Relationship r = new Relationship();
      r.setType(Relationship.Type.OFFICE_DOCUMENT);
      r.setId("rId" + nextRelationshipId++);
      r.setTarget("word/document.xml");
      
      mainRSet.put(r);
    }
    {
      Relationship r = new Relationship();
      r.setType(Relationship.Type.CORE_PROPERTIES);
      r.setId("rId" + nextRelationshipId++);
      r.setTarget("docProps/core.xml");
      
      mainRSet.put(r);
    }
    {
      Relationship r = new Relationship();
      r.setType(Relationship.Type.EXTENDED_PROPERTIES);
      r.setId("rId" + nextRelationshipId++);
      r.setTarget("docProps/app.xml");
      
      mainRSet.put(r);
    }
  }
  
  private int nextImageNumber = 1;
  
  private MSHelper msHelper = new MSHelper() {
    @Override
    public Relationship createRelationshipForImage(String ownerPartName,
        final InputSource inputSource) {
      getContent().checkExistsDefaultImagePng();
      
      final Relationship r = new Relationship();
      r.setType(Relationship.Type.IMAGE);
      r.setId("rId" + nextRelationshipId++);
      r.setTarget("media/image" + nextImageNumber++ + ".png");
      
      getRelationshipMap(ownerPartName).put(r);
      
      binaryParts.add(new BinaryFilePart() {
        @Override
        public InputStream openInputStream() throws Exception {
          return inputSource.openInputStream();
        }
        
        @Override
        public String getPartName() {
          return "/word/" + r.getTarget();
        }
      });
      
      return r;
    }
    
    @Override
    public Font getFont(String name) {
      return getFontTableContentElement().getFont(name);
    }
    
    @Override
    public Para getDefaultPara() {
      return defaultPara;
    }
  };
  
  private FontTableContentElement fontTableContentElement = null;
  
  private FontTableContentElement getFontTableContentElement() {
    if (fontTableContentElement == null) {
      fontTableContentElement = getContent().getFontTableContentElement();
      
      Document doc = getDocument();
      
      RelationshipMap rsm = getRelationshipMap(doc.getPartName());
      
      Relationship r = new Relationship();
      r.setTarget("fontTable.xml");
      r.setId("rId" + (nextRelationshipId++));
      r.setType(Relationship.Type.FONT_TABLE);
      
      rsm.put(r);
    }
    return fontTableContentElement;
  }
  
  private Para defaultPara = null;
  
  public Para getTemplatePara() {
    if (defaultPara == null) {
      defaultPara = new Para(null, msHelper);
    }
    return defaultPara;
  }
  
  private Content content = null;
  
  private Content getContent() {
    if (content == null) content = Content.createDefaultContent(msHelper);
    return content;
  }
  
  private boolean headerConnected = false, footerConnected = false;
  
  private RelationshipMap getRelationshipMap(String subjectPartName) {
    RelationshipMap ret = relationShipMaps.get(subjectPartName);
    if (ret == null) {
      ret = RelationshipMap.createBySubjectPartName(subjectPartName);
      relationShipMaps.put(subjectPartName, ret);
    }
    return ret;
  }
  
  public DocumentHeader getOrCreateHeader() {
    checkInit();
    
    DocumentHeader header = getContent().getOrCreateHeader();
    if (!headerConnected) {
      Document doc = getDocument();
      
      RelationshipMap rsm = getRelationshipMap(doc.getPartName());
      
      Relationship r = new Relationship();
      r.setTarget("header1.xml");
      r.setId("rId" + (nextRelationshipId++));
      r.setType(Relationship.Type.HEADER);
      rsm.put(r);
      
      Reference ref = new Reference();
      ref.setId(r.getId());
      ref.setTagName("w:headerReference");
      doc.addReference(ref);
      
      headerConnected = true;
    }
    return header;
  }
  
  public DocumentFooter getOrCreateFooter() {
    checkInit();
    
    DocumentFooter footer = getContent().getOrCreateFooter();
    if (!footerConnected) {
      Document doc = getDocument();
      
      RelationshipMap rsm = getRelationshipMap(doc.getPartName());
      
      Relationship r = new Relationship();
      r.setTarget("footer1.xml");
      r.setId("rId" + (nextRelationshipId++));
      r.setType(Relationship.Type.FOOTER);
      rsm.put(r);
      
      Reference ref = new Reference();
      ref.setId(r.getId());
      ref.setTagName("w:footerReference");
      doc.addReference(ref);
      
      footerConnected = true;
    }
    return footer;
  }
  
  private void outFilePart(FilePart from, ZipOutputStream to) throws IOException {
    to.putNextEntry(new ZipEntry(UtilOffice.killFirstSlash(from.getPartName())));
    {
      PrintStream out = new PrintStream(to, false, "UTF-8");
      from.write(out);
      out.flush();
    }
    to.closeEntry();
  }
  
  private List<BinaryFilePart> binaryParts = new ArrayList<BinaryFilePart>();
  
  public void write(OutputStream out) throws Exception {
    final ZipOutputStream zout;
    if (out instanceof ZipOutputStream) {
      zout = (ZipOutputStream)out;
    } else {
      zout = new ZipOutputStream(out);
    }
    {
      List<FilePart> parts = new ArrayList<FilePart>();
      parts.add(getContent());
      parts.addAll(relationShipMaps.values());
      getContent().addAllFileParts(parts);
      
      for (FilePart fp : parts) {
        outFilePart(fp, zout);
      }
    }
    {
      for (BinaryFilePart bfp : binaryParts) {
        zout.putNextEntry(new ZipEntry(UtilOffice.killFirstSlash(bfp.getPartName())));
        InputStream in = bfp.openInputStream();
        UtilOffice.copyStreams(in, zout);
        in.close();
        zout.closeEntry();
      }
    }
    zout.flush();
    zout.close();
  }
  
  private void write0(File outFile) throws Exception {
    write(new FileOutputStream(outFile));
  }
  
  public void write(File outFile) {
    try {
      write0(outFile);
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }
  
  public void write(String outFileName) {
    write(new File(outFileName));
  }
  
  public static void main(String[] args) {
    Docx w = new Docx();
    Document doc = w.getDocument();
    Para p = doc.createPara();
    Run r = p.createRun();
    r.setFontName(DefaultFontNames.TIMES_NEW_ROMAN);
    r.addText("Привет вир!!!");
    
    w.write("build/u.docx");
    
    System.out.println("Complete.....");
  }
}
