package kz.greetgo.msoffice.docx;

interface MSHelper {
  Relationship createRelationshipForImage(String ownerPartName, InputSource inputSource);
  
  Font getFont(String name);
  
  Para getDefaultPara();
}
