package kz.greetgo.msoffice;

/**
 * Форматы файлов с изображениями.
 */
public enum ImageType {
  
  PNG("png"), //
  JPG("jpg"), //
  BMP("bmp"), //
  GIF("gif"); //
  
  private final String ext; // расширение имени файла
  
  private ImageType(String ext) {
    this.ext = ext;
  }
  
  public String getExt() {
    return ext;
  }
  
  public static ImageType getByExt(String ext) {
    
    for (ImageType type : values())
      if (type.ext.equals(ext)) return type;
    
    if ("jpeg".equals(ext)) return JPG;
    
    return null;
  }
}