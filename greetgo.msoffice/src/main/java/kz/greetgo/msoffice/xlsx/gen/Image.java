package kz.greetgo.msoffice.xlsx.gen;

import kz.greetgo.msoffice.ImageType;

/**
 * Изображение. Нужен для повторной вставки изображения.
 */
public class Image {
  
  private final String filename; // имя файла с изображением
  private final ImageType type; // формат изображения
  
  Image(String filename, ImageType type) {
    this.filename = filename;
    this.type = type;
  }
  
  String getFilename() {
    return filename;
  }
  
  ImageType getImageType() {
    return type;
  }
  
  String getType() {
    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/"
        + filename;
  }
  
  @Override
  public int hashCode() {
    return filename.hashCode();
  }
  
  @Override
  public boolean equals(Object o) {
    
    if (this == o) return true;
    if (o == null) return false;
    if (!(o instanceof Image)) return false;
    
    return filename.equals(((Image)o).filename);
  }
}