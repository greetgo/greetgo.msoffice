package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.xlsx.reader.model.HorizontalAlign;
import kz.greetgo.msoffice.xlsx.reader.model.VerticalAlign;

import java.util.Date;

public interface Cell {
  String asText();

  Date asDate();

  boolean isDate();

  String format();

  Borders borders();

  HorizontalAlign horAlign();

  VerticalAlign vertAlign();
}
