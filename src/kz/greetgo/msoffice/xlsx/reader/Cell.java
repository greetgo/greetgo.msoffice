package kz.greetgo.msoffice.xlsx.reader;

import java.util.Date;

public interface Cell {
  String asText();

  Date asDate();

  String format();
}
