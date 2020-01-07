package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

import static kz.greetgo.msoffice.util.UtilOffice.toLettersNumber;

class MergeCell {
  int rowFrom, colFrom, rowTo, colTo;

  public void print(PrintStream out) {
    out.println("<mergeCell ref=\"" + toLettersNumber(colFrom) + rowFrom + ":"
      + toLettersNumber(colTo) + rowTo + "\" />");
  }
}
