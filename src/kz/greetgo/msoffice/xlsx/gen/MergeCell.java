package kz.greetgo.msoffice.xlsx.gen;

import static kz.greetgo.msoffice.UtilOffice.toLettersNumber;

import java.io.PrintStream;

class MergeCell {
  int rowFrom, colFrom, rowTo, colTo;
  
  public void print(PrintStream out) {
    out.println("<mergeCell ref=\"" + toLettersNumber(colFrom) + rowFrom + ":"
        + toLettersNumber(colTo) + rowTo + "\" />");
  }
}
