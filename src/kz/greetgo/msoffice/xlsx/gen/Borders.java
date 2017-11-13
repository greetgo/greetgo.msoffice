package kz.greetgo.msoffice.xlsx.gen;

import java.io.PrintStream;

public class Borders {
  private final Border left, right, top, bottom, diagonal;
  
  private boolean diagonalUp, diagonalDown;
  
  public Borders() {
    left = new Border();
    right = new Border();
    top = new Border();
    bottom = new Border();
    diagonal = new Border();
    diagonalUp = false;
    diagonalDown = false;
  }
  
  public void clean() {
    diagonalUp = false;
    diagonalDown = false;
    left.clean();
    right.clean();
    top.clean();
    bottom.clean();
    diagonal.clean();
  }
  
  public Borders(Borders x) {
    left = Border.copy(x.left);
    right = Border.copy(x.right);
    top = Border.copy(x.top);
    bottom = Border.copy(x.bottom);
    diagonal = Border.copy(x.diagonal);
    diagonalUp = x.diagonalUp;
    diagonalDown = x.diagonalDown;
  }
  
  public static Borders copy(Borders borders) {
    if (borders == null) return null;
    return new Borders(borders);
  }
  
  @Override
  public int hashCode() {
    final int prime = 31;
    int result = 1;
    result = prime * result + ((bottom == null) ? 0 :bottom.hashCode());
    result = prime * result + ((diagonal == null) ? 0 :diagonal.hashCode());
    result = prime * result + (diagonalDown ? 1231 :1237);
    result = prime * result + (diagonalUp ? 1231 :1237);
    result = prime * result + ((left == null) ? 0 :left.hashCode());
    result = prime * result + ((right == null) ? 0 :right.hashCode());
    result = prime * result + ((top == null) ? 0 :top.hashCode());
    return result;
  }
  
  @Override
  public boolean equals(Object obj) {
    if (this == obj) return true;
    if (obj == null) return false;
    if (getClass() != obj.getClass()) return false;
    Borders other = (Borders)obj;
    if (bottom == null) {
      if (other.bottom != null) return false;
    } else if (!bottom.equals(other.bottom)) return false;
    if (diagonal == null) {
      if (other.diagonal != null) return false;
    } else if (!diagonal.equals(other.diagonal)) return false;
    if (diagonalDown != other.diagonalDown) return false;
    if (diagonalUp != other.diagonalUp) return false;
    if (left == null) {
      if (other.left != null) return false;
    } else if (!left.equals(other.left)) return false;
    if (right == null) {
      if (other.right != null) return false;
    } else if (!right.equals(other.right)) return false;
    if (top == null) {
      if (other.top != null) return false;
    } else if (!top.equals(other.top)) return false;
    return true;
  }
  
  public Border left() {
    return left;
  }
  
  public Border right() {
    return right;
  }
  
  public Border top() {
    return top;
  }
  
  public Border bottom() {
    return bottom;
  }
  
  public Border diagonal() {
    return diagonal;
  }
  
  public boolean isDiagonalUp() {
    return diagonalUp;
  }
  
  public Borders diagonalUp() {
    diagonalUp = true;
    return this;
  }
  
  public Borders noDiagonalUp() {
    diagonalUp = false;
    return this;
  }
  
  public void setDiagonalUp(boolean diagonalUp) {
    this.diagonalUp = diagonalUp;
  }
  
  public boolean isDiagonalDown() {
    return diagonalDown;
  }
  
  public Borders diagonalDown() {
    diagonalDown = true;
    return this;
  }
  
  public Borders noDiagonalDown() {
    diagonalDown = false;
    return this;
  }
  
  public void setDiagonalDown(boolean diagonalDown) {
    this.diagonalDown = diagonalDown;
  }
  
  public void medium() {
    top.medium();
    right.medium();
    bottom.medium();
    left.medium();
  }
  
  public void thin() {
    top.thin();
    right.thin();
    bottom.thin();
    left.thin();
  }
  
  public void thick() {
    top.thick();
    right.thick();
    bottom.thick();
    left.thick();
  }
  
  public void print(PrintStream out) {
    out.print("<border");
    if (diagonalUp && diagonal != null) {
      out.print(" diagonalUp=\"1\"");
    }
    if (diagonalDown && diagonal != null) {
      out.print(" diagonalDown=\"1\"");
    }
    out.println(">");
    
    if (left == null) {
      out.print("<left />");
    } else {
      left.print(out, "left");
    }
    
    if (right == null) {
      out.print("<right />");
    } else {
      right.print(out, "right");
    }
    
    if (top == null) {
      out.print("<top />");
    } else {
      top.print(out, "top");
    }
    
    if (bottom == null) {
      out.print("<bottom />");
    } else {
      bottom.print(out, "bottom");
    }
    
    if (diagonal == null) {
      out.print("<diagonal />");
    } else {
      diagonal.print(out, "diagonal");
    }
    
    out.print("</border>");
  }
  
}
