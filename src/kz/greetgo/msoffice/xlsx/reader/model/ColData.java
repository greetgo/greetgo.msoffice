package kz.greetgo.msoffice.xlsx.reader.model;

import kz.greetgo.msoffice.util.BinUtil;

import static kz.greetgo.msoffice.util.BinUtil.SIZEOF_INT;
import static kz.greetgo.msoffice.util.BinUtil.SIZEOF_SHORT;
import static kz.greetgo.msoffice.util.BinUtil.readInt;
import static kz.greetgo.msoffice.util.BinUtil.readShort;
import static kz.greetgo.msoffice.util.BinUtil.readStr;
import static kz.greetgo.msoffice.util.BinUtil.writeInt;
import static kz.greetgo.msoffice.util.BinUtil.writeShort;
import static kz.greetgo.msoffice.util.BinUtil.writeStr;

public class ColData {
  public ValueType valueType;
  public int col;
  public int style;
  public String value;

  public static ColData empty(int colIndex) {
    ColData ret = new ColData();
    ret.col = colIndex;
    ret.style = 0;
    ret.valueType = ValueType.UNKNOWN;
    ret.value = null;
    return ret;
  }

  public int byteSize() {
    return SIZEOF_SHORT + SIZEOF_INT * 2 + BinUtil.sizeOfStr(value);
  }

  public void writeBytes(byte[] buffer, int offset) {
    short valueTypeInt = (short) (valueType == null ? ValueType.UNKNOWN.ordinal() : valueType.ordinal());

    writeShort(valueTypeInt, buffer, offset);
    writeInt(col, buffer, offset + SIZEOF_SHORT);
    writeInt(style, buffer, offset + SIZEOF_SHORT + SIZEOF_INT);
    writeStr(value, buffer, offset + SIZEOF_SHORT + SIZEOF_INT * 2);

  }

  public static ColData readBytes(byte[] buffer, int offset) {
    ColData colData = new ColData();
    colData.valueType = ValueType.values()[(int) readShort(buffer, offset)];
    colData.col = readInt(buffer, offset + SIZEOF_SHORT);
    colData.style = readInt(buffer, offset + SIZEOF_SHORT + SIZEOF_INT);
    colData.value = readStr(buffer, offset + SIZEOF_SHORT + SIZEOF_INT * 2);
    return colData;
  }

}
