package kz.greetgo.msoffice.xlsx.xlsx_reader.model;

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
  public String value;

  public int byteSize() {
    return SIZEOF_SHORT + SIZEOF_INT + BinUtil.sizeOfStr(value);
  }

  public void writeBytes(byte[] buffer, int offset) {
    short valueTypeInt = (short) (valueType == null ? ValueType.UNKNOWN.ordinal() : valueType.ordinal());

    writeShort(valueTypeInt, buffer, offset);
    writeInt(col, buffer, offset + SIZEOF_SHORT);
    writeStr(value, buffer, offset + SIZEOF_SHORT + SIZEOF_INT);

  }

  public static ColData readBytes(byte[] buffer, int offset) {
    ColData colData = new ColData();
    colData.valueType = ValueType.values()[(int) readShort(buffer, offset)];
    colData.col = readInt(buffer, offset + SIZEOF_SHORT);
    colData.value = readStr(buffer, offset + SIZEOF_SHORT + SIZEOF_INT);
    return colData;
  }

}
