package kz.greetgo.msoffice.xlsx.reader.model;

import org.testng.annotations.Test;

import java.util.Random;

import static kz.greetgo.msoffice.UtilTest.rndColData;
import static org.fest.assertions.api.Assertions.assertThat;

public class ColDataTest {

  @Test
  public void writeBytes__readBytes__fromZero() {

    Random rnd = new Random();

    ColData colData = rndColData(rnd);

    assertThat(colData.byteSize()).isGreaterThan(10);

    byte[] buffer = new byte[colData.byteSize()];
    rnd.nextBytes(buffer);

    //
    //
    colData.writeBytes(buffer, 0);
    ColData actual = ColData.readBytes(buffer, 0);
    //
    //

    assertThat(actual).isNotNull();
    assertThat(actual.col).isEqualTo(colData.col);
    assertThat(actual.valueType).isEqualTo(colData.valueType);
    assertThat(actual.value).isEqualTo(colData.value);

  }

  @Test
  public void writeBytes__readBytes() {

    Random rnd = new Random();

    ColData colData = rndColData(rnd);

    assertThat(colData.byteSize()).isGreaterThan(10);

    byte[] buffer = new byte[colData.byteSize() + 20];

    //
    //
    colData.writeBytes(buffer, 10);
    ColData actual = ColData.readBytes(buffer, 10);
    //
    //

    assertThat(actual).isNotNull();
    assertThat(actual.col).isEqualTo(colData.col);
    assertThat(actual.valueType).isEqualTo(colData.valueType);
    assertThat(actual.value).isEqualTo(colData.value);

  }
}
