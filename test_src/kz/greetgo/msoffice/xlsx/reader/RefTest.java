package kz.greetgo.msoffice.xlsx.reader;

import kz.greetgo.msoffice.test.RND;
import org.testng.annotations.Test;

import static org.fest.assertions.api.Assertions.assertThat;

public class RefTest {

  @Test
  public void writeTo__readFrom() {

    byte[] buffer = new byte[80];

    Ref ref1 = Ref.of(RND.rnd.nextLong(), RND.rnd.nextInt());
    Ref ref2 = Ref.of(RND.rnd.nextLong(), RND.rnd.nextInt());
    Ref ref3 = Ref.of(RND.rnd.nextLong(), RND.rnd.nextInt());
    Ref ref4 = Ref.of(RND.rnd.nextLong(), RND.rnd.nextInt());

    ref1.writeTo(buffer, 0);
    ref2.writeTo(buffer, 20);
    ref3.writeTo(buffer, 40);
    ref4.writeTo(buffer, 60);

    Ref actual1 = Ref.readFrom(buffer, 0);
    Ref actual2 = Ref.readFrom(buffer, 20);
    Ref actual3 = Ref.readFrom(buffer, 40);
    Ref actual4 = Ref.readFrom(buffer, 60);

    assertThat(actual1.offset).isEqualTo(ref1.offset);
    assertThat(actual1.length).isEqualTo(ref1.length);

    assertThat(actual2.offset).isEqualTo(ref2.offset);
    assertThat(actual2.length).isEqualTo(ref2.length);

    assertThat(actual3.offset).isEqualTo(ref3.offset);
    assertThat(actual3.length).isEqualTo(ref3.length);

    assertThat(actual4.offset).isEqualTo(ref4.offset);
    assertThat(actual4.length).isEqualTo(ref4.length);
  }

  @Test
  public void writeTo() {
    byte[] buffer = new byte[16];

    Ref ref = Ref.of(123, 543);

    ref.writeTo(buffer, 0);

    //@formatter:off
    assertThat(buffer[ 0]).isEqualTo((byte) 123);
    assertThat(buffer[ 1]).isEqualTo((byte) 0);
    assertThat(buffer[ 2]).isEqualTo((byte) 0);
    assertThat(buffer[ 3]).isEqualTo((byte) 0);
    assertThat(buffer[ 4]).isEqualTo((byte) 0);
    assertThat(buffer[ 5]).isEqualTo((byte) 0);
    assertThat(buffer[ 6]).isEqualTo((byte) 0);
    assertThat(buffer[ 7]).isEqualTo((byte) 0);
    assertThat(buffer[ 8]).isEqualTo((byte) 31);
    assertThat(buffer[ 9]).isEqualTo((byte) 2);
    assertThat(buffer[10]).isEqualTo((byte) 0);
    assertThat(buffer[11]).isEqualTo((byte) 0);
    assertThat(buffer[12]).isEqualTo((byte) 0);
    assertThat(buffer[13]).isEqualTo((byte) 0);
    assertThat(buffer[14]).isEqualTo((byte) 0);
    assertThat(buffer[15]).isEqualTo((byte) 0);
    //@formatter:on

  }
}
