package kz.greetgo.msoffice.xlsx.parse2;

import kz.greetgo.msoffice.test.RND;
import org.testng.annotations.Test;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import static java.lang.Math.abs;
import static org.fest.assertions.api.Assertions.assertThat;

public class StoredStringsTest {

  @Test
  public void readWrite() {
    long startedAt = System.currentTimeMillis();

    Path dir = Paths.get("build/StoredStringsTest/test1");
    dir.toFile().mkdirs();

    List<String> list = new ArrayList<>();

    {
      long t1 = System.currentTimeMillis();
      for (int i = 0, c = 3_000_000 + RND.plusInt(5000); i < c; i++) {
        list.add(RND.str(7 + RND.plusInt(7)));
      }
      long t2 = System.currentTimeMillis();
      System.out.println("3h25v325 :: create time     = " + (t2 - t1));
    }

    Path refFile = dir.resolve("ref" + abs(RND.rnd.nextLong()));
    Path contentFile = dir.resolve("content" + abs(RND.rnd.nextLong()));

    try (StoredStrings ss = new StoredStrings(refFile, contentFile, 10, 15)) {
      {
        long t1 = System.currentTimeMillis();
        long i = 0;
        for (String str : list) {
          long index = ss.append(str);
          assertThat(index).isEqualTo(i++);
        }
        long t2 = System.currentTimeMillis();
        System.out.println("3h25v325 :: append time     = " + (t2 - t1));
      }

      {
        long t1 = System.currentTimeMillis();
        for (int i = 0; i < list.size(); i++) {
          String actual = ss.get(i);
          String expected = list.get(i);
          assertThat(actual).isEqualTo(expected);
        }
        long t2 = System.currentTimeMillis();
        System.out.println("3h25v325 :: order read time = " + (t2 - t1));
      }

      {
        long t1 = System.currentTimeMillis();
        for (int i = 0; i < list.size(); i++) {
          int j = RND.plusInt(list.size());
          String actual = ss.get(j);
          String expected = list.get(j);
          assertThat(actual).isEqualTo(expected);
        }
        long t2 = System.currentTimeMillis();
        System.out.println("3h25v325 :: rnd read time   = " + (t2 - t1));
      }

      System.out.println("3h25v325 :: total test time = " + (System.currentTimeMillis() - startedAt));

      assertThat(ss.strCount()).isEqualTo(list.size());


    }

  }

}
