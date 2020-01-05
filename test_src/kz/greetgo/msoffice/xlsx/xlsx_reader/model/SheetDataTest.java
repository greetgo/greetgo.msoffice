package kz.greetgo.msoffice.xlsx.xlsx_reader.model;

import kz.greetgo.msoffice.UtilTest;
import org.testng.annotations.Test;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.function.Function;

import static org.fest.assertions.api.Assertions.assertThat;

public class SheetDataTest extends AbstractModelTest {

  @SuppressWarnings("SameParameterValue")
  private Function<String, Path> tmpDir(String pathInBuild) {
    String x = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss-SSS").format(new Date());
    Path dir = Paths.get("build").resolve("test-data").resolve(pathInBuild + "-" + x);
    dir.toFile().mkdirs();
    return prefix -> dir.resolve(prefix + "-" + Math.abs(rnd.nextLong()) + ".bin");
  }

  @Test
  public void checkRowsLogic() {

    List<RowData> rows = new ArrayList<>();

    rows.add(RowData.empty(0));
    rows.add(RowData.empty(1));
    rows.add(RowData.empty(2));

    SheetData sheet = new SheetData("wow", tmpDir("SheetDataTest"));

    for (int i = 0; i < 10; i++) {
      RowData row = UtilTest.rndRowData(rnd, rows.size());
      rows.add(row);
      sheet.addRow(row);
    }

    for (int i = 0; i < rows.size(); i++) {
      RowData expected = rows.get(i);
      RowData actual = sheet.getRowData(i);
      assertEquals(actual, expected, "row index = " + i);
    }

    assertThat(sheet.rowSize()).isEqualTo(rows.size());

  }

}
