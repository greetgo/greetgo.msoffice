package kz.greetgo.msoffice;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import static org.fest.assertions.api.Assertions.assertThat;

public class UtilOfficeTest {

  @DataProvider
  Object[][] dataProvider__parseLettersNumber() {
    return new Object[][]{
      {"L", 11},
      {"AA", 26},
    };
  }

  @Test(dataProvider = "dataProvider__parseLettersNumber")
  public void parseLettersNumber(String letterNumber, int number) {
    assertThat(UtilOffice.parseLettersNumber(letterNumber)).isEqualTo(number);
  }

  @DataProvider
  Object[][] dataProvider__cellCoordinates() {
    return new Object[][]{
      {"A1", 1, 1},
      {"A2", 1, 2},
      {"B1", 2, 1},
      {"B2", 2, 2},
      {"L18", 12, 18},
      {"AA200", 27, 200},
    };
  }

  @Test(dataProvider = "dataProvider__cellCoordinates")
  public void cellCoordinates(String cellCoordinate, int col, int row) {
    assertThat(UtilOffice.parseCellCoordinate(cellCoordinate)[0]).isEqualTo(col);
    assertThat(UtilOffice.parseCellCoordinate(cellCoordinate)[1]).isEqualTo(row);
  }

}
