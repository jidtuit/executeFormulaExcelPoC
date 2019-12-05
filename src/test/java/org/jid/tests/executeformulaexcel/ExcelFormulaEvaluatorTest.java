package org.jid.tests.executeformulaexcel;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.nio.file.Path;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.jupiter.api.Test;

class ExcelFormulaEvaluatorTest {

  private static final Path excel1Path = Path.of(".", "src", "test", "resources", "example1.xlsx");

  @Test
  void shouldReturnExpectedValueWithSimpleFunction() throws IOException, InvalidFormatException {

    // Formula: A2 + B2 = C2

    ExcelFormulaEvaluator evaluator = ExcelFormulaEvaluator.create(excel1Path);
    int sheet = 0;

    evaluator.setNumericValue(sheet, "A2", 2);
    evaluator.setNumericValue(sheet, "B2", 3);

    assertThat(evaluator.getNumericValue(sheet, "C2"))
        .isNotNull()
        .isNotEmpty()
        .contains(5d);

    evaluator.close();

  }

  @Test
  void shouldReturnExpectedValueWithFormulasConcatenated() throws IOException, InvalidFormatException {

    // Formulas: A5 + B5 = C5  ---- C5 * D5 = E5

    ExcelFormulaEvaluator evaluator = ExcelFormulaEvaluator.create(excel1Path);
    int sheet = 0;

    evaluator.setNumericValue(sheet, "A5", 2);
    evaluator.setNumericValue(sheet, "B5", 3);
    evaluator.setNumericValue(sheet, "D5", 4);

    assertThat(evaluator.getNumericValue(sheet, "E5"))
        .isNotNull()
        .isNotEmpty()
        .contains(20d);

    evaluator.close();

  }


  @Test
  void shouldReturnExpectedValueWithValuesInDifferentSheets() throws IOException, InvalidFormatException {

    // Formula: Sheet0-A9 + Sheet1-B2 = Sheet0-B9

    ExcelFormulaEvaluator evaluator = ExcelFormulaEvaluator.create(excel1Path);
    int sheet = 0;
    int paramsSheet = 1;

    evaluator.setNumericValue(sheet, "A9", 2);
    evaluator.setNumericValue(paramsSheet, "B2", 3);

    assertThat(evaluator.getNumericValue(sheet, "B9"))
        .isNotNull()
        .isNotEmpty()
        .contains(5d);

    evaluator.close();

  }

}