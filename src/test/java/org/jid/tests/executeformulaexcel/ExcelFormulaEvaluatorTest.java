package org.jid.tests.executeformulaexcel;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.nio.file.Path;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.jupiter.api.Test;

class ExcelFormulaEvaluatorTest {

  private static final Path excel1Path = Path.of(".", "src", "test", "resources", "example1.xlsx");

  @Test
  void shouldReturnCorrectValueWithASimpleFunction() throws IOException, InvalidFormatException {

    // Formula: A2 + B2 = C2

    ExcelFormulaEvaluator evaluator = ExcelFormulaEvaluator.load(excel1Path);
    int sheet = 0;

    evaluator.setNumericValue(sheet, "A2", 2);
    evaluator.setNumericValue(sheet, "B2", 3);

    assertThat(evaluator.getNumericValue(sheet, "C2"))
        .isNotNull()
        .isNotEmpty()
        .contains(5d);

  }

  @Test
  void shouldReturnCorrectValueWithAConcatenationOfFormulas() throws IOException, InvalidFormatException {

    // Formulas: A5 + B5 = C5  ---- C5 * D5 = E5

    ExcelFormulaEvaluator evaluator = ExcelFormulaEvaluator.load(excel1Path);
    int sheet = 0;

    evaluator.setNumericValue(sheet, "A5", 2);
    evaluator.setNumericValue(sheet, "B5", 3);
    evaluator.setNumericValue(sheet, "D5", 4);

    assertThat(evaluator.getNumericValue(sheet, "E5"))
        .isNotNull()
        .isNotEmpty()
        .contains(20d);
    
  }


}