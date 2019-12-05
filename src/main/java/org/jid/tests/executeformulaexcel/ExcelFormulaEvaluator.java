package org.jid.tests.executeformulaexcel;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Optional;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFormulaEvaluator implements AutoCloseable{

  private XSSFWorkbook wb;
  private FormulaEvaluator evaluator;

  public static ExcelFormulaEvaluator create(Path path) throws IOException, InvalidFormatException {

    if(path == null || !Files.exists(path)) {
      throw new IOException("File path must exist");
    }

    XSSFWorkbook wb = new XSSFWorkbook(path.toFile());

    return new ExcelFormulaEvaluator(wb);
  }


  public Optional<Double> getNumericValue(int sheetPosition, String cellReference) {

    XSSFSheet sheet = wb.getSheetAt(sheetPosition);
    CellReference cellRef = new CellReference(cellReference);
    Row row = sheet.getRow(cellRef.getRow());
    Cell cell = row.getCell(cellRef.getCol());

    Optional<Double> resp= Optional.empty();
    if (cell!=null) {
      // It can be solved with an if-else but I wanted to be explicit about the different types for this example
      switch (cell.getCellType()) {
        case NUMERIC:
          resp = Optional.of(cell.getNumericCellValue());
          break;
        case FORMULA:
          resp = Optional.ofNullable(evaluator.evaluate(cell).getNumberValue());
          break;
        case BOOLEAN:
        case STRING:
        case BLANK:
        case ERROR:
        case _NONE:
          break; // Optional of empty
      }
    }

    return resp;
  }


  public void setNumericValue(int sheetPosition, String cellReference, double value) {

    XSSFSheet sheet = wb.getSheetAt(sheetPosition);
    CellReference cellRef = new CellReference(cellReference);
    Row row = sheet.getRow(cellRef.getRow());
    Cell cell = row.getCell(cellRef.getCol());

    cell.setCellValue(value);
  }


  private ExcelFormulaEvaluator(XSSFWorkbook wb)  {
    this.wb = wb;
    this.evaluator = wb.getCreationHelper().createFormulaEvaluator();
  }


  @Override
  public void close() throws IOException {
    if(wb != null)
      wb.close();
  }
}
