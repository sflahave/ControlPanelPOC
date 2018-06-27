import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiPoc {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(AsposePoc.class);

        // Instantiating a Workbook object that represents a Microsoft Excel file.
        String filePath = dataDir + "junk.xlsx";

        Workbook wb = new XSSFWorkbook(filePath);

        Sheet sheet = wb.getSheetAt(0);

        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

//        CellReference inputCellRef = new CellReference("A1");
//        Row inputRow = sheet.getRow(inputCellRef.getRow());
//        Cell inputCell = inputRow.getCell(inputCellRef.getCol());
//        inputCell.setCellValue(12);

        Row r1 = sheet.getRow(0);
        Cell a1 = r1.getCell(0);
        a1.setCellValue(13);

        // suppose your formula is in A2
        CellReference cellReference = new CellReference("A2");
        Row row = sheet.getRow(cellReference.getRow());
        Cell cell = row.getCell(cellReference.getCol());

        CellValue cellValue = evaluator.evaluate(cell);

        Double value = cellValue.getNumberValue();

        System.out.println("Output: " + value);

    }

}
