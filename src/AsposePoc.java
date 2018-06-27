import java.io.FileInputStream;

import com.aspose.cells.Cell;
import com.aspose.cells.FileFormatType;
import com.aspose.cells.License;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AsposePoc {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(AsposePoc.class);

        // Creating a file input stream to reference the license file
//        FileInputStream fstream = new FileInputStream(dataDir + "Aspose.Cells.lic");
//
//        // Create a License object
//        License license = new License();
//
//        // Applying the Aspose.Cells license
//        license.setLicense(fstream);

        // Instantiating a Workbook object that represents a Microsoft Excel file.
        String filePath = dataDir + "junk.xlsx";
        Workbook wb = new Workbook(filePath);

        // Note when you create a new workbook, a default worksheet, "Sheet1", is by default added to the workbook. Accessing the
        // first worksheet in the book ("Sheet1").
        Worksheet sheet = wb.getWorksheets().get(0);

        // Access cell "A1" in the sheet.
        Cell inputCell = sheet.getCells().get("A1");

        // Input the "Hello World!" text into the "A1" cell
        inputCell.setValue(3);

        Cell outputCell = sheet.getCells().get("A2");
        String formula = outputCell.getFormula();
        Double result = (Double) sheet.calculateFormula(formula);

        System.out.print("Output: " + result);

        // Save the Microsoft Excel file.
//        wb.save(filePath);
//        wb.save(dataDir + "CreatingWorkbook_out.xls", FileFormatType.EXCEL_97_TO_2003);
//        wb.save(dataDir + "CreatingWorkbook_out.xlsx");
//        wb.save(dataDir + "CreatingWorkbook_out.ods");

    }
}