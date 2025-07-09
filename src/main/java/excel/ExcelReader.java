package excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;     // do .xls
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;     // do .xlsx

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {
    public static void readExcelFile(String folderPath, String fileName) {
        File excelFile = new File(folderPath, fileName);

        if (!excelFile.exists()) {
            System.out.println("❌ File '" + fileName + "' not found in the specified folder.");
            return;
        }

        try (FileInputStream fis = new FileInputStream(excelFile);
             Workbook workbook = createWorkbook(fis, fileName)) {

            Sheet sheet = workbook.getSheetAt(0);
            System.out.println("\n✅ Excel data:");

            for (Row row : sheet) {
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        default:
                            System.out.print("?\t");
                    }
                }
                System.out.println();
            }

        } catch (IOException e) {
            System.out.println("❌ Error reading the Excel file: " + e.getMessage());
        }
    }

    private static Workbook createWorkbook(FileInputStream fis, String fileName) throws IOException {
        if (fileName.toLowerCase().endsWith(".xlsx")) {
            return new XSSFWorkbook(fis);
        } else if (fileName.toLowerCase().endsWith(".xls")) {
            return new HSSFWorkbook(fis);
        } else {
            throw new IllegalArgumentException("Unsupported file type: " + fileName);
        }
    }
}
