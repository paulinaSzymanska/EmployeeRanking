package excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.List;

public class ExcelReader {
    public static void readExcelFileIfExists(String folderPath) {
        List<File> excelFiles = ExcelFileFinder.createListOfExcelFiles(folderPath);
        if (excelFiles.isEmpty()) {
            System.out.println("❌ Excel files not found in the specified folder.");
            return;
        }

        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

        for (File excelFile : excelFiles) {
            System.out.println("\nFound excel file " + excelFile.getAbsolutePath());
            try (FileInputStream fis = new FileInputStream(excelFile);
                 Workbook workbook = createWorkbook(fis, excelFile.getName())) {

                Sheet sheet = workbook.getSheetAt(0);
                System.out.println("\n✅ Excel data:");

                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (cell == null) {
                            System.out.print("\t");
                            continue;
                        }
                        switch (cell.getCellType()) {
                            case STRING:
                                System.out.print(cell.getStringCellValue() + "\t");
                                break;
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    System.out.print(sdf.format(cell.getDateCellValue()) + "\t");
                                } else {
                                    System.out.print(cell.getNumericCellValue() + "\t");
                                }
                                break;
                            case BOOLEAN:
                                System.out.print(cell.getBooleanCellValue() + "\t");
                                break;
                            case FORMULA:
                                System.out.print(cell.getCellFormula() + "\t");
                                break;
                            case BLANK:
                                System.out.print("\t");
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
