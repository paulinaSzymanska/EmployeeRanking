package excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelReader {
    public static List<Entry> readEntriesFromExcel(String folderPath) {
        List<File> excelFiles = ExcelFileFinder.createListOfExcelFiles(folderPath);
        List<Entry> entries = new ArrayList<>();

        for (File excelFile : excelFiles) {
            try (FileInputStream fis = new FileInputStream(excelFile);
                 Workbook workbook = createWorkbook(fis, excelFile.getName())) {

                Sheet sheet = workbook.getSheetAt(0);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) continue; // skip header
                    try {
                        Cell dateCell = row.getCell(0);
                        Cell nameCell = row.getCell(1);
                        Cell hoursCell = row.getCell(2);

                        if (dateCell == null || nameCell == null || hoursCell == null) continue;

                        Date date = DateUtil.isCellDateFormatted(dateCell)
                                ? dateCell.getDateCellValue()
                                : null;

                        String name = nameCell.getStringCellValue();
                        double hours = hoursCell.getNumericCellValue();

                        if (date != null) {
                            entries.add(new Entry(date, name, hours));
                        }
                    } catch (Exception ex) {
                        //do nothing
                    }
                }

            } catch (IOException e) {
                System.out.println("❌ Error reading Excel file: " + e.getMessage());
            }
        }
        return entries;
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
