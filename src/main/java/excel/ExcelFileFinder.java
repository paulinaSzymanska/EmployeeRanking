package excel;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class ExcelFileFinder {

    public static List<File> createListOfExcelFiles(String folderPath) {
        List<File> excelFiles = new ArrayList<>();
        File folder = new File(folderPath);

        if (!folder.exists() || !folder.isDirectory()) {
            System.out.println("Folder does not exist or is not a directory: " + folderPath);
            return excelFiles;
        }

        searchExcelFilesRecursively(folder, excelFiles);
        return excelFiles;
    }

    private static void searchExcelFilesRecursively(File folder, List<File> excelFiles) {
        File[] filesAndDirs = folder.listFiles();
        if (filesAndDirs == null) return;

        for (File f : filesAndDirs) {
            if (f.isDirectory()) {
                searchExcelFilesRecursively(f, excelFiles);
            } else if (f.isFile()) {
                String name = f.getName().toLowerCase();
                if (name.endsWith(".xls") || name.endsWith(".xlsx")) {
                    excelFiles.add(f);
                }
            }
        }
    }
}
