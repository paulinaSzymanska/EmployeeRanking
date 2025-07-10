import excel.ExcelReader;
import excel.Entry;
import excel.StatisticsGenerator;

import java.util.List;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        System.out.print("Enter the path to the folder containing Excel files: ");
        String folderPath = scanner.nextLine();

        List<Entry> entries = ExcelReader.readEntriesFromExcel(folderPath);

        if (entries.isEmpty()) {
            System.out.println("❌ No data entries found.");
            return;
        }

        StatisticsGenerator.generateStatistics(entries);
    }
}
