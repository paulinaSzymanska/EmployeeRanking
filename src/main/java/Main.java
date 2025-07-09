import input.UserInputHandler;
import excel.ExcelReader;

import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        String folderPath = UserInputHandler.askForDirectoryPath();

        System.out.print("Enter the Excel file name (e.g. data.xls or data.xlsx): ");
        String fileName = scanner.nextLine();

        ExcelReader.readExcelFile(folderPath, fileName);
    }
}
