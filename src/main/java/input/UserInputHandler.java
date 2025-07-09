package input;

import java.io.File;
import java.util.Scanner;

public class UserInputHandler {
    public static String askForDirectoryPath() {
        Scanner scanner = new Scanner(System.in);
        System.out.print("Enter the path to the folder containing the Excel file: ");
        String folderPath = scanner.nextLine();

        File folder = new File(folderPath);

        while (!folder.exists() || !folder.isDirectory()) {
            System.out.println("❌ Invalid path. Please enter a valid directory:");
            folderPath = scanner.nextLine();
            folder = new File(folderPath);
        }

        return folderPath;
    }
}
