import input.UserInputHandler;
import excel.ExcelReader;

public class Main {
    public static void main(String[] args) {
        String folderPath = UserInputHandler.askForDirectoryPath();
        ExcelReader.readExcelFileIfExists(folderPath);
    }
}
