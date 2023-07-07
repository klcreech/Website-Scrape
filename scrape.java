import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.NoSuchElementException;
import java.io.File;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.Date;

public class Main {
    public static void main(String[] args) {
        WebDriver driver = new ChromeDriver();
        int start_index = 12;
        int skip_value = 12;
        try {
            Workbook workbook = WorkbookFactory.create(new File("E:\\xx.xlsx"));
            Sheet sheet = workbook.getSheetAt(0);
        } catch (IOException e) {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("xx");
        }
        int empty_row = sheet.getLastRowNum() + 1;
        if (empty_row == 2) {
            String[] column_names = {
                "xx"
            };
            for (int column_index = 0; column_index < column_names.length; column_index++) {
                String column_name = column_names[column_index];
                Row row = sheet.getRow(0);
                if (row == null) {
                    row = sheet.createRow(0);
                }
                Cell cell = row.createCell(column_index);
                cell.setCellValue(column_name);
            }
        }
