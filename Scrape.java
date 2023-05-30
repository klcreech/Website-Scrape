// Java Web scrape - Kerry

// To use this code:

//  1.Replace the placeholder values with your desired URLs, CSS selectors, file paths, and sheet names.
//  2.Set the start_index and skip_value variables if you need to iterate through multiple pages.
//  3.Execute the code to launch the Chrome browser and start scraping the data.
//  4.The scraped data will be stored in the Excel file specified.
//  5.After the scraping is complete, the browser will be closed, and the Excel file will be saved.




import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.NoSuchElementException;
import java.io.File;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WebScraper {
    public static void main(String[] args) {
        // Launch the Chrome browser
        WebDriver driver = new ChromeDriver();

        // Start page index and skip value
        int start_index = 12; // Use this if you have a skip value in the URL
        int skip_value = 12; // Use this if you have a skip value in the URL

        try {
            // Load existing Excel workbook or create a new one if it doesn't exist
            Workbook workbook;
            Sheet sheet;
            File file = new File("E:\\xx.xlsx"); // Replace with the desired file path
            if (file.exists()) {
                workbook = WorkbookFactory.create(file);
                sheet = workbook.getSheetAt(0);
            } else {
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("xx"); // Replace with the desired sheet name
            }

            // Find the first empty row in the Excel sheet
            int empty_row = sheet.getLastRowNum() + 1;

            // Write the column headers if the sheet is empty
            if (empty_row == 1) {
                String[] column_names = {
                    "xx" // Replace with column header names, add as many as needed
                };
                Row headerRow = sheet.createRow(0);
                for (int column_index = 0; column_index < column_names.length; column_index++) {
                    String column_name = column_names[column_index];
                    Cell cell = headerRow.createCell(column_index);
                    cell.setCellValue(column_name);
                }
            }

            // Scrape data and populate the Excel sheet
            while (true) {
                // Update the URL with the skip value
                String url = "https://www.thesite.org"; // Replace with the desired URL

                // Navigate to the website
                driver.get(url);

                // Find all xx.xx elements
                List<WebElement> item_elements = driver.findElements(By.cssSelector("xx.xx")); // Replace with the desired CSS selector

                // If no item elements are found, exit the loop
                if (item_elements.isEmpty()) {
                    break;
                }

                // Iterate over each item element
                for (WebElement item_element : item_elements) {
                    // Find the x[class='x'] element within the current item element
                    try {
                        WebElement first_element = item_element.findElement(By.cssSelector("x[class='x']")); // Replace with the desired CSS selector
                        String first_data = first_element.getText();

                        // Add the above as needed
                        Row dataRow = sheet.createRow(empty_row);
                        Cell cell = dataRow.createCell(0);
                        cell.setCellValue(first_data);
                        
                        empty_row++;
                    } catch (NoSuchElementException e) {
                        // Handle the case when the element is not found
                    }
                }

                // Increment the start index for the next page ** for multiple page scrapes
                start_index += skip_value;
            }

            // Save the Excel file
            workbook.write(new FileOutputStream("E:\\xx.xlsx")); // Replace with the desired file path

            // Close the workbook and browser
            workbook.close();
            driver.quit();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
