package org.example;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

public class GoogleSearch {

    public static void main(String[] args) throws IOException {
        if (args.length < 1) {
            System.out.println("Usage: java GoogleSearch <path_to_excel_file>");
            System.exit(1);
        }

        String excelFilePath = args[0];
        FileInputStream file = new FileInputStream(new File(excelFilePath));
        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheetAt(0);

        // Prepare to write results back to Excel
        Workbook resultWorkbook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkbook.createSheet("Results");
        int resultRowNum = 0;

        for (Row row : sheet) {
            Cell cell = row.getCell(0); // Assuming keywords are in the first column
            if (cell != null) {
                String keyword = cell.getStringCellValue().trim();
                if (!keyword.isEmpty()) {
                    String firstUrl = searchGoogle(keyword);

                    // Write the keyword and first URL to the result sheet
                    Row resultRow = resultSheet.createRow(resultRowNum++);
                    resultRow.createCell(0).setCellValue(keyword);
                    resultRow.createCell(1).setCellValue(firstUrl);
                }
            }
        }

        // Save results to a new Excel file
        FileOutputStream outputStream = new FileOutputStream("search_results.xlsx");
        resultWorkbook.write(outputStream);
        resultWorkbook.close();

        workbook.close();
        file.close();
    }

    public static String searchGoogle(String keyword) {
        WebDriverManager.firefoxdriver().setup();
        WebDriver driver = new FirefoxDriver();
        driver.get("https://www.google.com");

        WebElement searchBox = driver.findElement(By.name("q"));
        searchBox.sendKeys(keyword);
        searchBox.sendKeys(Keys.RETURN);

        // Wait for search results to load
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("h3")));

        // Get the first search result URL
        List<WebElement> searchResults = driver.findElements(By.cssSelector("h3"));
        String url = "";
        if (!searchResults.isEmpty()) {
            WebElement firstResult = searchResults.get(0);
            url = firstResult.findElement(By.xpath("./..")).getAttribute("href");
        }

        driver.quit();
        return url;
    }
}
