package com.allcloud;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CustomerCIFID {

    WebDriver driver;
    String cellValue;
    @Test(priority=1)
    public void setup() {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.manage().deleteAllCookies();
    }


    public void ss() throws IOException, InterruptedException {
    	

        String filePath = System.getProperty("user.dir") + "\\src\\test\\CIFData.xlsx";
        FileInputStream fileInputStream = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("Sheet1"); // Replace with your sheet name
        // Specify the starting row and column index
        int startRowIndex = 0; // Replace with the starting row index
        int columnIndex = 0; // Replace with the column index (0 for the first column)

        // Iterate through the cells to find the first non-null value
        String cellValue = null;
        for (int rowIndex = startRowIndex; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    if (cell.getCellType() == CellType.STRING) {
                        cellValue = cell.getStringCellValue();
                    }
                }
            }
            // If a non-null value is found, break the loop
            if (cellValue != null && !cellValue.isEmpty()) {
                break;
            }
        }
        System.out.println(cellValue);

        // Perform actions with the found cellValue
        if (cellValue != null) {
        	String url = "https://apps.allcloud.in/magfinserv/Customer/CustomerDetails/" + cellValue ;
            driver.get(url);
         

        }

            System.out.println(cellValue);
            this.cellValue = cellValue;
            String targetValue = cellValue; // Replace with the value you are searching for

            // Iterate through the cells to search for the target value
            Cell foundCell = null;
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        String cellValue1 = cell.getStringCellValue();
                        if (cellValue1.equals(targetValue)) {
                            foundCell = cell;
                            break;
                        }
                    }
                }
                if (foundCell != null) {
                    break;
                }
            }

            int rowIndex2 = foundCell.getRowIndex();
            int columnIndex2 = foundCell.getColumnIndex();

            System.out.println(rowIndex2);
            System.out.println(columnIndex2);

            Row row = sheet.getRow(foundCell.getRowIndex());
            if (row != null) {
                Cell cell1 = row.getCell(0);
                if (cell1 != null) {
                    // Clear the cell value
                    cell1.setCellValue("");

                    // Close the input stream
                    fileInputStream.close();

                    // Save the modified Excel file
                    FileOutputStream outputStream = new FileOutputStream(filePath);
                    workbook.write(outputStream);
                    outputStream.close();
                   
                }
            }
        }
    
    @Test(priority = 2)
   public void CIF() throws IOException, InterruptedException {
   	CustomerDetails details = new CustomerDetails();
       // Do something with the cellValue
   	   ss();
  	   details.OpenBrowser();
       details.ExtractCustomerDetails();
       driver.close();
   }
}
