package com.allcloud;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.temporal.TemporalAmount;

import org.apache.commons.io.FileUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import com.google.common.collect.Table.Cell;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CustomerDetails {
	public WebDriver driver;
	
	
	
	public void Setup() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();

	}
	

	public void OpenBrowser() throws IOException {
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));

		
		WebElement Username = driver.findElement(By.id("UserName"));
		wait.until(ExpectedConditions.visibilityOf(Username));
		Username.sendKeys("data.migration");
		
		WebElement Password = driver.findElement(By.id("Password"));
		Password.sendKeys("RoopA@1990");
		
		driver.findElement(By.id("btnSignIn")).click();
		
		
		  try {
				WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(5));

		        WebElement toastMessage = wait1.until(ExpectedConditions.presenceOfElementLocated(By.className("toast-title")));
		        // If the toast message is displayed, capture a screenshot
		        File screenshotFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		        String path = System.getProperty("user.dir") + "\\src\\test\\LoginFailed.png";
		        FileUtils.copyFile(screenshotFile, new File(path));
		        System.out.println("Screenshot captured and saved successfully. Toast message: " + toastMessage.getText());
		    } catch (org.openqa.selenium.TimeoutException e) {
		        // If the toast message is not displayed, consider the login successful
		        System.out.println("Login Successful");
		    }
		
	}
	
	
	public void ExtractCustomerDetails() {
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath("//a[contains(text(),'Customer Centres')]")))).isDisplayed();	
		WebElement Branch = driver.findElement(By.xpath("//a[contains(text(),'Customer Centres')]"));
		Branch.click();
		System.out.println(Branch.getText());
		wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath("//h4[.='Centre Names']")))).isDisplayed();
		WebElement BranchName = driver.findElement(By.xpath("//*[@id=\"modelSmallBody\"]/table/tbody/tr/td"));
		WebElement Dummy = driver.findElement(By.xpath("//*[@id=\"divFirstBlock\"]/div[1]/div[3]/div"));
		System.out.println(Dummy.getText());
		
		//close the popup
		driver.findElement(By.xpath("//*[@id=\"modelSmall\"]/div/div/div[1]/button")).click();
		
		String filePath = System.getProperty("user.dir") + "\\src\\test\\AllCloudCustomerDetails.xlsx";
        String SheetName = BranchName.getText();
		  String[] columnNames = {"Created On","CustomerName", "DOB", "Gender", "MobileNo", "Cu Name","CuCareOF", "CuAddLine1", "CuAddLine2", "CuArea", "CuCity", "CuTaluka", "CuState", "CuPincode", "CuCountry", "Par Name","ParCareOF", "ParAddLine1", "ParAddLine2", "ParArea", "ParCity", "ParTaluka", "ParState", "ParPincode", "ParCountry", "KYCDL", "KYCRationCard", "KYCCKYC", "Other"}; // Add your column names here

        try {
            createSheet(filePath, SheetName);
            System.out.println("Sheet created successfully or ignored if already exists.");
            createColumns(filePath, SheetName, columnNames);
            System.out.println("Columns created successfully or ignored if already exist.");
            enterDetailsInExcelSheet(filePath, SheetName);
            System.out.println("Details Captured");
        } catch (IOException e) {
            e.printStackTrace();
        }
	}
    


	private void createSheet(String filePath, String SheetName) throws EncryptedDocumentException, IOException {
		// TODO Auto-generated method stub
		 FileInputStream fis = new FileInputStream(new File(filePath));
	        Workbook workbook = WorkbookFactory.create(fis);

	        // Check if the sheet already exists
	        if (workbook.getSheetIndex(SheetName) == -1) {
	            // Create a new sheet
	            org.apache.poi.ss.usermodel.Sheet newSheet = workbook.createSheet(SheetName);

	            // Do additional operations on the new sheet if needed

	            // Save the changes back to the Excel file
	            try (FileOutputStream fos = new FileOutputStream(new File(filePath))) {
	                workbook.write(fos);
	            }
	        } else {
	            System.out.println("Sheet with name '" + SheetName + "' already exists. Ignoring.");
	        }

	        fis.close();
	    }
	
	
	    public static void createColumns(String filePath, String sheetName, String[] columnNames) throws IOException {
	        FileInputStream fis = new FileInputStream(new File(filePath));
	        Workbook workbook = WorkbookFactory.create(fis);

	        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet(sheetName);

	        if (sheet == null) {
	            // If the sheet doesn't exist, create a new sheet
	            sheet = workbook.createSheet(sheetName);
	        }

	        Row headerRow = sheet.getRow(0);
	        if (headerRow == null) {
	            // If the header row doesn't exist, create a new row for column names
	            headerRow = sheet.createRow(0);

	            // Create columns with field names
	            for (int i = 0; i < columnNames.length; i++) {
	                org.apache.poi.ss.usermodel.Cell cell = headerRow.createCell(i);
	                cell.setCellValue(columnNames[i]);
	            }

	            // Do additional operations on the header row if needed

	            // Save the changes back to the Excel file
	            try (FileOutputStream fos = new FileOutputStream(new File(filePath))) {
	                workbook.write(fos);
	            }
	        } else {
	            System.out.println("Columns already exist. Ignoring.");
	        }

	        fis.close();
	}
	    
	    public void enterDetailsInExcelSheet(String filePath, String sheetName) throws IOException {
	        FileInputStream fis = new FileInputStream(filePath);
	        Workbook workbook = new XSSFWorkbook(fis);
	        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet(sheetName); // Specify the sheet name

	        // Example: Find a WebElement by its ID (you can change this based on your webpage structure)
	        WebElement CreatedOn = driver.findElement(By.xpath("//*[@id=\"divSecondBlock\"]/div[1]/div")); 
	        WebElement CustomerName = driver.findElement(By.xpath("//*[@id=\"divFirstBlock\"]/div[1]/div[6]/div"));
	        WebElement DOB = driver.findElement(By.xpath("//*[@id=\"divFirstBlock\"]/div[1]/div[3]/div"));
	        WebElement Gender = driver.findElement(By.xpath("//*[@id=\"divFirstBlock\"]/div[1]/div[2]/div"));
	        WebElement MobileNo = driver.findElement(By.xpath("//*[@id=\"divSecondBlock\"]/div[3]/div"));
	        WebElement CuName = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[1]/div[2]"));
	        WebElement CuCareOF = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[1]/div[4]"));
	        WebElement CuAddLine1 = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[1]/div[6]"));
	        WebElement CuAddLine2 = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[1]/div[8]"));
	        WebElement CuArea = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[1]/div[10]"));
	        WebElement CuCity = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[1]/div[12]"));
	        WebElement CuTaluka = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[1]/div[14]"));
	        WebElement CuState = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[1]/div[16]"));
	        WebElement CuPincode = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[1]/div[18]"));
	        WebElement CuCountry = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[1]/div[26]"));
	        WebElement ParName = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[2]/div[2]"));
	        WebElement ParCareOF = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[2]/div[4]"));
	        WebElement ParAddLine1 = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[2]/div[6]"));
	        WebElement ParAddLine2 = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[2]/div[8]"));
	        WebElement ParArea = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[2]/div[10]"));
	        WebElement ParCity = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[2]/div[12]"));
	        WebElement ParTaluka = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[2]/div[14]"));
	        WebElement ParState = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[2]/div[16]"));
	        WebElement ParPincode = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[2]/div[18]"));
	        WebElement ParCountry = driver.findElement(By.xpath("//*[@id=\"divThirdBlock\"]/div[2]/div[26]"));
	        WebElement KYCDL = driver.findElement(By.xpath("//*[@id=\"divFifthBlock\"]/div[3]/div"));
	        WebElement KYCRationCard = driver.findElement(By.xpath("//*[@id=\"divFifthBlock\"]/div[2]/div[1]"));
	        WebElement KYCCKYC = driver.findElement(By.xpath("//*[@id=\"divFifthBlock\"]/div[5]/div"));
	        WebElement Other = driver.findElement(By.xpath("//*[@id=\"divFifthBlock\"]/div[9]/label/b"));
	        

	        // Get the text from the WebElement
	        String CreatedOnText = CreatedOn.getText();
	        String CustomerNameText = CustomerName.getText();
	        String DOBText = DOB.getText();
	        String GenderText = Gender.getText();
	        String MobileNoText = MobileNo.getText();
	        String CuNameText = CuName.getText();
	        String CuCareOFText = CuCareOF.getText();
	        String CuAddLine1Text = CuAddLine1.getText();
	        String CuAddLine2Text = CuAddLine2.getText();
	        String CuAreaText = CuArea.getText();
	        String CuCityText = CuCity.getText();
	        String CuTalukaText = CuTaluka.getText();
	        String CuStateText = CuState.getText();
	        String CuPincodeText = CuPincode.getText();
	        String CuCountryText = CuCountry.getText();
	        String ParNameText = ParName.getText();
	        String ParCareOFText = ParCareOF.getText();
	        String ParAddLine1Text = ParAddLine1.getText();
	        String ParAddLine2Text = ParAddLine2.getText();
	        String ParAreaText = ParArea.getText();
	        String ParCityText = ParCity.getText();
	        String ParTalukaText = ParTaluka.getText();
	        String ParStateText = ParState.getText();
	        String ParPincodeText = ParPincode.getText();
	        String ParCountryText = ParCountry.getText();
	        String KYCDLText = KYCDL.getText();
	        String KYCRationCardText = KYCRationCard.getText();
	        String KYCCKYCText = KYCCKYC.getText();
	        String OtherText = Other.getText();

	        // Find the first blank row and store the text in the first cell
	        int rowNum = findFirstBlankRow(sheet);
	        Row row = sheet.getRow(rowNum);
	        if (row == null) {
	            // If the row doesn't exist, create it
	            row = sheet.createRow(rowNum);
	        }
	        org.apache.poi.ss.usermodel.Cell cell0 = row.createCell(0); 
	        cell0.setCellValue(CreatedOnText);
	        org.apache.poi.ss.usermodel.Cell cell1 = row.createCell(1); 
	        cell1.setCellValue(CustomerNameText);
	        org.apache.poi.ss.usermodel.Cell cell2 = row.createCell(2); 
	        cell2.setCellValue(DOBText);
	        org.apache.poi.ss.usermodel.Cell cell3 = row.createCell(3); 
	        cell3.setCellValue(GenderText);
	        org.apache.poi.ss.usermodel.Cell cell4 = row.createCell(4); 
	        cell4.setCellValue(MobileNoText);
	        org.apache.poi.ss.usermodel.Cell cell5 = row.createCell(5); 
	        cell5.setCellValue(CuNameText);
	        org.apache.poi.ss.usermodel.Cell cell6 = row.createCell(6); 
	        cell6.setCellValue(CuCareOFText);
	        org.apache.poi.ss.usermodel.Cell cell7 = row.createCell(7); 
	        cell7.setCellValue(CuAddLine1Text);
	        org.apache.poi.ss.usermodel.Cell cell8 = row.createCell(8); 
	        cell8.setCellValue(CuAddLine2Text);
	        org.apache.poi.ss.usermodel.Cell cell9 = row.createCell(9); 
	        cell9.setCellValue(CuAreaText);
	        org.apache.poi.ss.usermodel.Cell cell10 = row.createCell(10); 
	        cell10.setCellValue(CuCityText);
	        org.apache.poi.ss.usermodel.Cell cell11 = row.createCell(11); 
	        cell11.setCellValue(CuTalukaText);
	        org.apache.poi.ss.usermodel.Cell cell12 = row.createCell(12); 
	        cell12.setCellValue(CuStateText);
	        org.apache.poi.ss.usermodel.Cell cell13 = row.createCell(13); 
	        cell13.setCellValue(CuPincodeText);
	        org.apache.poi.ss.usermodel.Cell cell14 = row.createCell(14); 
	        cell14.setCellValue(CuCountryText);
	        org.apache.poi.ss.usermodel.Cell cell15 = row.createCell(15); 
	        cell15.setCellValue(ParNameText);
	        org.apache.poi.ss.usermodel.Cell cell16 = row.createCell(16); 
	        cell16.setCellValue(ParCareOFText);
	        org.apache.poi.ss.usermodel.Cell cell17 = row.createCell(17); 
	        cell17.setCellValue(ParAddLine1Text);
	        org.apache.poi.ss.usermodel.Cell cell18 = row.createCell(18); 
	        cell18.setCellValue(ParAddLine2Text);
	        org.apache.poi.ss.usermodel.Cell cell19 = row.createCell(19); 
	        cell19.setCellValue(ParAreaText);
	        org.apache.poi.ss.usermodel.Cell cell20 = row.createCell(20); 
	        cell20.setCellValue(ParCityText);
	        org.apache.poi.ss.usermodel.Cell cell21 = row.createCell(21); 
	        cell21.setCellValue(ParTalukaText);
	        org.apache.poi.ss.usermodel.Cell cell22 = row.createCell(22); 
	        cell22.setCellValue(ParStateText);
	        org.apache.poi.ss.usermodel.Cell cell23 = row.createCell(23); 
	        cell23.setCellValue(ParPincodeText);
	        org.apache.poi.ss.usermodel.Cell cell24 = row.createCell(24); 
	        cell24.setCellValue(ParCountryText);
	        org.apache.poi.ss.usermodel.Cell cell25 = row.createCell(25); 
	        cell25.setCellValue(KYCDLText);
	        org.apache.poi.ss.usermodel.Cell cell26 = row.createCell(26); 
	        cell26.setCellValue(KYCRationCardText);
	        org.apache.poi.ss.usermodel.Cell cell27 = row.createCell(27); 
	        cell27.setCellValue(KYCCKYCText);
	        org.apache.poi.ss.usermodel.Cell cell28 = row.createCell(28); 
	        cell28.setCellValue(OtherText);

	        // Save the changes back to the Excel file
	        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
	        workbook.write(fileOutputStream);
	        fileOutputStream.close();
	    }

	    private static int findFirstBlankRow(org.apache.poi.ss.usermodel.Sheet sheet) {
	        int lastRowNum = sheet.getLastRowNum();

	        for (int i = 0; i <= lastRowNum; i++) {
	            Row row = sheet.getRow(i);

	            // Check if the first cell of the row is blank
	            org.apache.poi.ss.usermodel.Cell cell = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
	            if (cell.getCellType() == CellType.BLANK) {
	                return i;
	            }
	        }

	        // If no blank row is found, return the next row
	        return lastRowNum + 1;
	    }
}
	
	
	
	
	


