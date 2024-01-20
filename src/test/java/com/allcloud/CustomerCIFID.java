package com.allcloud;


import java.awt.AWTException;
import java.awt.RenderingHints.Key;
import java.awt.datatransfer.StringSelection;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.NoSuchElementException;
import java.util.Set;

import org.apache.commons.io.FileUtils;
import org.apache.commons.math3.analysis.function.Exp;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
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
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CustomerCIFID {

    WebDriver driver;
    String cellValue;
   
    public void setup() {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.manage().deleteAllCookies();
 }
   

    public void processExcel() throws IOException, InterruptedException {
        String filePath = System.getProperty("user.dir") + "\\src\\test\\CIFData.xlsx";
        FileInputStream fileInputStream = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("Sheet1");

        int startRowIndex = 0;
        int columnIndex = 0;

        readCellValue(sheet, startRowIndex, columnIndex);

        // Perform actions with the found cellValue
        if (cellValue != null) {
            String url = "https://apps.allcloud.in/magfinserv/Customer/CustomerDetails/" + cellValue;
            driver.get(url);
        }

        clearCellValue(sheet, cellValue, fileInputStream, filePath, workbook);
    }

    private void readCellValue(org.apache.poi.ss.usermodel.Sheet sheet, int startRowIndex, int columnIndex) {
        String cellValue = null;

        for (int rowIndex = startRowIndex; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    cellValue = cell.getStringCellValue();
                }
            }
            if (cellValue != null && !cellValue.isEmpty()) {
                break;
            }
        }

        System.out.println(cellValue);
        this.cellValue = cellValue;
    }

    private void clearCellValue(org.apache.poi.ss.usermodel.Sheet sheet, String targetValue,
            FileInputStream fileInputStream, String filePath, Workbook workbook) throws IOException {
        Cell foundCell = findCell(sheet, targetValue);

        if (foundCell != null) {
            int rowIndex2 = foundCell.getRowIndex();
            int columnIndex2 = foundCell.getColumnIndex();

            System.out.println(rowIndex2);
            System.out.println(columnIndex2);

            Row row = sheet.getRow(rowIndex2);
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
    }

    private Cell findCell(org.apache.poi.ss.usermodel.Sheet sheet, String targetValue) {
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
        return foundCell;
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
	
	
	public void ExtractCustomerDetails() throws InterruptedException, AWTException {
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath("//a[contains(text(),'Customer Centres')]")))).isDisplayed();	

  
		
		String filePath = System.getProperty("user.dir") + "\\src\\test\\AllCloudCustomerDetails.xlsx";
        String SheetName = "CustomerDetails";
		  try {
           
            enterDetailsInExcelSheet(filePath, SheetName);
           
           
            
            
        } catch (IOException e) {
            e.printStackTrace();
        }
	}
	
	public void CustomerImage() throws AWTException, InterruptedException {
		try {
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			String ParentWindow = driver.getWindowHandle();

			WebElement customerImage = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//i[contains(@class,'fa fa-cloud-download')]")));
			customerImage.isDisplayed();
			customerImage.click();

    Set<String> handles = driver.getWindowHandles();
    
    for (String handle :handles ) {
    	if(!handle.equals(ParentWindow)) {
    		driver.switchTo().window(handle);
    		Robot r = new Robot();
			Thread.sleep(2000);
			r.keyPress(KeyEvent.VK_CONTROL);
			r.keyPress(KeyEvent.VK_S);
			r.keyRelease(KeyEvent.VK_CONTROL);
			r.keyRelease(KeyEvent.VK_S);
			Thread.sleep(1000);

			r.keyPress(KeyEvent.VK_TAB);
			r.keyRelease(KeyEvent.VK_TAB);
			Thread.sleep(1000);
			r.keyPress(KeyEvent.VK_TAB);
			r.keyRelease(KeyEvent.VK_TAB);
			Thread.sleep(1000);
			r.keyPress(KeyEvent.VK_TAB);
			r.keyRelease(KeyEvent.VK_TAB);
			Thread.sleep(1000);
			r.keyPress(KeyEvent.VK_TAB);
			r.keyRelease(KeyEvent.VK_TAB);
			Thread.sleep(1000);
			r.keyPress(KeyEvent.VK_TAB);
			r.keyRelease(KeyEvent.VK_TAB);
			Thread.sleep(2000);
			r.keyPress(KeyEvent.VK_TAB);
			r.keyRelease(KeyEvent.VK_TAB);
			Thread.sleep(1000);


			r.keyPress(KeyEvent.VK_ENTER);
			r.keyRelease(KeyEvent.VK_ENTER);
			Thread.sleep(1000);
			
			String Folderpath = System.getProperty("user.dir") + "\\src\\test\\CustomerImages";
			StringSelection check=  new StringSelection(Folderpath);
			  Toolkit.getDefaultToolkit().getSystemClipboard().setContents(check, null);
r.keyPress(KeyEvent.VK_CONTROL);
			  r.keyPress(KeyEvent.VK_V);
			  r.keyRelease(KeyEvent.VK_CONTROL);
			  r.keyRelease(KeyEvent.VK_V);
			  

			  
			  Thread.sleep(1000);
			  r.keyPress(KeyEvent.VK_ENTER);
			  r.keyRelease(KeyEvent.VK_ENTER);
			  Thread.sleep(1000);
	//		  r.keyPress(KeyEvent.VK_ENTER);

			  //		  r.keyRelease(KeyEvent.VK_ENTER);
	//		  Thread.sleep(1000);

			  r.keyPress(KeyEvent.VK_ENTER);

			  r.keyRelease(KeyEvent.VK_ENTER);
			  Thread.sleep(1000);

			  r.keyPress(KeyEvent.VK_ENTER);
			  r.keyRelease(KeyEvent.VK_ENTER);
			  Thread.sleep(1000);
			  r.keyPress(KeyEvent.VK_ENTER);
			  r.keyRelease(KeyEvent.VK_ENTER);
			  Thread.sleep(1000);
			  r.keyPress(KeyEvent.VK_ENTER);
			  r.keyRelease(KeyEvent.VK_ENTER);
			  Thread.sleep(1000);
			  
			  driver.close();
			  Thread.sleep(3000);
			  driver.switchTo().window(ParentWindow);
    	}

			}
		}
			catch (org.openqa.selenium.TimeoutException e) {
			    // If the toast message is not displayed, consider the login successful
			    System.out.println("No Image found");
			}
    	}
    
			
	
    
	
	   
	    
	    public void enterDetailsInExcelSheet(String filePath, String sheetName) throws IOException, NoSuchElementException {
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

	    	FileInputStream fis = new FileInputStream(filePath);
	        Workbook workbook = new XSSFWorkbook(fis);
	        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet(sheetName); // Specify the sheet name

	        // Example: Find a WebElement by its ID (you can change this based on your webpage structure)
	        WebElement Branch = driver.findElement(By.xpath("//a[contains(text(),'Customer Centres')]"));
			Branch.click();

			wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath("//h4[.='Centre Names']")))).isDisplayed();
			WebElement BranchName = driver.findElement(By.xpath("//*[@id=\"modelSmallBody\"]/table/tbody/tr/td"));
			  // Find the first blank row and store the text in the first cell
	        int rowNum = findFirstBlankRow(sheet);
	        Row row = sheet.getRow(rowNum);
	        if (row == null) {
	            // If the row doesn't exist, create it
	            row = sheet.createRow(rowNum);
	        }
	        String BranchNameText = BranchName.getText();
	        org.apache.poi.ss.usermodel.Cell cell1 = row.createCell(1); 
	        cell1.setCellValue(BranchNameText);
			
			//close the popup
			wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath("//*[@id=\"modelSmall\"]/div/div/div[1]/button"))));
			driver.findElement(By.xpath("//*[@id=\"modelSmall\"]/div/div/div[1]/button")).click();
	      
			WebElement CIFID = driver.findElement(By.xpath("//*[@id=\"divFirstBlock\"]/div[1]/div[1]/div"));
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
	        String CIFIDText = CIFID.getText();
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

	      
	        org.apache.poi.ss.usermodel.Cell cell0 = row.createCell(0); 
	        cell0.setCellValue(CIFIDText);
	   
	        System.out.println(BranchNameText);
	        org.apache.poi.ss.usermodel.Cell cell2 = row.createCell(2); 
	        cell2.setCellValue(CreatedOnText);
	        org.apache.poi.ss.usermodel.Cell cell3 = row.createCell(3); 
	        cell3.setCellValue(CustomerNameText);
	        org.apache.poi.ss.usermodel.Cell cell4 = row.createCell(4); 
	        cell4.setCellValue(DOBText);
	        org.apache.poi.ss.usermodel.Cell cell5 = row.createCell(5); 
	        cell5.setCellValue(GenderText);
	        org.apache.poi.ss.usermodel.Cell cell6 = row.createCell(6); 
	        cell6.setCellValue(MobileNoText);
	        org.apache.poi.ss.usermodel.Cell cell7 = row.createCell(7); 
	        cell7.setCellValue(CuNameText);
	        org.apache.poi.ss.usermodel.Cell cell8 = row.createCell(8); 
	        cell8.setCellValue(CuCareOFText);
	        org.apache.poi.ss.usermodel.Cell cell9 = row.createCell(9); 
	        cell9.setCellValue(CuAddLine1Text);
	        org.apache.poi.ss.usermodel.Cell cell10 = row.createCell(10); 
	        cell10.setCellValue(CuAddLine2Text);
	        org.apache.poi.ss.usermodel.Cell cell11 = row.createCell(11); 
	        cell11.setCellValue(CuAreaText);
	        org.apache.poi.ss.usermodel.Cell cell12 = row.createCell(12); 
	        cell12.setCellValue(CuCityText);
	        org.apache.poi.ss.usermodel.Cell cell13 = row.createCell(13); 
	        cell13.setCellValue(CuTalukaText);
	        org.apache.poi.ss.usermodel.Cell cell14 = row.createCell(14); 
	        cell14.setCellValue(CuStateText);
	        org.apache.poi.ss.usermodel.Cell cell15 = row.createCell(15); 
	        cell15.setCellValue(CuPincodeText);
	        org.apache.poi.ss.usermodel.Cell cell16 = row.createCell(16); 
	        cell16.setCellValue(CuCountryText);
	        org.apache.poi.ss.usermodel.Cell cell17 = row.createCell(17); 
	        cell17.setCellValue(ParNameText);
	        org.apache.poi.ss.usermodel.Cell cell18 = row.createCell(18); 
	        cell18.setCellValue(ParCareOFText);
	        org.apache.poi.ss.usermodel.Cell cell19 = row.createCell(19); 
	        cell19.setCellValue(ParAddLine1Text);
	        org.apache.poi.ss.usermodel.Cell cell20 = row.createCell(20); 
	        cell20.setCellValue(ParAddLine2Text);
	        org.apache.poi.ss.usermodel.Cell cell21 = row.createCell(21); 
	        cell21.setCellValue(ParAreaText);
	        org.apache.poi.ss.usermodel.Cell cell22 = row.createCell(22); 
	        cell22.setCellValue(ParCityText);
	        org.apache.poi.ss.usermodel.Cell cell23 = row.createCell(23); 
	        cell23.setCellValue(ParTalukaText);
	        org.apache.poi.ss.usermodel.Cell cell24 = row.createCell(24); 
	        cell24.setCellValue(ParStateText);
	        org.apache.poi.ss.usermodel.Cell cell25 = row.createCell(25); 
	        cell25.setCellValue(ParPincodeText);
	        org.apache.poi.ss.usermodel.Cell cell26 = row.createCell(26); 
	        cell26.setCellValue(ParCountryText);
	        org.apache.poi.ss.usermodel.Cell cell27 = row.createCell(27); 
	        cell27.setCellValue(KYCDLText);
	        org.apache.poi.ss.usermodel.Cell cell28 = row.createCell(28); 
	        cell28.setCellValue(KYCRationCardText);
	        org.apache.poi.ss.usermodel.Cell cell29 = row.createCell(29); 
	        cell29.setCellValue(KYCCKYCText);
	        org.apache.poi.ss.usermodel.Cell cell30 = row.createCell(30); 
	        cell30.setCellValue(OtherText);
	        try {
	        	try {
	    			WebDriverWait wait3 = new WebDriverWait(driver, Duration.ofSeconds(03));

	    			WebElement customerImage = wait3.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//i[contains(@class,'fa fa-cloud-download')]")));
	    			customerImage.isDisplayed();
	    			org.apache.poi.ss.usermodel.Cell cell31 = row.createCell(31); 
	    		        cell31.setCellValue(CIFIDText + ".jpg");
	    		        CustomerImage();
	        } catch (Exception e) {
	          System.out.println("Customer Image not found");
	        	// Handle NoSuchElementException if the element is not found
	          //  e.printStackTrace(); // Or log the error message
	        }
	        	try {
	    			WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(03));
	    			WebElement KYC1 = wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[2]/div[2]/div/div/div/div/div/div/div/div[3]/div/div[8]/div/div/div/table/tbody/tr[1]/td[1]")));
                    KYC1.isDisplayed();
               	 org.apache.poi.ss.usermodel.Cell cell32 = row.createCell(32); 
    		        cell32.setCellValue(CIFIDText + KYC1.getText()+ ".pdf");
    		        CustomerKYC1();
    	           
        			
        		}catch (Exception e) {
      	          System.out.println("KYC1 Image not found");
  	        	// Handle NoSuchElementException if the element is not found
  	          //  e.printStackTrace(); // Or log the error message
  	        }
	        	try {
	    			WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(3));
	    			WebElement KYC2 = wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[2]/div[2]/div/div/div/div/div/div/div/div[3]/div/div[8]/div/div/div/table/tbody/tr[2]/td[1]")));
                    KYC2.isDisplayed();
               	 org.apache.poi.ss.usermodel.Cell cell33 = row.createCell(33); 
    		        cell33.setCellValue(CIFIDText + KYC2.getText()+ ".pdf");
    		        CustomerKYC2();           
    	           
        			
        		}catch (Exception e) {
      	          System.out.println("KYC2 Image not found");
  	        	// Handle NoSuchElementException if the element is not found
  	          //  e.printStackTrace(); // Or log the error message
  	        }
	        	try {
	    			WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(3));
	    			WebElement KYC3 = wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[2]/div[2]/div/div/div/div/div/div/div/div[3]/div/div[8]/div/div/div/table/tbody/tr[3]/td[1]")));
                    KYC3.isDisplayed();
               	 org.apache.poi.ss.usermodel.Cell cell34 = row.createCell(34); 
    		        cell34.setCellValue(CIFIDText + KYC3.getText()+ ".pdf");
    		        CustomerKYC3();
        		}catch (Exception e) {
      	          System.out.println("KYC3 Image not found");
  	        	// Handle NoSuchElementException if the element is not found
  	          //  e.printStackTrace(); // Or log the error message
  	        }
	        	
	       

	        // Save the changes back to the Excel file
	        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
	        workbook.write(fileOutputStream);
	        fileOutputStream.close();
	        }
	        finally {
	          
	        	// Add any cleanup steps here
	            // This block will be executed whether an exception occurs or not
	        }}
	        	
	    

	    private void printStackTrace() {
			// TODO Auto-generated method stub
			
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

	    public void CIFIDConfirm() {
	        try {
	            WebElement error1 = driver.findElement(By.xpath("//h2[@class='error']"));

	            while (error1.isDisplayed()) {
	                
	                // Assuming setup(), processExcel(), and OpenBrowser() are custom methods
	                // Fill in the details of these methods according to your implementation
	                setup();
	                processExcel();
	                OpenBrowser();
	                WebElement error2 = driver.findElement(By.xpath("//h2[@class='error']"));
	                error2.isDisplayed();
	                driver.quit();
	               
	            }
	            System.out.println("Customer Details found");
	        } catch (NoSuchElementException e) {
	            // Handle NoSuchElementException if the element is not found
	            e.printStackTrace(); // Or log the error message
	        } catch (Exception e) {
	            // Handle other exceptions
	            e.printStackTrace(); // Or log the error message
	        }
	    }

	    public void CustomerKYC1() throws AWTException, InterruptedException {
	    	
try {
	
	WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(3));

	WebElement KYC1 = wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[2]/div[2]/div/div/div/div/div/div/div/div[3]/div/div[8]/div/div/div/table/tbody/tr[1]/td[1]")));
KYC1.isDisplayed();
String ParentWindow = driver.getWindowHandle();
	WebElement ClickOnKYC1 = driver.findElement(By.xpath("/html/body/div[2]/div[2]/div/div/div/div/div/div/div/div[3]/div/div[8]/div/div/div/table/tbody/tr[1]/td[2]/a"));
	ClickOnKYC1.click();
	wait1.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath("/html/body/div[5]/div/div/div[2]/div/div/table/tbody/tr/td[5]/a[1]"))));
	WebElement ClickOnDownloadLink = driver.findElement(By.xpath("/html/body/div[5]/div/div/div[2]/div/div/table/tbody/tr/td[5]/a[1]"));
	ClickOnDownloadLink.click();
	 Set<String> handles = driver.getWindowHandles();
	    
	    Robot r = new Robot();
	Thread.sleep(1000);
	r.keyPress(KeyEvent.VK_CONTROL);
	r.keyPress(KeyEvent.VK_S);
	r.keyRelease(KeyEvent.VK_CONTROL);
	r.keyRelease(KeyEvent.VK_S);
	Thread.sleep(1000);
	    	
	
	    	
	WebElement CIFID = driver.findElement(By.xpath("//*[@id=\"divFirstBlock\"]/div[1]/div[1]/div"));
	String ChangeName1  = CIFID.getText() + KYC1.getText();

	StringSelection check1=  new StringSelection(ChangeName1);
	  Toolkit.getDefaultToolkit().getSystemClipboard().setContents(check1, null);
	  
	  r.keyPress(KeyEvent.VK_CONTROL);
	  r.keyPress(KeyEvent.VK_V);
	  r.keyRelease(KeyEvent.VK_CONTROL);
	  r.keyRelease(KeyEvent.VK_V);
	  
	  Thread.sleep(2000);
	r.keyPress(KeyEvent.VK_ENTER);
	r.keyRelease(KeyEvent.VK_ENTER);
	Thread.sleep(1000);
	
	
	  

	  for (String handle :handles ) {
	    	if(!handle.equals(ParentWindow)) {
	    		driver.switchTo().window(handle); 
	 
	  
	  driver.close();
	  
	  driver.switchTo().window(ParentWindow);
	  
	  WebElement Close = driver.findElement(By.xpath("//button[.='Close']"));
	  wait1.until(ExpectedConditions.visibilityOf(Close));
	  Close.click();
	  

	}
	    }}
	    
catch (org.openqa.selenium.TimeoutException e) {
    // If the toast message is not displayed, consider the login successful
    System.out.println("No Image found");
}
	    }

public void CustomerKYC2() throws AWTException, InterruptedException {
	    	
	    	try {
	    		
	    		WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(5));
	    		String ParentWindow = driver.getWindowHandle();
	    		WebElement KYC2 = wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[2]/div[2]/div/div/div/div/div/div/div/div[3]/div/div[8]/div/div/div/table/tbody/tr[2]/td[1]")));
	    	KYC2.isDisplayed();
	    		WebElement ClickOnKYC2 = driver.findElement(By.xpath("/html/body/div[2]/div[2]/div/div/div/div/div/div/div/div[3]/div/div[8]/div/div/div/table/tbody/tr[2]/td[2]/a"));
	    		ClickOnKYC2.click();
	    		wait1.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath("/html/body/div[5]/div/div/div[2]/div/div/table/tbody/tr/td[5]/a[1]"))));
	    		WebElement ClickOnDownloadLink = driver.findElement(By.xpath("/html/body/div[5]/div/div/div[2]/div/div/table/tbody/tr/td[5]/a[1]"));
	    		ClickOnDownloadLink.click();
	    		 Set<String> handles = driver.getWindowHandles();
	    		    
	    		    Robot r = new Robot();
	    		Thread.sleep(500);
	    		r.keyPress(KeyEvent.VK_CONTROL);
	    		r.keyPress(KeyEvent.VK_S);
	    		r.keyRelease(KeyEvent.VK_CONTROL);
	    		r.keyRelease(KeyEvent.VK_S);
	    		Thread.sleep(500);
	    		    	
	    		
	    		    	
	    		WebElement CIFID = driver.findElement(By.xpath("//*[@id=\"divFirstBlock\"]/div[1]/div[1]/div"));
	    		String ChangeName2  = CIFID.getText() + KYC2.getText();

	    		StringSelection check2=  new StringSelection(ChangeName2);
	    		  Toolkit.getDefaultToolkit().getSystemClipboard().setContents(check2, null);
	    		  Thread.sleep(1000);
	    		  r.keyPress(KeyEvent.VK_CONTROL);
	    		  r.keyPress(KeyEvent.VK_V);
	    		  r.keyRelease(KeyEvent.VK_CONTROL);
	    		  r.keyRelease(KeyEvent.VK_V);
	    		  
	    		  Thread.sleep(2000);
	    		r.keyPress(KeyEvent.VK_ENTER);
	    		r.keyRelease(KeyEvent.VK_ENTER);
	    		Thread.sleep(1000);
	    		
	    		
	    		  

	    		  for (String handle :handles ) {
	    		    	if(!handle.equals(ParentWindow)) {
	    		    		driver.switchTo().window(handle); 
	    		 
	    		  
	    		  driver.close();
	    		  
	    		  driver.switchTo().window(ParentWindow);
	    		  
	    		  WebElement Close = driver.findElement(By.xpath("//button[.='Close']"));
	    		  wait1.until(ExpectedConditions.visibilityOf(Close));
	    		  Close.click();
	    		  

	    		}
	    		    }}

	    		catch (org.openqa.selenium.TimeoutException e) {
	    		    // If the toast message is not displayed, consider the login successful
	    		    System.out.println("No Doucment2 found");
	    		}
}

public void CustomerKYC3() throws AWTException, InterruptedException {
	    	
try {
  	WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(5));

	WebElement KYC3 = wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[2]/div[2]/div/div/div/div/div/div/div/div[3]/div/div[8]/div/div/div/table/tbody/tr[3]/td[1]")));
	      	KYC3.isDisplayed();
    		String ParentWindow = driver.getWindowHandle();
	    		WebElement ClickOnKYC3 = driver.findElement(By.xpath("/html/body/div[2]/div[2]/div/div/div/div/div/div/div/div[3]/div/div[8]/div/div/div/table/tbody/tr[3]/td[2]/a"));
	    		ClickOnKYC3.click();
	    		wait1.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath("/html/body/div[5]/div/div/div[2]/div/div/table/tbody/tr/td[5]/a[1]"))));
	    		WebElement ClickOnDownloadLink = driver.findElement(By.xpath("/html/body/div[5]/div/div/div[2]/div/div/table/tbody/tr/td[5]/a[1]"));
	    		ClickOnDownloadLink.click();
	    		 Set<String> handles = driver.getWindowHandles();
	    		    
	    		    Robot r = new Robot();
	    		Thread.sleep(500);
	    		r.keyPress(KeyEvent.VK_CONTROL);
	    		r.keyPress(KeyEvent.VK_S);
	    		r.keyRelease(KeyEvent.VK_CONTROL);
	    		r.keyRelease(KeyEvent.VK_S);
	    		Thread.sleep(500);
	    		    	
	    		
	    		    	
	    		WebElement CIFID = driver.findElement(By.xpath("//*[@id=\"divFirstBlock\"]/div[1]/div[1]/div"));
	    		String ChangeName  = CIFID.getText() + KYC3.getText();

	    		StringSelection check=  new StringSelection(ChangeName);
	    		  Toolkit.getDefaultToolkit().getSystemClipboard().setContents(check, null);
	    		  Thread.sleep(1000);
	    		  r.keyPress(KeyEvent.VK_CONTROL);
	    		  r.keyPress(KeyEvent.VK_V);
	    		  r.keyRelease(KeyEvent.VK_CONTROL);
	    		  r.keyRelease(KeyEvent.VK_V);
	    		  
	    		  Thread.sleep(2000);
	    		r.keyPress(KeyEvent.VK_ENTER);
	    		r.keyRelease(KeyEvent.VK_ENTER);
	    		Thread.sleep(1000);
	    		
	    		
	    		  

	    		  for (String handle :handles ) {
	    		    	if(!handle.equals(ParentWindow)) {
	    		    		driver.switchTo().window(handle); 
	    		 
	    		  
	    		  driver.close();
	    		  
	    		  driver.switchTo().window(ParentWindow);
	    		  
	    		  WebElement Close = driver.findElement(By.xpath("//button[.='Close']"));
	    		  wait1.until(ExpectedConditions.visibilityOf(Close));
	    		  Close.click();
	    		  

	    		}
	    		    }
}

	    		catch (org.openqa.selenium.TimeoutException e) {
	    		    // If the toast message is not displayed, consider the login successful
	    		    System.out.println("No Doucment3 found");
	    		}
}

	    		    	

	    	
	    	
	    
    
   
   public void CIF() throws IOException, InterruptedException, AWTException {
       // Do something with the cellValue
   	   setup();
	   processExcel();
  	   OpenBrowser();
       CIFIDConfirm();
       ExtractCustomerDetails();
       driver.quit();
   }
   @Test(priority = 2)
    public void runCIFMultipleTimes() throws IOException, InterruptedException, AWTException {
	   int numberOfIterations = 10
			   ;
	   for (int i = 0; i < numberOfIterations; i++) {
            CIF();
        }
    }
}
