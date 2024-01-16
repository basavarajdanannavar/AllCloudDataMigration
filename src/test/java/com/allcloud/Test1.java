package com.allcloud;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Test1 {
	public WebDriver driver;
	public Workbook workbook;
	public Sheet sheet;

	@SuppressWarnings("deprecation")
	@BeforeClass
	public void setUp() {
		// Set the path to ChromeDriver based on your system configuration
		ChromeOptions options = new ChromeOptions();
		options.setAcceptInsecureCerts(true);

		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver(options);
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

	}

	@Test(priority = 0)
	public void RememberMe() throws InterruptedException, IOException {
		FileInputStream file = new FileInputStream(".//DataFiles//Login.xlsx");
        Workbook workbook = WorkbookFactory.create(file);
        @SuppressWarnings("rawtypes")
		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("sheet1");

        // Loop through the rows and perform the login test cases
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String username = getCellValueAsString(row.getCell(0));
            String password = getCellValueAsString(row.getCell(1));
            
            
            driver.get("https://tickets.technovative.in/scp/login.php");

    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

    		// Enter username and password
    		WebElement usernameField = driver.findElement(By.id("name"));
    		WebElement passwordField = driver.findElement(By.id("pass"));
    		usernameField.sendKeys(username);
    		passwordField.sendKeys(password);

    		// Submit the form
    		WebElement loginButton = driver.findElement(By.cssSelector("[type='submit']"));
    		loginButton.click();
    		
    		
    		Thread.sleep(3000);
    		driver.findElement(By.cssSelector("[href='/scp/admin.php']")).click();
    		Thread.sleep(300);
    		driver.findElement(By.cssSelector("[href='/scp/staff.php']")).click();
    		Thread.sleep(300);
    		driver.findElement(By.xpath("//*[@id=\"content\"]/div[2]/div[2]/div/div[2]/a")).click();
    		
		}
	}

	@Test(priority = 1)
	public void ClickOnCreateClient() throws InterruptedException, EncryptedDocumentException, IOException, AWTException {
		FileInputStream file = new FileInputStream(".//DataFiles//TicketUserlist.xlsx");
        Workbook workbook = WorkbookFactory.create(file);
        @SuppressWarnings("rawtypes")
		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("sheet1");

        // Loop through the rows and perform the login test cases
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String email1 = getCellValueAsString(row.getCell(1));
            String name1 = getCellValueAsString(row.getCell(0));
            String last = getCellValueAsString(row.getCell(2));
        
  WebElement name=driver.findElement(By.xpath("//*[@id=\"account\"]/table/tbody[1]/tr/td/div/table/tbody/tr[1]/td[2]/input[1]"));
  WebElement email=driver.findElement(By.xpath("//*[@id=\"account\"]/table/tbody[1]/tr/td/div/table/tbody/tr[2]/td[2]/input"));
  WebElement lastname=driver.findElement(By.xpath("//*[@id=\"account\"]/table/tbody[1]/tr/td/div/table/tbody/tr[1]/td[2]/input[2]"));
  WebElement username=driver.findElement(By.xpath("//*[@id=\"account\"]/table/tbody[2]/tr[2]/td[2]/input"));
             
  
  name.sendKeys(name1);
  email.sendKeys(email1);
  lastname.sendKeys(last);
  username.sendKeys(name1);
  
  Thread.sleep(2000);
  
/*  driver.findElement(By.xpath("//*[@id=\"account\"]/table/tbody[2]/tr[2]/td[2]/button")).click();
  
  Thread.sleep(1000);
 
 
 Robot r=new Robot();
 
 r.mousePress(KeyEvent.VK_ESCAPE);
 r.mouseRelease(KeyEvent.VK_ESCAPE);

  r.mousePress(KeyEvent.VK_TAB);
  r.mouseRelease(KeyEvent.VK_TAB);
  Thread.sleep(300);
  r.mousePress(KeyEvent.VK_TAB);
  r.mouseRelease(KeyEvent.VK_TAB);
  Thread.sleep(300);
  r.mousePress(KeyEvent.VK_TAB);
  r.mouseRelease(KeyEvent.VK_TAB);
  Thread.sleep(300);
  r.mousePress(KeyEvent.VK_TAB);
  r.mouseRelease(KeyEvent.VK_TAB);
  Thread.sleep(300);
  r.mousePress(KeyEvent.VK_TAB);
  r.mouseRelease(KeyEvent.VK_TAB);
  Thread.sleep(300);
  r.mousePress(KeyEvent.VK_TAB);
  r.mouseRelease(KeyEvent.VK_TAB);
  Thread.sleep(300);
  r.mousePress(KeyEvent.VK_TAB);
  r.mouseRelease(KeyEvent.VK_TAB);
  Thread.sleep(300);
  r.mousePress(KeyEvent.VK_TAB);
  r.mouseRelease(KeyEvent.VK_TAB);
  Thread.sleep(300);
  r.mousePress(KeyEvent.VK_TAB);
  r.mouseRelease(KeyEvent.VK_TAB);
  Thread.sleep(300);
  r.mousePress(KeyEvent.VK_TAB);
  r.mouseRelease(KeyEvent.VK_TAB);
  Thread.sleep(300);
  r.mousePress(KeyEvent.VK_TAB);
  r.mouseRelease(KeyEvent.VK_TAB);
  Thread.sleep(2000);
  
  driver.findElement(By.xpath("//*[@id=\"_4ff6ea51d5e670\"]")).sendKeys("123456");
  
  driver.findElement(By.xpath("//*[@id=\"_80bd0c9b1a20fa\"]")).sendKeys("123456");
  
  Thread.sleep(500);
  
  driver.findElement(By.xpath("//*[@id=\"field_0cb8403c5695cf\"]/label/label")).click();
  Thread.sleep(500);
  
  driver.findElement(By.xpath("//*[@id=\"popup\"]/div[2]/form/p/span[2]/input")).click();  */
            
            
            driver.findElement(By.cssSelector("[href='#access']")).click();
            
            Thread.sleep(1000);
            
            driver.findElement(By.id("dept_id")).click();
            driver.findElement(By.xpath("//*[@id=\"dept_id\"]/option[4]")).click();
            driver.findElement(By.xpath("//*[@id=\"access\"]/table/tbody[1]/tr[2]/td[2]/select/option[6]")).click();
            
            
            driver.findElement(By.id("add_access")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//*[@id=\"add_access\"]/option[5]")).click();
            driver.findElement(By.xpath("//*[@id=\"add_extended_access\"]/td/button")).click();
            driver.findElement(By.xpath("//*[@id=\"access\"]/table/tbody[3]/tr[2]/td[2]/select/option[6]")).click();
            
            Thread.sleep(2000);
            driver.findElement(By.id("add_access")).click();
            driver.findElement(By.xpath("//*[@id=\"add_access\"]/option[5]")).click();
            driver.findElement(By.xpath("//*[@id=\"add_extended_access\"]/td/button")).click();
            driver.findElement(By.xpath("//*[@id=\"access\"]/table/tbody[3]/tr[3]/td[2]/select/option[6]")).click();
            Thread.sleep(2000);
            driver.findElement(By.id("add_access")).click();
            driver.findElement(By.xpath("//*[@id=\"add_access\"]/option[2]")).click();
            driver.findElement(By.xpath("//*[@id=\"add_extended_access\"]/td/button")).click();
            driver.findElement(By.xpath("//*[@id=\"access\"]/table/tbody[3]/tr[4]/td[2]/select/option[6]")).click();
            
            Thread.sleep(2000);
            
            driver.findElement(By.xpath("//*[@id=\"content\"]/form/p/input[1]")).click();
            Thread.sleep(5000);
        	driver.findElement(By.cssSelector("[href='/scp/staff.php']")).click();
    		Thread.sleep(300);
    		driver.findElement(By.xpath("//*[@id=\"content\"]/div[2]/div[2]/div/div[2]/a")).click();
    		
        
        
        }

	}

	
			
			private static String getCellValueAsString(Cell cell) {
		if (cell == null) {
			return "";
		}

		if (cell.getCellType() == CellType.STRING) {
			return cell.getStringCellValue();
		} else if (cell.getCellType() == CellType.NUMERIC) {
			return String.valueOf((int) cell.getNumericCellValue());
		} else {
			return "";
		}
	}
}
