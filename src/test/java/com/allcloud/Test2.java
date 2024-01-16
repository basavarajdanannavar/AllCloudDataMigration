package com.allcloud;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Set;
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
import org.testng.Assert;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Test2 {
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
		FileInputStream file = new FileInputStream(".//DataFiles//AllCloudLogin.xlsx");
		Workbook workbook = WorkbookFactory.create(file);
		@SuppressWarnings("rawtypes")
		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("sheet1");

		// Loop through the rows and perform the login test cases
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			String username = getCellValueAsString(row.getCell(0));
			String password = getCellValueAsString(row.getCell(1));

			driver.get("https://apps.allcloud.in/magfinserv/Account/Login?");

			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

			// Enter username and password
			WebElement usernameField = driver.findElement(By.id("UserName"));
			WebElement passwordField = driver.findElement(By.id("Password"));
			usernameField.sendKeys(username);
			Thread.sleep(2000);
			passwordField.sendKeys(password);
			Thread.sleep(2000);

			// Submit the form
			WebElement loginButton = driver.findElement(By.cssSelector("[type='submit']"));
			loginButton.click();
		}
	}

	@Test(priority = 1)
	public void ClickOnCreateClient()
			throws InterruptedException, EncryptedDocumentException, IOException, AWTException {
		// click on MFi in Home page
		driver.findElement(By.xpath("/html/body/div[2]/div[1]/div/div/div[2]/div/div/ul/li[3]/a")).click();
		// click on Group
		driver.findElement(By.cssSelector("[href='/magfinserv/MFIGroup/Details']")).click();
		// Switch to the new tab or window
		Set<String> windowHandles = driver.getWindowHandles();
		for (String windowHandle : windowHandles) {
			driver.switchTo().window(windowHandle);

			// You can check the title or URL of the window here
			String windowTitle = driver.getTitle();
			if (windowTitle.equals("Groups"))
				;

			// click on View group
			driver.findElement(By.cssSelector("[name='SubmitButton']")).click();
			// Enter the search details
			driver.findElement(By.cssSelector("[type='search']")).sendKeys("test4");

			WebElement RecordNotFound = driver.findElement(By.xpath("//*[@id=\"Groups\"]/tbody/tr/td"));
			System.out.println(RecordNotFound.getText());

			
			
			if(RecordNotFound.getText().equals("Record not found")) {
				Creategroup();
				}
			else {
				AddMember();
			}
		}
	}

	
	public void Creategroup() throws AWTException {
		// click on add button
		driver.findElement(By.cssSelector("[href='/magfinserv/MFIGroup/AddorUpdate']")).click();

		Set<String> windowHandles1 = driver.getWindowHandles();
		for (String windowHandle1 : windowHandles1) {
			driver.switchTo().window(windowHandle1);

			// You can check the title or URL of the window here
			String windowTitle1 = driver.getTitle();
			if (windowTitle1.equals("Add Group"));

			// select Centre
			driver.findElement(By.id("select2-CentreId-container")).click();
			driver.findElement(By.xpath("/html/body/span/span/span[1]/input")).sendKeys("Wai");
			Robot r = new Robot();
			r.keyPress(KeyEvent.VK_ENTER);
			r.keyPress(KeyEvent.VK_ENTER);

//Enter the Group name
			driver.findElement(By.id("GroupName")).sendKeys("test6");
//select the scheme 
			driver.findElement(By.id("select2-SchemeId-container")).click();

			driver.findElement(By.xpath("/html/body/span/span/span[1]/input")).sendKeys("bi");
			r.keyPress(KeyEvent.VK_ENTER);
			r.keyPress(KeyEvent.VK_ENTER);

//click on Add
//driver.findElement(By.id("btnAdd")).click();

		}
	}

	
	public void AddMember() throws AWTException, InterruptedException {
		
		WebElement Groupname= driver.findElement(By.xpath("//*[@id=\"Groups\"]/tbody/tr/td[5]"));
		Assert.assertEquals(Groupname.getText(), "test4");
	driver.findElement(By.xpath("//*[@id=\"Groups\"]/tbody/tr/td[12]/a[2]")).click();
	//click on add memeber
	driver.findElement(By.id("AddMember")).click();
	
	
	
		Robot r1=new Robot();
		r1.keyPress(KeyEvent.VK_ALT);
		r1.keyPress(KeyEvent.VK_TAB);
		r1.keyRelease(KeyEvent.VK_ALT);
		r1.keyRelease(KeyEvent.VK_TAB);

		Thread.sleep(6000);
		driver.manage().window().maximize();
		
		r1.keyPress(KeyEvent.VK_ALT);
		r1.keyPress(KeyEvent.VK_TAB);
		r1.keyRelease(KeyEvent.VK_ALT);
		r1.keyRelease(KeyEvent.VK_TAB);
		
		driver.close();
		
		driver.switchTo().parentFrame();
		
		Thread.sleep(1000);
		r1.keyPress(KeyEvent.VK_ALT);
		r1.keyPress(KeyEvent.VK_TAB);
		r1.keyRelease(KeyEvent.VK_ALT);
		r1.keyRelease(KeyEvent.VK_TAB);
		
		Set<String> windowHandles2 = driver.getWindowHandles();
		for (String windowHandle2 : windowHandles2) {
			driver.switchTo().window(windowHandle2);

			// You can check the title or URL of the window here
			String windowTitle2 = driver.getTitle();
			if (windowTitle2.equals("Add - Google Chrome"));
			
			System.out.println(driver.getTitle());
			
			driver.getTitle();
		
		driver.manage().window().maximize();
	
		Thread.sleep(2000);
		
		driver.findElement(By.xpath("//*[@id=\"txtSearchAll\"]")).sendKeys("hello");
	
	
	driver.findElement(By.xpath("/html/body/div[2]/div[2]/div/div/div/div/div/div/form/div/div[1]/div[1]/span[3]/label")).click();
	
	driver.findElement(By.xpath("//*[@id=\"idRadioExistingCustomer\"]")).click();
	
	
	
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
