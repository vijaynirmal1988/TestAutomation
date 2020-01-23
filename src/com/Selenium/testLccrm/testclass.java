package com.Selenium.testLccrm;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;

public class testclass {

	WebDriver wd;
	WebDriverWait wait;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFCell cell;
	MissingCellPolicy blank = Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

	@BeforeSuite
	public void beforeSuite() {

		System.setProperty("webdriver.chrome.driver", "E:\\DOWNLOADS\\chromedriver.exe");
		wd = new ChromeDriver();
		wait = new WebDriverWait(wd, 30);
		wd.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		wd.manage().window().maximize();
		wd.get("http://localhost:80/lccrm/lccrmUI/user/userLogin.php");
		wd.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	}

	@Test(priority = 0)
	public void LeadData() throws IOException, InterruptedException {
		wd.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

		File src = new File("C:\\Users\\hp\\eclipse workspace\\TestAutomation\\TestData.xlsx");
		FileInputStream finput = new FileInputStream(src);
		workbook = new XSSFWorkbook(finput);
		sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

		cell = sheet.getRow(1).getCell(0, blank);
		cell.setCellType(CellType.STRING);
		wd.findElement(By.id("txtemail")).sendKeys(cell.getStringCellValue());

		cell = sheet.getRow(1).getCell(1, blank);
		cell.setCellType(CellType.STRING);
		wd.findElement(By.id("txtpwd")).sendKeys(cell.getStringCellValue());
		wd.findElement(By.id("btnSave")).click();
		Thread.sleep(1000);
	}

	@Test(priority = 1)
	public void RegisterData() throws IOException, InterruptedException {

		File src = new File("E:\\DOWNLOADS\\TestData.xlsx");
		FileInputStream finput = new FileInputStream(src);
		workbook = new XSSFWorkbook(finput);
		sheet = workbook.getSheetAt(5);
		//Counting Rows at Excel Sheet 
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		//Thread.sleep(500);
		new Select(wd.findElement(By.name("cmLeadStatusTable_length"))).selectByValue("25");
		WebElement table = wd.findElement(By.xpath("//table[@id='cmLeadStatusTable']"));
		List<WebElement> leadStatusTable = table.findElements(By.tagName("tr"));
		
		for (int i = 1; i < rowCount + 1; i++) { // loop for no. of rows in excel sheet
			Row row = sheet.getRow(i);
			row.setRowNum(i);
			String nameInput = row.getCell(0,blank).getStringCellValue(); 
			String leadstatusInput = row.getCell(1,blank).getStringCellValue(); 
			
				for (int k = 1; k < leadStatusTable.size(); k++) { // loop for no of rows in webpage table
				Thread.sleep(500);	
			 String nameWeb =wd.findElement(By.xpath("//*[@id='cmLeadStatusTable']/tbody/tr["+k+"]/td[3]")).getText();
		     String leadStatusWeb =wd.findElement(By.xpath("//*[@id='cmLeadStatusTable']/tbody/tr["+k+"]/td[8]")).getText();

				if(nameInput.equals(nameWeb) && leadstatusInput.equals(leadStatusWeb)) {
					if(leadstatusInput.equals("Converted")) {
						
					 continue;
					 
				} else if(leadstatusInput.equals("Unassigned")) {
			 Thread.sleep(500);
		     wd.findElement(By.xpath("//*[@id='cmLeadStatusTable']/tbody/tr["+k+"]/td[1]")).click(); 
			 Thread.sleep(1000);
			// code for converting unassigned to inprogress
			 wd.findElement(By.id("btnAssign")).click();
			 Thread.sleep(1000);
			 new Select(wd.findElement(By.id("selAsnLeadStatus"))).selectByValue(row.getCell(2,blank).getStringCellValue());
			 Thread.sleep(1000);
			 new Select(wd.findElement(By.id("selAsnCounselorID"))).selectByValue(row.getCell(3,blank).getStringCellValue());
			 wd.findElement(By.id("txtAsnRemarks")).sendKeys(row.getCell(4,blank).getStringCellValue());
	         Thread.sleep(1000);
			 wd.findElement(By.xpath("//*[@id=\"viewAssignModal\"]/div/div/div/div[2]/div/div/div[1]/button")).click();
			     } else if(leadstatusInput.equals("Inprogress")) {
			 String courseInput = row.getCell(2,blank).getStringCellValue(); 
			 wd.findElement(By.xpath("//*[@id='cmLeadStatusTable']/tbody/tr["+k+"]/td[2]/a")).click(); 
			 Thread.sleep(2000);
					    	// code for converting inprogress to registered
			 wd.findElement(By.id("btnRegister")).click();
			 Thread.sleep(2000);
			 wd.findElement(By.id("btnAddStudent")).click();
			 Thread.sleep(2000);
			 table = wd.findElement(By.xpath("//*[@id=\"courseTable\"]"));
		     WebDriverWait waitForElement = new WebDriverWait(wd, 30 );
			 List<WebElement> courseTable = table.findElements(By.tagName("tr"));
				for(int l=1;l<= courseTable.size();l++) {
					String courseWeb =wd.findElement(By.xpath("//*[@id=\"courseTable\"]/tbody/tr["+l+"]/td[2]")).getText();
						     if(courseInput.equals(courseWeb)) {
						    	  wd.findElement(By.xpath("//*[@id=\"courseTable\"]/tbody/tr["+l+"]/td[1]")).click();
						    	  break;
						     }
						  
						     }
						    Thread.sleep(1000); 
						    wd.findElement(By.id("btnPay")).click();
						    Thread.sleep(1000); 

						    new Select(wd.findElement(By.id("selPaymentFor"))).selectByValue(row.getCell(3,blank).getStringCellValue());
						    Thread.sleep(500);
						    new Select(wd.findElement(By.id("selPreferredYear"))).selectByValue(row.getCell(4,blank).getStringCellValue());
						    Thread.sleep(500);
						    new Select(wd.findElement(By.id("selCourseSchedulePref"))).selectByValue(row.getCell(5,blank).getStringCellValue());
						    Thread.sleep(500);
						    new Select(wd.findElement(By.id("selExpectedCourseMonth"))).selectByValue(row.getCell(6,blank).getStringCellValue());
						    Thread.sleep(500);
						    new Select(wd.findElement(By.id("selModeOfPayment"))).selectByValue(row.getCell(7,blank).getStringCellValue());
						    Thread.sleep(500);
						    WebElement Date = wd.findElement(By.id("payDate"));
						    Date.sendKeys(Keys.chord(Keys.CONTROL,"a")); 
							Date.sendKeys(Keys.BACK_SPACE);
						    Date.sendKeys(row.getCell(8,blank).getStringCellValue());
						    Date.sendKeys(Keys.TAB); 
						    Thread.sleep(500);
						    wd.findElement(By.id("transactionPayRef")).sendKeys(row.getCell(9,blank).getStringCellValue());
						    Thread.sleep(500);
						    wd.findElement(By.id("feeAmount")).sendKeys(row.getCell(10,blank).getStringCellValue());
						    wd.findElement(By.id("btnSubmitPayment")).click();
						    Thread.sleep(500);
						    wd.findElement(By.id("btnRegisterCancel")).click();
						    Thread.sleep(4000); 
						  // move to dashboard page
							wd.findElement(By.xpath("//text()[contains(.,'Leads')]/ancestor::p[1]")).click(); 
							wd.findElement(By.xpath("//p[text()='Dashboard']")).click(); 
							Thread.sleep(2000);

							new Select(wd.findElement(By.name("cmLeadStatusTable_length"))).selectByValue("25");

			         	} else {
				            	continue;
								}
				 			} else {
				 				continue;
				 			}

						}		
				WebElement element = wd.findElement(By.tagName("body"));
				JavascriptExecutor js = (JavascriptExecutor)wd;
				js.executeScript("arguments[0].scrollIntoView();", element); 
					
		}
	}
	}
