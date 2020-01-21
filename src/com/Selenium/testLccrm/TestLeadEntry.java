	
package com.Selenium.testLccrm;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
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

public class TestLeadEntry {

	WebDriver wd;
	WebDriverWait wait;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFCell cell;
	MissingCellPolicy blank = Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

	private static String timestamp() {
		return new SimpleDateFormat("dd-MM-yyyy HH-mm-ss").format(new Date());
	}

	@BeforeSuite
	public void beforeSuite() {

		System.setProperty("webdriver.chrome.driver", "E:\\DOWNLOADS\\chromedriver.exe");
		wd = new ChromeDriver();
		wait = new WebDriverWait(wd, 30);
		wd.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		wd.manage().window().maximize();
		// Start the new url from
		wd.get("http://localhost:80/lccrm/lccrmUI/user/userLogin.php");
		// wd.get("http://stagingenvironmentlccrm.cloudaccess.host/lccrm/lccrmUI/user/userLogin.php");
		//wait = new WebDriverWait(wd, 30);
		wd.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	}
	
	@Test(priority = 0)
	public void LoginData() throws IOException, InterruptedException {
		wd.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);				

		// Import excel sheet.
		File src = new File("E:\\DOWNLOADS\\TestData.xlsx");
		// Load the file.
		FileInputStream finput = new FileInputStream(src);
		// Load he workbook.
		workbook = new XSSFWorkbook(finput);

		// Load the sheet in which data is stored.
		sheet = workbook.getSheetAt(0);
		
	   int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
     // System.out.println(rowCount);

		// for(int j=1 ; j < rowCount + 1 ; j++) {
		// Row row = sheet.getRow(j);
		// row.setRowNum(j);

		// Import data for Email.
		cell = sheet.getRow(1).getCell(0, blank);
		cell.setCellType(CellType.STRING);
		wd.findElement(By.id("txtemail")).sendKeys(cell.getStringCellValue());

		// Import data for password.
		cell = sheet.getRow(1).getCell(1, blank);
		cell.setCellType(CellType.STRING);
		wd.findElement(By.id("txtpwd")).sendKeys(cell.getStringCellValue());
		// Enter to Login
		wd.findElement(By.id("btnSave")).click();
		Thread.sleep(2000);
		// Alert alert = wd.switchTo().alert();
		// String alertMessage= wd.switchTo().alert().getText();
		// alert.accept();
		// }
	}
	/*	@Test(priority = 1)
		public void LeadData() throws IOException, InterruptedException {
			wd.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);				

			// Import excel sheet.
			File src = new File("E:\\DOWNLOADS\\TestData.xlsx");
			// Load the file.
			FileInputStream finput = new FileInputStream(src);
			// Load he workbook.
			workbook = new XSSFWorkbook(finput);

			// Load the sheet in which data is stored.
			sheet = workbook.getSheetAt(1);
			
		   int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();

		  for (int i = 1; i < rowCount+1; i++) { 
			  Row row = sheet.getRow(i);
	              row.setRowNum(i); 
		 for (int k = 0; k < row.getLastCellNum();k++) {
		 		 Thread.sleep(2000); 

		  //select to the center to create a new leads 
		 new Select(wd.findElement(By.id("selCenter"))).selectByValue(row.getCell(k++,blank).getStringCellValue()); 
	     if(k < sheet.getLastRowNum()) { 
	    	 wd.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);				
	     
	// wd.findElement(By.xpath("//a[@class='nav-link collapsed']//p[contains(text(),'Leads')]")).click(); 
		  wd.findElement(By.xpath("//p[contains(text(),'Add New Lead')]")).click();
	     }
 
		  wd.findElement(By.id("txtFirstName")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		  wd.findElement(By.id("txtMiddleName")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		  wd.findElement(By.id("txtLastName")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		   WebElement Date = wd.findElement(By.xpath("//input[@type='text'][@name='Dob']"));
		  Date.sendKeys(Keys.chord(Keys.CONTROL,"a")); 
		  Date.sendKeys(Keys.BACK_SPACE);
		  Date.sendKeys(String.valueOf(row.getCell(k++,blank).getStringCellValue()));
		  Date.sendKeys(Keys.TAB); 
		  WebElement gender1 = wd.findElement(By.id("rdMale"));
		  WebElement gender2 = wd.findElement(By.id("rdFemale")); 
		  WebElement gender3 = wd.findElement(By.id("rdOthers"));
		 if(row.getCell(k,blank).getStringCellValue().equals("Male")) {
		   ((JavascriptExecutor) wd).executeScript("arguments[0].checked = true;", gender1);
		 } else if(row.getCell(k,blank).getStringCellValue().equals("Female")) {
		 ((JavascriptExecutor) wd).executeScript("arguments[0].checked = true;",gender2);
		 } else if(row.getCell(k,blank).getStringCellValue().equals("Transgender")) {
		   ((JavascriptExecutor) wd).executeScript("arguments[0].checked = true;",gender3); 
		   } k++;
		   Thread.sleep(100);
		  wd.findElement(By.name("parentguardianname")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		  wd.findElement(By.name("address")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		  wd.findElement(By.name("city")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		  wd.findElement(By.name("state")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		  wd.findElement(By.name("pincode")).sendKeys(row.getCell(k++,blank).getStringCellValue());
     	  wd.findElement(By.name("mobilenumber")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		  wd.findElement(By.name("landline")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		  wd.findElement(By.name("email")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		  wd.findElement(By.name("pgMobileNo")).sendKeys(row.getCell(k++,blank).getStringCellValue());
	      wd.findElement(By.id("txtParentsEmail")).sendKeys(row.getCell(k++,blank).getStringCellValue()); 
	      Thread.sleep(100); 
	      WebElement Prospect1 =wd.findElement(By.id("rdGreen")); 
	      WebElement Prospect2 = wd.findElement(By.id("rdYellow")); 
	      WebElement Prospect3 = wd.findElement(By.id("rdRed"));
		 if(row.getCell(k,blank).getStringCellValue().equals("G")) {
		     ((JavascriptExecutor) wd).executeScript("arguments[0].checked = true;",Prospect1); 
		 } else if(row.getCell(k,blank).getStringCellValue().equals("Y"))
		   { ((JavascriptExecutor) wd).executeScript("arguments[0].checked = true;",Prospect2);
		 } else if(row.getCell(k,blank).getStringCellValue().equals("R"))
		    { ((JavascriptExecutor) wd).executeScript("arguments[0].checked = true;",Prospect3); 
		 } k++; 
		 new Select(wd.findElement(By.id("selCurrentEducationStatus"))).selectByValue(row.getCell(k++,blank).getStringCellValue());
		 wd.findElement(By.id("txtInstitutionName")).sendKeys(row.getCell(k++,blank).getStringCellValue()); 
		 new Select(wd.findElement(By.id("selAnnualIncome"))).selectByValue(row.getCell(k++,blank).getStringCellValue());
		 new Select(wd.findElement(By.id("selStreamofEducation"))).selectByValue(row.getCell(k++,blank).getStringCellValue()); 
		 new Select(wd.findElement(By.id("selTypeofIndustry"))).selectByValue(row.getCell(k++,blank).getStringCellValue());
		 wd.findElement(By.name("WorkExperience")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		 wd.findElement(By.id("txtCompanyName")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		 //Course selection 
		  WebElement checkcourse1 = wd.findElement(By.xpath("//text()[contains(.,'GRE')]/ancestor::label[1]"));
		  WebElement checkcourse2 = wd.findElement(By.xpath("//text()[contains(.,'GMAT')]/ancestor::label[1]"));
		  WebElement checkcourse3 = wd.findElement(By.xpath("//text()[contains(.,'SAT')]/ancestor::label[1]")); 
		  WebElement checkcourse4 = wd.findElement(By.xpath("//text()[contains(.,'ACT')]/ancestor::label[1]"));
		  WebElement checkcourse5 = wd.findElement(By.xpath("//text()[contains(.,'TOEFL')]/ancestor::label[1]"));
		  WebElement checkcourse6 = wd.findElement(By.xpath("//text()[contains(.,'IELTS')]/ancestor::label[1]"));
		  WebElement checkcourse7 = wd.findElement(By.xpath("//text()[contains(.,'Admission Counseling')]/ancestor::label[1]"));
		  WebElement checkcourse8 = wd.findElement(By.xpath("//text()[contains(.,'Others')]/ancestor::label[1]"));
		 
		 String courses = row.getCell(k,blank).getStringCellValue(); 
		 String[]coursesSplit = courses.split(",");
		 
		 //looping through the course names separated by , 
		 for (int x = 0; x < coursesSplit.length; x++) {
		 
		 if(coursesSplit[x].equals("GRE")) { 
		     ((JavascriptExecutor)wd).executeScript("arguments[0].click();", checkcourse1); 
		 } else if(coursesSplit[x].equals("GMAT")) { 
			 ((JavascriptExecutor) wd).executeScript("arguments[0].click();", checkcourse2); 
		 } else if(coursesSplit[x].equals("SAT")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", checkcourse3);
		 } else if(coursesSplit[x].equals("ACT")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", checkcourse4); 
		 } else if(coursesSplit[x].equals("TOFEL")) {
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", checkcourse5); 
		 } else if(coursesSplit[x].equals("IELTS")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", checkcourse6);
		 } else if(coursesSplit[x].equals("Admission Counseling")) {
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", checkcourse7); 
		 } else if(coursesSplit[x].equals("Others")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", checkcourse8);
		     wd.findElement(By.id("txtCourseOthers")).sendKeys(coursesSplit[++x]); 
		  }
		 
		  } k++;
		   //country selection 
		   WebElement chkcountry1 = wd.findElement(By.xpath("/html/body/div[2]/div/div/div/div/div[2]/form/div[13]/div/div[1]/label/input")); 
		   WebElement chkcountry2 = wd.findElement(By.xpath("//text()[contains(.,'Canada')]/ancestor::label[1]"));
		   WebElement chkcountry3 = wd.findElement(By.xpath("/html/body/div[2]/div/div/div/div/div[2]/form/div[13]/div/div[3]/label/input")); 
		   WebElement chkcountry4 = wd.findElement(By.xpath("//text()[contains(.,'Australia')]/ancestor::label[1]")); 
		   WebElement chkcountry5 = wd.findElement(By.xpath("//text()[contains(.,'Singapore')]/ancestor::label[1]")); 
		   WebElement chkcountry6 = wd.findElement(By.xpath("//text()[contains(.,'Germany')]/ancestor::label[1]")); 
		   WebElement chkcountry7 = wd.findElement(By.xpath("/html[1]/body[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[14]/div[1]/div[1]/label[2]/span[1]/span[1]")); 
		   Thread.sleep(100); 
	       String country = row.getCell(k,blank).getStringCellValue(); 
	       String[] countrySplit =country.split(",");
		 
		 //looping through the Country names separated by , 
	       for (int y = 0; y < countrySplit.length; y++) {
		    if(countrySplit[y].equals("US")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", chkcountry1);
		 } else if(countrySplit[y].equals("Canada")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", chkcountry2); 
			} else if(countrySplit[y].equals("UK")) { 
			((JavascriptExecutor)wd).executeScript("arguments[0].click();", chkcountry3);
		 } else if(countrySplit[y].equals("Australia")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", chkcountry4); 
		 } else if(countrySplit[y].equals("Singapore")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", chkcountry5); 
		 } else if(countrySplit[y].equals("Germany")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", chkcountry6);
		 } else if(countrySplit[y].equals("Others")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", chkcountry7);
		 wd.findElement(By.id("txtCountryOthers")).sendKeys(countrySplit[++y]);
		 }
		 } k++;
		 
		 WebElement KnowAbout1 = wd.findElement(By.xpath("//text()[contains(.,'Flyers')]/ancestor::label[1]")); 
		 WebElement KnowAbout2 = wd.findElement(By.xpath("//text()[contains(.,'Paper Ad')]/ancestor::label[1]")); 
		 WebElement KnowAbout3 =wd.findElement(By.xpath("//text()[contains(.,'Internet')]/ancestor::label[1]")); 
		 WebElement KnowAbout4 = wd.findElement(By.xpath("//text()[contains(.,'Social Net')]/ancestor::label[1]")); 
		 WebElement KnowAbout5 = wd.findElement(By.xpath("//text()[contains(.,'College Seminar')]/ancestor::label[1]"));
		 WebElement KnowAbout6 = wd.findElement(By.xpath("//text()[contains(.,'Friends')]/ancestor::label[1]")); 
		 WebElement KnowAbout7 = wd.findElement(By.xpath("//text()[contains(.,'Hoarding')]/ancestor::label[1]")); 
		 WebElement KnowAbout8 = wd.findElement(By.xpath("//text()[contains(.,'Poster/Banner')]/ancestor::label[1]")); 
		 WebElement KnowAbout9 = wd.findElement(By.xpath("//text()[contains(.,'Email from Magoosh')]/ancestor::label[1]"));
		 WebElement KnowAbout10 = wd.findElement(By.xpath("//text()[contains(.,'News Paper')]/ancestor::label[1]")); 
		 WebElement KnowAbout11 = wd.findElement(By. xpath("/html[1]/body[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/div[17]/div[1]/div[1]/label[2]" )); 
		 Thread.sleep(100); 
		 String knowabt = row.getCell(k,blank).getStringCellValue(); 
		 String[] knowabtSplit = knowabt.split(",");
		 
		 //looping through the Country names separated by , 
		 for (int z = 0; z <knowabtSplit.length; z++) { 
			 if(knowabtSplit[z].equals("Flyers")) {
		     ((JavascriptExecutor) wd).executeScript("arguments[0].click();", KnowAbout1);
		 } else if(knowabtSplit[z].equals("Paper Ad")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", KnowAbout2); 
		 } else if(knowabtSplit[z].equals("Internet")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", KnowAbout3); 
		 } else if(knowabtSplit[z].equals("Social Net")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", KnowAbout4); 
		 } else if(knowabtSplit[z].equals("College Seminar")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", KnowAbout5); 
		 } else if(knowabtSplit[z].equals("Friends")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", KnowAbout6); 
		 } else if(knowabtSplit[z].equals("Hoarding")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", KnowAbout7); 
		 } else if(knowabtSplit[z].equals("Poster/Banner")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", KnowAbout8); 
		 } else if(knowabtSplit[z].equals("Email from Magoosh")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", KnowAbout9); 
		 } else if(knowabtSplit[z].equals("News Paper")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", KnowAbout10); 
		 } else if(knowabtSplit[z].equals("Others")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", KnowAbout11);
		   wd.findElement(By.id("txtKnowAboutOthers")).sendKeys(knowabtSplit[++z]);
		 } 
			 } k++;
		 
		 wd.findElement(By.id("txtAddFriendName1")).sendKeys(row.getCell(k++,blank).getStringCellValue());
	     wd.findElement(By.id("txtAddFriendMobile1")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		 wd.findElement(By.id("txtAddFriendEmail1")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		 wd.findElement(By.id("txtAddFriendName2")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		 wd.findElement(By.id("txtAddFriendMobile2")).sendKeys(row.getCell(k++,blank).getStringCellValue());
		 wd.findElement(By.id("txtAddFriendEmail2")).sendKeys(row.getCell(k++,blank).getStringCellValue()); 
		 Thread.sleep(100); 
		 new Select(wd.findElement(By.id("selCourseEstimatedStart"))).selectByValue(row.getCell(k++,blank).getStringCellValue()); 
		 new Select(wd.findElement(By.id("selClassSchedulePref"))).selectByValue(row.getCell(k++,blank).getStringCellValue()); 
		 new Select(wd.findElement(By.id("selHasMagooshAccount"))).selectByValue(row.getCell(k++,blank).getStringCellValue());
		 
		 // Select Source of Lead 
		 WebElement SOL1 = wd.findElement(By.xpath("//div[@class='col-sm-12']//div[1]//label[1]")); 
		 WebElement SOL2 = wd.findElement(By.xpath("//div[@class='col-sm-12']//div[2]//label[1]")); 
		 WebElement SOL3 = wd.findElement(By.xpath("//div[@class='col-sm-12']//div[3]//label[1]")); 
		 WebElement SOL4 = wd.findElement(By.xpath("//div[@class='col-sm-12']//div[4]//label[1]")); 
		 WebElement SOL5 = wd.findElement(By.xpath("//div[@class='col-sm-12']//div[5]//label[1]")); 
		 WebElement SOL6 = wd.findElement(By.xpath("//div[@class='col-sm-12']//div[6]//label[1]")); 
		 WebElement SOL7 = wd.findElement(By.xpath("//div[@class='col-sm-12']//div[7]//label[1]")); 
		 WebElement SOL8 = wd.findElement(By.xpath("//div[@class='col-sm-12']//div[8]//label[1]")); 
		 WebElement SOL9 = wd.findElement(By.xpath("//div[@class='col-sm-12']//div[9]//label[1]")); 
		 WebElement SOL10 = wd.findElement(By.xpath("//div[@class='col-sm-12']//div[10]//label[1]")); 
		 WebElement SOL11 = wd.findElement(By.xpath("//div[@class='col-sm-12']//div[11]//label[1]")); 
		 Thread.sleep(100);
		 
		 String SOL = row.getCell(k,blank).getStringCellValue(); 
		 String[] SOLSplit = SOL.split(",");
		 
		 //looping through the course names separated by , 
		 for (int w = 0; w < SOLSplit.length; w++) {
		 
		 if(SOLSplit[w].equals("Telephone")) { 
		 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", SOL1); 
		 } else if(SOLSplit[w].equals("Whatsapp")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", SOL2); 
		 } else if(SOLSplit[w].equals("Direct Walkin")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", SOL3); 
		 } else  if(SOLSplit[w].equals("College Seminar")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", SOL4); 
			 } else if(SOLSplit[w].equals("Student Reference")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", SOL5); 
		 } else if(SOLSplit[w].equals("Activities")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", SOL6); 
		 } else if(SOLSplit[w].equals("Sulekha")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", SOL7);
		 } else if(SOLSplit[w].equals("Thinkvidya")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", SOL8); 
	     } else if(SOLSplit[w].equals("Just Dial")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", SOL9);
	     } else if(SOLSplit[w].equals("Urban Pro")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", SOL10);
	     } else if(SOLSplit[w].equals("Others")) { 
			 ((JavascriptExecutor)wd).executeScript("arguments[0].click();", SOL11);
		     wd.findElement(By.id("txtSOLOthers")).sendKeys(SOLSplit[++w]); 
		 } 
		 }
		 wd.findElement(By.id("btnSave")).click(); 
		 Thread.sleep(200); 
		 File scrFile =((TakesScreenshot)wd).getScreenshotAs(OutputType.FILE);
		 FileUtils.copyFile(scrFile, 
		 new File("E:/DOWNLOADS/Screenshot/Newlead_ "+timestamp()+".png"));
		 SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy HH-mm-ss");
		 Thread.sleep(200); 
		 } 
		// wd.findElement(By.xpath("//button[text()='OK']")).click(); 
		 }
		 
		 
	}*/
	@Test(priority = 2)
	public void RegisterData() throws IOException, InterruptedException {
	 // File src = new File("E:\\DOWNLOADS\\TestData.xlsx");
	  File src = new File("C:\\Users\\hp\\eclipse workspace\\TestAutomation\\TestData.xlsx");
	  FileInputStream finput = new FileInputStream(src);
	  workbook = new XSSFWorkbook(finput);
	  sheet = workbook.getSheetAt(5);
	 // sheet = workbook.getSheetAt(6);
	//Counting Rows at Excel Sheet 
	  int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
	 Thread.sleep(500);
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
		
/*	@Test(priority = 3)
	public void BatchDataRead() throws IOException, InterruptedException {
		
		// Import excel sheet.
		File src = new File("E:\\DOWNLOADS\\TestData.xlsx");
		// Load the file.
		FileInputStream finput = new FileInputStream(src);
		// Load he workbook.
		workbook = new XSSFWorkbook(finput);

		// Load the sheet in which data is stored.
		sheet = workbook.getSheetAt(4);
			
		 int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
	    
		for (int i = 1; i < rowCount + 1; i++) {
			Row row = sheet.getRow(i);
		    row.setRowNum(i);
		for (int k = 0; k < row.getLastCellNum(); k++) {
		wd.findElement(By.xpath("//text()[contains(.,'Batch Details')]/ancestor::p[1]")).click();
		// Enter to Batch Group Master
		wd.findElement(By.xpath("//p[text()='Batch Group']")).click();
		Thread.sleep(100);
		// add Batch Group
		
		wd.findElement(By.id("btnBGroupAdd")).click();
		new Select(wd.findElement(By.id("selBatchGroupYear"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("selBatchGroupQuater"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("btnBGroupSave")).click();
					
		// Enter to batch Master
		Thread.sleep(2000);
		wd.findElement(By.xpath("//p[text()='Batch']")).click();
					
		// Enter to add batch details
		wd.findElement(By.id("btnAddBatch")).click();
		new Select(wd.findElement(By.id("selBatchGroupCode"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("selCourseList"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("selBatchType"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("txtBatchSizeMin")).sendKeys(Keys.chord(Keys.CONTROL, "a"), row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("txtBatchSizeMax")).sendKeys(Keys.chord(Keys.CONTROL, "a"), row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("selBatchSession"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		WebElement StartDate = wd.findElement(By.id("txtBatchStartDate"));
		StartDate.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartDate.sendKeys(Keys.BACK_SPACE);
	    StartDate.sendKeys(row.getCell(k++, blank).getStringCellValue());StartDate.sendKeys(Keys.TAB);
		WebElement EndDate = wd.findElement(By.id("txtBatchEndDate"));
		EndDate.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndDate.sendKeys(Keys.BACK_SPACE);
		EndDate.sendKeys(row.getCell(k++, blank).getStringCellValue());EndDate.sendKeys(Keys.TAB);
		wd.findElement(By.id("btnAdd")).click();
		Thread.sleep(1000);
		
		// Enter to add Course Details
		new Select(wd.findElement(By.id("facultyVerbal"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("noClassVerbal")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("noTestVerbal")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("facultyQuants"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("noClassQuants")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("noTestQuants")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("btnCourseAdd")).click();
		Thread.sleep(500);
	    
		// Enter to add Class schedule
		WebElement ClassDate1 = wd.findElement(By.id("dtclass1"));
		ClassDate1.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject1"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty1"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic1")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent1")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime1 = wd.findElement(By.id("classStart1"));
		StartTime1.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime1.sendKeys(Keys.BACK_SPACE);
		StartTime1.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime1.sendKeys(Keys.TAB);
		WebElement EndTime1 = wd.findElement(By.id("classEnd1"));
		EndTime1.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime1.sendKeys(Keys.BACK_SPACE);
		EndTime1.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime1.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row1\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		WebElement ClassDate2 = wd.findElement(By.id("dtclass2"));
		ClassDate2.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject2"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty2"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic2")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent2")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime2 = wd.findElement(By.id("classStart2"));
		StartTime2.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime2.sendKeys(Keys.BACK_SPACE);
		StartTime2.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime2.sendKeys(Keys.TAB);
		WebElement EndTime2 = wd.findElement(By.id("classEnd2"));
		EndTime2.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime2.sendKeys(Keys.BACK_SPACE);
		EndTime2.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime2.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row2\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		WebElement ClassDate3 = wd.findElement(By.id("dtclass3"));
		ClassDate3.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject3"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty3"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic3")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent3")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime3 = wd.findElement(By.id("classStart3"));
		StartTime3.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime3.sendKeys(Keys.BACK_SPACE);
		StartTime3.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime3.sendKeys(Keys.TAB);
		WebElement EndTime3 = wd.findElement(By.id("classEnd3"));
		EndTime3.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime3.sendKeys(Keys.BACK_SPACE);
		EndTime3.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime3.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row3\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);	
		
		WebElement ClassDate4 = wd.findElement(By.id("dtclass4"));
		ClassDate4.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject4"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty4"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic4")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent4")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime4 = wd.findElement(By.id("classStart4"));
		StartTime4.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime4.sendKeys(Keys.BACK_SPACE);
		StartTime4.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime4.sendKeys(Keys.TAB);
		WebElement EndTime4 = wd.findElement(By.id("classEnd4"));
		EndTime4.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime4.sendKeys(Keys.BACK_SPACE);
		EndTime4.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime4.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row4\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		WebElement ClassDate5 = wd.findElement(By.id("dtclass5"));
		ClassDate5.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject5"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty5"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic5")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent5")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime5 = wd.findElement(By.id("classStart5"));
		StartTime5.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime5.sendKeys(Keys.BACK_SPACE);
		StartTime5.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime5.sendKeys(Keys.TAB);
		WebElement EndTime5 = wd.findElement(By.id("classEnd5"));
		EndTime5.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime5.sendKeys(Keys.BACK_SPACE);
		EndTime5.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime5.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row5\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		WebElement ClassDate6 = wd.findElement(By.id("dtclass6"));
		ClassDate6.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject6"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty6"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic6")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent6")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime6 = wd.findElement(By.id("classStart6"));
		StartTime6.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime6.sendKeys(Keys.BACK_SPACE);
		StartTime6.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime6.sendKeys(Keys.TAB);
		WebElement EndTime6 = wd.findElement(By.id("classEnd6"));
		EndTime6.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime6.sendKeys(Keys.BACK_SPACE);
		EndTime6.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime6.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row6\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		WebElement ClassDate7 = wd.findElement(By.id("dtclass7"));
		ClassDate7.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject7"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty7"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic7")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent7")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime7 = wd.findElement(By.id("classStart7"));
		StartTime7.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime7.sendKeys(Keys.BACK_SPACE);
		StartTime7.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime7.sendKeys(Keys.TAB);
		WebElement EndTime7 = wd.findElement(By.id("classEnd7"));
		EndTime7.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime7.sendKeys(Keys.BACK_SPACE);
		EndTime7.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime7.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row7\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
	    try {
		
		WebElement ClassDate8 = wd.findElement(By.id("dtclass8"));
		ClassDate8.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject8"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty8"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic8")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent8")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime8 = wd.findElement(By.id("classStart8"));
		StartTime8.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime8.sendKeys(Keys.BACK_SPACE);
		StartTime8.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime8.sendKeys(Keys.TAB);
		WebElement EndTime8 = wd.findElement(By.id("classEnd8"));
		EndTime8.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime8.sendKeys(Keys.BACK_SPACE);
		EndTime8.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime8.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row8\"]/td[9]/button[1]\r\n")).click();
	      } catch (org.openqa.selenium.ElementNotInteractableException e) {
		         e.printStackTrace();
	     }
		
	    Thread.sleep(2000);
	    try {
		WebElement ClassDate9 = wd.findElement(By.id("dtclass9"));
		ClassDate9.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject9"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty9"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic9")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent9")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime9 = wd.findElement(By.id("classStart9"));
		StartTime9.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime9.sendKeys(Keys.BACK_SPACE);
		StartTime9.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime9.sendKeys(Keys.TAB);
		WebElement EndTime9 = wd.findElement(By.id("classEnd9"));
		EndTime9.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime9.sendKeys(Keys.BACK_SPACE);
		EndTime9.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime9.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row9\"]/td[9]/button[1]\r\n")).click();
	      } catch (org.openqa.selenium.ElementNotInteractableException e) {
			
			e.printStackTrace();
		}
		Thread.sleep(2000);
		
		try {
		WebElement ClassDate10 = wd.findElement(By.id("dtclass10"));
		ClassDate10.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject10"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty10"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic10")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent10")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime10 = wd.findElement(By.id("classStart10"));
		StartTime10.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime10.sendKeys(Keys.BACK_SPACE);
		StartTime10.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime10.sendKeys(Keys.TAB);
		WebElement EndTime10 = wd.findElement(By.id("classEnd10"));
		EndTime10.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime10.sendKeys(Keys.BACK_SPACE);
		EndTime10.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime10.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row10\"]/td[9]/button[1]\r\n")).click();
		} catch (Exception e) {
			
			e.printStackTrace();
		}
		Thread.sleep(2000);
		
		WebElement ClassDate11 = wd.findElement(By.id("dtclass11"));
		ClassDate11.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject11"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty11"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic11")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent11")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime11 = wd.findElement(By.id("classStart11"));
		StartTime11.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime11.sendKeys(Keys.BACK_SPACE);
		StartTime11.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime11.sendKeys(Keys.TAB);
		WebElement EndTime11 = wd.findElement(By.id("classEnd11"));
		EndTime11.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime11.sendKeys(Keys.BACK_SPACE);
		EndTime11.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime11.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row11\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		WebElement ClassDate12 = wd.findElement(By.id("dtclass12"));
		ClassDate12.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject12"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty12"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic12")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent12")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime12 = wd.findElement(By.id("classStart12"));
		StartTime12.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime12.sendKeys(Keys.BACK_SPACE);
		StartTime12.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime12.sendKeys(Keys.TAB);
		WebElement EndTime12 = wd.findElement(By.id("classEnd12"));
		EndTime12.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime12.sendKeys(Keys.BACK_SPACE);
		EndTime12.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime12.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row12\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		WebElement ClassDate13 = wd.findElement(By.id("dtclass13"));
		ClassDate13.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject13"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty13"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic13")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent13")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime13 = wd.findElement(By.id("classStart13"));
		StartTime13.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime13.sendKeys(Keys.BACK_SPACE);
		StartTime13.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime13.sendKeys(Keys.TAB);
		WebElement EndTime13 = wd.findElement(By.id("classEnd13"));
		EndTime13.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime13.sendKeys(Keys.BACK_SPACE);
		EndTime13.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime13.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row13\"]/td[9]/button[1]\r\n")).click();
	    Thread.sleep(2000);
	    
		WebElement ClassDate14 = wd.findElement(By.id("dtclass14"));
		ClassDate14.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject14"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty14"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic14")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent14")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime14 = wd.findElement(By.id("classStart14"));
		StartTime14.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime14.sendKeys(Keys.BACK_SPACE);
		StartTime14.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime14.sendKeys(Keys.TAB);
		WebElement EndTime14 = wd.findElement(By.id("classEnd14"));
		EndTime14.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime14.sendKeys(Keys.BACK_SPACE);
		EndTime14.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime14.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row14\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		WebElement ClassDate15 = wd.findElement(By.id("dtclass15"));
		ClassDate15.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject15"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty15"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic15")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent15")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime15 = wd.findElement(By.id("classStart15"));
		StartTime15.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime15.sendKeys(Keys.BACK_SPACE);
		StartTime15.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime15.sendKeys(Keys.TAB);
		WebElement EndTime15 = wd.findElement(By.id("classEnd15"));
		EndTime15.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime15.sendKeys(Keys.BACK_SPACE);
		EndTime15.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime15.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row15\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		WebElement ClassDate16 = wd.findElement(By.id("dtclass16"));
		ClassDate16.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject16"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty16"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic16")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent16")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime16 = wd.findElement(By.id("classStart16"));
		StartTime16.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime16.sendKeys(Keys.BACK_SPACE);
		StartTime16.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime16.sendKeys(Keys.TAB);
		WebElement EndTime16 = wd.findElement(By.id("classEnd16"));
		EndTime16.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime16.sendKeys(Keys.BACK_SPACE);
		EndTime16.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime16.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row16\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		WebElement ClassDate17 = wd.findElement(By.id("dtclass17"));
		ClassDate17.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject17"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty17"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic17")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent17")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime17 = wd.findElement(By.id("classStart17"));
		StartTime17.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime17.sendKeys(Keys.BACK_SPACE);
		StartTime17.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime17.sendKeys(Keys.TAB);
		WebElement EndTime17 = wd.findElement(By.id("classEnd17"));
		EndTime17.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime17.sendKeys(Keys.BACK_SPACE);
		EndTime17.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime17.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row17\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		WebElement ClassDate18 = wd.findElement(By.id("dtclass18"));
		ClassDate18.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject18"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty18"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic18")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent18")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime18 = wd.findElement(By.id("classStart18"));
		StartTime18.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime18.sendKeys(Keys.BACK_SPACE);
		StartTime18.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime18.sendKeys(Keys.TAB);
		WebElement EndTime18 = wd.findElement(By.id("classEnd18"));
		EndTime18.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime18.sendKeys(Keys.BACK_SPACE);
		EndTime18.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime18.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row18\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		WebElement ClassDate19 = wd.findElement(By.id("dtclass19"));
		ClassDate19.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject19"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty19"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic19")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent19")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime19 = wd.findElement(By.id("classStart19"));
		StartTime19.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime19.sendKeys(Keys.BACK_SPACE);
		StartTime19.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime19.sendKeys(Keys.TAB);
		WebElement EndTime19 = wd.findElement(By.id("classEnd19"));
		EndTime19.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime19.sendKeys(Keys.BACK_SPACE);
		EndTime19.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime19.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row19\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		WebElement ClassDate20 = wd.findElement(By.id("dtclass20"));
		ClassDate20.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classSubject20"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("classfaculty20"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classTopic20")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("classContent20")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement StartTime20 = wd.findElement(By.id("classStart20"));
		StartTime20.sendKeys(Keys.chord(Keys.CONTROL, "a"));StartTime20.sendKeys(Keys.BACK_SPACE);
		StartTime20.sendKeys(row.getCell(k++, blank).getStringCellValue());StartTime20.sendKeys(Keys.TAB);
		WebElement EndTime20 = wd.findElement(By.id("classEnd20"));
		EndTime20.sendKeys(Keys.chord(Keys.CONTROL, "a"));EndTime20.sendKeys(Keys.BACK_SPACE);
		EndTime20.sendKeys(row.getCell(k++, blank).getStringCellValue());EndTime20.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row20\"]/td[9]/button[1]\r\n")).click();
		Thread.sleep(2000);
		
		// Enter to add Test Schedule
		WebElement TestDate1 = wd.findElement(By.id("dttest1"));
		TestDate1.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject1"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty1"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic1")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent1")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime1 = wd.findElement(By.id("testStart1"));
		TestStartTime1.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime1.sendKeys(Keys.BACK_SPACE);
		TestStartTime1.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime1.sendKeys(Keys.TAB);
		WebElement TestEndTime1 = wd.findElement(By.id("testEnd1"));
		TestEndTime1.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime1.sendKeys(Keys.BACK_SPACE);
		TestEndTime1.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime1.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row1\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate2 = wd.findElement(By.id("dttest2"));
		TestDate2.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject2"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty2"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic2")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent2")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime2 = wd.findElement(By.id("testStart2"));
		TestStartTime2.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime2.sendKeys(Keys.BACK_SPACE);
		TestStartTime2.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime2.sendKeys(Keys.TAB);
		WebElement TestEndTime2 = wd.findElement(By.id("testEnd2"));
		TestEndTime2.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime2.sendKeys(Keys.BACK_SPACE);
		TestEndTime2.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime2.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row2\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate3 = wd.findElement(By.id("dttest3"));
		TestDate3.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject3"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty3"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic3")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent3")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime3 = wd.findElement(By.id("testStart3"));
		TestStartTime3.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime3.sendKeys(Keys.BACK_SPACE);
		TestStartTime3.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime3.sendKeys(Keys.TAB);
		WebElement TestEndTime3 = wd.findElement(By.id("testEnd3"));
		TestEndTime3.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime3.sendKeys(Keys.BACK_SPACE);
		TestEndTime3.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime3.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row3\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate4 = wd.findElement(By.id("dttest4"));
		TestDate4.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject4"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty4"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic4")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent4")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime4 = wd.findElement(By.id("testStart4"));
		TestStartTime4.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime4.sendKeys(Keys.BACK_SPACE);
		TestStartTime4.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime4.sendKeys(Keys.TAB);
		WebElement TestEndTime4 = wd.findElement(By.id("testEnd4"));
		TestEndTime4.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime4.sendKeys(Keys.BACK_SPACE);
		TestEndTime4.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime4.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row4\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate5 = wd.findElement(By.id("dttest5"));
		TestDate5.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject5"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty5"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic5")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent5")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime5 = wd.findElement(By.id("testStart5"));
		TestStartTime5.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime5.sendKeys(Keys.BACK_SPACE);
		TestStartTime5.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime5.sendKeys(Keys.TAB);
		WebElement TestEndTime5 = wd.findElement(By.id("testEnd5"));
		TestEndTime5.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime5.sendKeys(Keys.BACK_SPACE);
		TestEndTime5.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime5.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row5\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate6 = wd.findElement(By.id("dttest6"));
		TestDate6.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject6"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty6"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic6")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent6")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime6 = wd.findElement(By.id("testStart6"));
		TestStartTime6.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime6.sendKeys(Keys.BACK_SPACE);
		TestStartTime6.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime6.sendKeys(Keys.TAB);
		WebElement TestEndTime6 = wd.findElement(By.id("testEnd6"));
		TestEndTime6.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime6.sendKeys(Keys.BACK_SPACE);
		TestEndTime6.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime6.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row6\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate7 = wd.findElement(By.id("dttest7"));
		TestDate7.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject7"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty7"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic7")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent7")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime7 = wd.findElement(By.id("testStart7"));
		TestStartTime7.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime7.sendKeys(Keys.BACK_SPACE);
		TestStartTime7.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime7.sendKeys(Keys.TAB);
		WebElement TestEndTime7 = wd.findElement(By.id("testEnd7"));
		TestEndTime7.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime7.sendKeys(Keys.BACK_SPACE);
		TestEndTime7.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime7.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row7\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate8 = wd.findElement(By.id("dttest8"));
		TestDate8.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject8"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty8"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic8")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent8")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime8 = wd.findElement(By.id("testStart8"));
		TestStartTime8.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime8.sendKeys(Keys.BACK_SPACE);
		TestStartTime8.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime8.sendKeys(Keys.TAB);
		WebElement TestEndTime8 = wd.findElement(By.id("testEnd8"));
		TestEndTime8.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime8.sendKeys(Keys.BACK_SPACE);
		TestEndTime8.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime8.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row8\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate9 = wd.findElement(By.id("dttest9"));
		TestDate9.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject9"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty9"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic9")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent9")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime9 = wd.findElement(By.id("testStart9"));
		TestStartTime9.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime9.sendKeys(Keys.BACK_SPACE);
		TestStartTime9.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime9.sendKeys(Keys.TAB);
		WebElement TestEndTime9 = wd.findElement(By.id("testEnd9"));
		TestEndTime9.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime9.sendKeys(Keys.BACK_SPACE);
		TestEndTime9.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime9.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row9\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate10 = wd.findElement(By.id("dttest10"));
		TestDate10.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject10"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty10"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic10")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent10")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime10 = wd.findElement(By.id("testStart10"));
		TestStartTime10.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime10.sendKeys(Keys.BACK_SPACE);
		TestStartTime10.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime10.sendKeys(Keys.TAB);
		WebElement TestEndTime10 = wd.findElement(By.id("testEnd10"));
		TestEndTime10.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime10.sendKeys(Keys.BACK_SPACE);
		TestEndTime10.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime10.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row10\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate11 = wd.findElement(By.id("dttest11"));
		TestDate11.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject11"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty11"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic11")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent11")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime11 = wd.findElement(By.id("testStart11"));
		TestStartTime11.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime11.sendKeys(Keys.BACK_SPACE);
		TestStartTime11.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime11.sendKeys(Keys.TAB);
		WebElement TestEndTime11 = wd.findElement(By.id("testEnd11"));
		TestEndTime11.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime11.sendKeys(Keys.BACK_SPACE);
		TestEndTime11.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime11.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row11\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate12 = wd.findElement(By.id("dttest12"));
		TestDate12.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject12"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty12"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic12")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent12")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime12 = wd.findElement(By.id("testStart12"));
		TestStartTime12.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime12.sendKeys(Keys.BACK_SPACE);
		TestStartTime12.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime12.sendKeys(Keys.TAB);
		WebElement TestEndTime12 = wd.findElement(By.id("testEnd12"));
		TestEndTime12.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime12.sendKeys(Keys.BACK_SPACE);
		TestEndTime12.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime12.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row12\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		try {
		WebElement TestDate13 = wd.findElement(By.id("dttest13"));
		TestDate13.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject13"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty13"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic13")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent13")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime13 = wd.findElement(By.id("testStart13"));
		TestStartTime13.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime13.sendKeys(Keys.BACK_SPACE);
		TestStartTime13.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime13.sendKeys(Keys.TAB);
		WebElement TestEndTime13 = wd.findElement(By.id("testEnd13"));
		TestEndTime13.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime13.sendKeys(Keys.BACK_SPACE);
		TestEndTime13.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime13.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row13\"]/td[9]/button[1]")).click();
			} catch (org.openqa.selenium.StaleElementReferenceException e) {
				e.printStackTrace();
			}
		
		Thread.sleep(4000);
		
		WebElement TestDate14 = wd.findElement(By.id("dttest14"));
		TestDate14.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject14"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty14"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic14")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent14")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime14 = wd.findElement(By.id("testStart14"));
		TestStartTime14.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime14.sendKeys(Keys.BACK_SPACE);
		TestStartTime14.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime14.sendKeys(Keys.TAB);
		WebElement TestEndTime14 = wd.findElement(By.id("testEnd14"));
		TestEndTime14.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime14.sendKeys(Keys.BACK_SPACE);
		TestEndTime14.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime14.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row14\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate15 = wd.findElement(By.id("dttest15"));
		TestDate15.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject15"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty15"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic15")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent15")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime15 = wd.findElement(By.id("testStart15"));
		TestStartTime15.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime15.sendKeys(Keys.BACK_SPACE);
		TestStartTime15.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime15.sendKeys(Keys.TAB);
		WebElement TestEndTime15 = wd.findElement(By.id("testEnd15"));
		TestEndTime15.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime15.sendKeys(Keys.BACK_SPACE);
		TestEndTime15.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime15.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row15\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate16 = wd.findElement(By.id("dttest16"));
		TestDate16.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject16"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty16"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic16")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent16")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime16 = wd.findElement(By.id("testStart16"));
		TestStartTime16.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime16.sendKeys(Keys.BACK_SPACE);
		TestStartTime16.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime16.sendKeys(Keys.TAB);
		WebElement TestEndTime16 = wd.findElement(By.id("testEnd16"));
		TestEndTime16.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime16.sendKeys(Keys.BACK_SPACE);
		TestEndTime16.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime16.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row16\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate17 = wd.findElement(By.id("dttest17"));
		TestDate17.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject17"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty17"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic17")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent17")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime17 = wd.findElement(By.id("testStart17"));
		TestStartTime17.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime17.sendKeys(Keys.BACK_SPACE);
		TestStartTime17.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime17.sendKeys(Keys.TAB);
		WebElement TestEndTime17 = wd.findElement(By.id("testEnd17"));
		TestEndTime17.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime17.sendKeys(Keys.BACK_SPACE);
		TestEndTime17.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime17.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row17\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate18 = wd.findElement(By.id("dttest18"));
		TestDate18.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject18"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty18"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic18")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent18")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime18 = wd.findElement(By.id("testStart18"));
		TestStartTime18.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime18.sendKeys(Keys.BACK_SPACE);
		TestStartTime18.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime18.sendKeys(Keys.TAB);
		WebElement TestEndTime18 = wd.findElement(By.id("testEnd18"));
		TestEndTime18.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime18.sendKeys(Keys.BACK_SPACE);
		TestEndTime18.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime18.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row18\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate19 = wd.findElement(By.id("dttest19"));
		TestDate19.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject19"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty19"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic19")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent19")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime19 = wd.findElement(By.id("testStart19"));
		TestStartTime19.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime19.sendKeys(Keys.BACK_SPACE);
		TestStartTime19.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime19.sendKeys(Keys.TAB);
		WebElement TestEndTime19 = wd.findElement(By.id("testEnd19"));
		TestEndTime19.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime19.sendKeys(Keys.BACK_SPACE);
		TestEndTime19.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime19.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row19\"]/td[9]/button[1]")).click();
		Thread.sleep(2000);
		
		WebElement TestDate20 = wd.findElement(By.id("dttest20"));
		TestDate20.sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testSubject20"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("testfaculty20"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testTopic20")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("testContent20")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement TestStartTime20 = wd.findElement(By.id("testStart20"));
		TestStartTime20.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestStartTime20.sendKeys(Keys.BACK_SPACE);
		TestStartTime20.sendKeys(row.getCell(k++, blank).getStringCellValue());TestStartTime20.sendKeys(Keys.TAB);
		WebElement TestEndTime20 = wd.findElement(By.id("testEnd20"));
		TestEndTime20.sendKeys(Keys.chord(Keys.CONTROL, "a"));TestEndTime20.sendKeys(Keys.BACK_SPACE);
		TestEndTime20.sendKeys(row.getCell(k++, blank).getStringCellValue());TestEndTime20.sendKeys(Keys.TAB);
		wd.findElement(By.xpath("//*[@id=\"row20\"]/td[9]/button[1]")).click();
		Thread.sleep(1000);
		
		File scrFile = ((TakesScreenshot) wd).getScreenshotAs(OutputType.FILE);
		FileUtils.copyFile(scrFile, new File("E:/DOWNLOADS/Screenshot/Batch_ " + timestamp() + ".png"));
		SimpleDateFormat date = new SimpleDateFormat("dd-MM-yyyy HH-mm-ss");
			
		      }
    	    }
		  
		}
*//*	@Test(priority= 4)
	public void EmployeeData() throws IOException, InterruptedException {
		File src = new File("E:\\DOWNLOADS\\TestData.xlsx");
		FileInputStream finput = new FileInputStream(src);
		workbook = new XSSFWorkbook(finput);
		sheet = workbook.getSheetAt(3);
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		for (int i = 1; i < rowCount + 1; i++) {
			Row row = sheet.getRow(i);
			row.setRowNum(i);
		for (int k = 0; k < row.getLastCellNum(); k++) {
    	wd.findElement(By.xpath("//text()[contains(.,'Admin')]/ancestor::p[1]")).click();
    	wd.findElement(By.xpath("//p[text()='Employee']")).click();
		wd.findElement(By.id("btnAddProfile")).click();
        wd.findElement(By.name("FirstName")).sendKeys(row.getCell(k++, blank).getStringCellValue());
    	wd.findElement(By.name("MiddleName")).sendKeys(row.getCell(k++, blank).getStringCellValue());
	    wd.findElement(By.name("LastName")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement Date = wd.findElement(By.xpath("//input[@id='dtDOB']"));
		Date.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		Date.sendKeys(Keys.BACK_SPACE);
		Date.sendKeys(row.getCell(k++, blank).getStringCellValue());
		Date.sendKeys(Keys.TAB);
		wd.findElement(By.name("FatherName")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.name("SpouseName")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("txtAddress")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("txtEmail")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("txtCity")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("txtState")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("txtPincode")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("selDepartment"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("selBusinessTitle"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		new Select(wd.findElement(By.id("selCenterName"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("txtMobileNo")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("txtAnotherContactNo")).sendKeys(row.getCell(k++, blank).getStringCellValue());
		WebElement DOJ = wd.findElement(By.id("dtJoiningDate"));
		DOJ.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		DOJ.sendKeys(Keys.BACK_SPACE);
		DOJ.sendKeys(row.getCell(k++, blank).getStringCellValue());
        DOJ.sendKeys(Keys.TAB);
		new Select(wd.findElement(By.id("selActive"))).selectByValue(row.getCell(k++, blank).getStringCellValue());

		wd.findElement(By.id("btnSaveProfile")).click();
		Thread.sleep(5000);
		
              }

         }
	}
	@Test(priority= 5)
	public void MasterData() throws IOException, InterruptedException {
		File src = new File("E:\\DOWNLOADS\\TestData.xlsx");
		FileInputStream finput = new FileInputStream(src);
		workbook = new XSSFWorkbook(finput);
		sheet = workbook.getSheetAt(2);
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		for (int i = 1; i < rowCount + 1; i++) {
			Row row = sheet.getRow(i);
			row.setRowNum(i);
		for (int k = 0; k < row.getLastCellNum(); k++) {
    	wd.findElement(By.xpath("//text()[contains(.,'Admin')]/ancestor::p[1]")).click();
    	wd.findElement(By.xpath("//p[text()='Center']")).click();
		wd.findElement(By.id("btnAdd")).click();
   	    new Select(wd.findElement(By.id("sltOptionType"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
		wd.findElement(By.id("txtOptionName")).sendKeys(row.getCell(k++, blank).getStringCellValue());
	    wd.findElement(By.id("btnSave")).click();
	    wd.findElement(By.xpath("//text()[contains(.,'Admin')]/ancestor::p[1]")).click();
	    wd.findElement(By.xpath("//p[text()='Course']")).click();
		wd.findElement(By.id("txtCourseName")).sendKeys(row.getCell(k++, blank).getStringCellValue());
   	    new Select(wd.findElement(By.id("sltSubject"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
	    wd.findElement(By.id("btnSave")).click();
                 }
		     }
		 }
	@Test(priority= 6)
	public void ReportsData() throws IOException, InterruptedException {
		
    	wd.findElement(By.xpath("//text()[contains(.,'Reports')]/ancestor::p[1]")).click();
    	wd.findElement(By.xpath("//p[text()='Lead Report']")).click();
    	wd.findElement(By.xpath("//p[text()='Attendance Report']")).click();
		 new Select(wd.findElement(By.id("txtPreferredYear"))).selectByValue("");
		 new Select(wd.findElement(By.id("selBatchGroup"))).selectByIndex(0);
		 new Select(wd.findElement(By.id("selBatchCode"))).selectByValue("");


    	//
    	wd.findElement(By.xpath("//p[text()='Payment Report']")).click();


	}*/
	
		}
		
	
	
	
	