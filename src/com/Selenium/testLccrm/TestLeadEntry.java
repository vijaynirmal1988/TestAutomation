		
package com.Selenium.testLccrm;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
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
	public void LeadData() throws IOException, InterruptedException {
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

		
//		  for (int i = 1; i < rowCount+1; i++) { 
//			  Row row = sheet.getRow(i);
//	              row.setRowNum(i); 
//		 for (int k = 2; k < row.getLastCellNum();k++) {
//		 		 Thread.sleep(2000); 

		  //select to the center to create a new leads 
		 //new Select(wd.findElement(By.id("selCenter"))).selectByValue(row.getCell(k++,blank).getStringCellValue()); 
//	     if(k < sheet.getLastRowNum()) { 
//	    	 wd.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);				
//	     
//		//  wd.findElement(By.xpath("//a[@class='nav-link collapsed']//p[contains(text(),'Leads')]")).click(); 
//		 // wd.findElement(By.xpath("//p[contains(text(),'Add New Lead')]")).click();
//	     }
 
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
		 //} 
		// wd.findElement(By.xpath("//button[text()='OK']")).click(); 
		// }
		 
		 
	}
	@Test(priority=1)
	public void RegisterData() throws IOException, InterruptedException {
		int j=0;
		JavascriptExecutor js = (JavascriptExecutor) wd;
		
		File src = new File("E:\\DOWNLOADS\\TestData.xlsx");
		FileInputStream finput = new FileInputStream(src);
		workbook = new XSSFWorkbook(finput);
		sheet = workbook.getSheetAt(5);
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		for (int i = 1; i < rowCount + 1; i++) {
			Row row = sheet.getRow(i);
			row.setRowNum(i);
		for (int k = 0; k < row.getLastCellNum(); k++) {	
    	for (j = 1; j <= 8 ; j++) {
    		 String nameXL = row.getCell(k,blank).getStringCellValue(); 
    		 String[] nameXLSplit = nameXL.split(",");
    		String name =wd.findElement(By.xpath("//*[@id='cmLeadStatusTable']/tbody/tr["+j+"]/td[3]")).getText();

    		String leadStatus =wd.findElement(By.xpath("//*[@id='cmLeadStatusTable']/tbody/tr["+j+"]/td[7]")).getText();
    		if(leadStatus.equals("Unassigned")) {
    			
    			for(int a=0;a<nameXLSplit.length;a++) {
        			if(nameXLSplit[a].equals(name)) {
        				 wd.findElement(By.xpath("//*[@id='cmLeadStatusTable']/tbody/tr["+j+"]/td[1]")).click(); 
        			}
        		}
    		} else if(leadStatus.equals("Inprogress")) {
    			 wd.findElement(By.xpath("//*[@id='cmLeadStatusTable']/tbody/tr["+j+"]/td[2]")).click(); 
    			 Thread.sleep(1000);
    			 wd.findElement(By.id("btnRegister")).click();
    			 
    		}
    	    wd.findElement(By.id("btnAssign")).click();
   	    	 new Select(wd.findElement(By.id("selAsnLeadStatus"))).selectByValue(row.getCell(k++,blank).getStringCellValue());
		     new Select(wd.findElement(By.id("selAsnCounselorID"))).selectByValue(row.getCell(k++,blank).getStringCellValue());
 		     wd.findElement(By.xpath("//*[@id=\"viewAssignModal\"]/div/div/div/div[2]/div/div/div[1]/button")).click();
		
    	}
    	for (j = 1; j <= 8 ; j++) {
    		String leadStatus =wd.findElement(By.xpath("//*[@id='cmLeadStatusTable']/tbody/tr["+j+"]/td[7]")).getText();
    		if(leadStatus.equals("Inprogress")) {
   			 wd.findElement(By.xpath("//*[@id='cmLeadStatusTable']/tbody/tr["+j+"]/td[2]")).click(); 
   			 Thread.sleep(1000);
   			 wd.findElement(By.id("btnRegister")).click();
   		}
    	}
    	
		}
                 }
		     }
		 
	@Test(priority = 2)
	public void BatchDataRead() throws IOException, InterruptedException {
		try {
			// Import excel sheet.
			File src = new File("E:\\DOWNLOADS\\TestData.xlsx");
			// Load the file.
			FileInputStream finput = new FileInputStream(src);
			// Load he workbook.
			workbook = new XSSFWorkbook(finput);

			// Load the sheet in which data is stored.
			sheet = workbook.getSheetAt(0);
			
		 int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
	        System.out.println(rowCount);
			
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
					StartDate.sendKeys(Keys.chord(Keys.CONTROL, "a"));
					StartDate.sendKeys(Keys.BACK_SPACE);
					StartDate.sendKeys(row.getCell(k++, blank).getStringCellValue());
					StartDate.sendKeys(Keys.TAB);
					WebElement EndDate = wd.findElement(By.id("txtBatchEndDate"));
					EndDate.sendKeys(Keys.chord(Keys.CONTROL, "a"));
					EndDate.sendKeys(Keys.BACK_SPACE);
					EndDate.sendKeys(row.getCell(k++, blank).getStringCellValue());
					EndDate.sendKeys(Keys.TAB);
					wd.findElement(By.id("btnAdd")).click();
					// Enter to add Course Details
					new Select(wd.findElement(By.id("facultyVerbal"))).selectByValue(row.getCell(k++, blank).getStringCellValue());
					wd.findElement(By.id("noClassVerbal")).sendKeys("1");
					wd.findElement(By.id("noTestVerbal")).sendKeys("1");
					wd.findElement(By.id("btnCourseAdd")).click();

					// Enter to add Class schedule
					WebElement ClassDate = wd.findElement(By.id("dtclass1"));
					ClassDate.sendKeys("05/04/2020");
					new Select(wd.findElement(By.id("classSubject1"))).selectByValue("Verbal");
					new Select(wd.findElement(By.id("classfaculty1"))).selectByValue("1011");

					wd.findElement(By.id("classTopic1")).sendKeys("verb");
					wd.findElement(By.id("classContent1")).sendKeys("verb");
					WebElement StartTime = wd.findElement(By.id("classStart1"));
					StartTime.sendKeys(Keys.chord(Keys.CONTROL, "a"));
					StartTime.sendKeys(Keys.BACK_SPACE);
					StartTime.sendKeys("02:45PM");
					StartTime.sendKeys(Keys.TAB);
					WebElement EndTime = wd.findElement(By.id("classEnd1"));
					EndTime.sendKeys(Keys.chord(Keys.CONTROL, "a"));
					EndTime.sendKeys(Keys.BACK_SPACE);
					EndTime.sendKeys("04:45PM");
					EndTime.sendKeys(Keys.TAB);
					wd.findElement(By.xpath("//*[@id=\"row1\"]/td[9]/button[1]\r\n")).click();

					// Enter to add Test Schedule
					WebElement TestDate = wd.findElement(By.id("dttest1"));
					TestDate.sendKeys("05/30/2020");
					new Select(wd.findElement(By.id("testSubject1"))).selectByValue("Verbal");
					new Select(wd.findElement(By.id("testfaculty1"))).selectByValue("1011");
					wd.findElement(By.id("testTopic1")).sendKeys("verb");
					wd.findElement(By.id("testContent1")).sendKeys("verb");
					WebElement TestStartTime = wd.findElement(By.id("testStart1"));
					TestStartTime.sendKeys(Keys.chord(Keys.CONTROL, "a"));
					TestStartTime.sendKeys(Keys.BACK_SPACE);
					TestStartTime.sendKeys("02:45PM");
					TestStartTime.sendKeys(Keys.TAB);
					WebElement TestEndTime = wd.findElement(By.id("testEnd1"));
					TestEndTime.sendKeys(Keys.chord(Keys.CONTROL, "a"));
					TestEndTime.sendKeys(Keys.BACK_SPACE);
					TestEndTime.sendKeys("04:45PM");
					TestEndTime.sendKeys(Keys.TAB);
					wd.findElement(By.xpath("//*[@id=\"row1\"]/td[9]/button[1]")).click();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		File scrFile = ((TakesScreenshot) wd).getScreenshotAs(OutputType.FILE);
		FileUtils.copyFile(scrFile, new File("E:/DOWNLOADS/Screenshot/Batch_ " + timestamp() + ".png"));
		SimpleDateFormat date = new SimpleDateFormat("dd-MM-yyyy HH-mm-ss");
		Thread.sleep(200);

	}
	@Test(priority=2)
	public void EmployeeData() throws IOException, InterruptedException {
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
	@Test(priority=3)
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
	@Test(priority=3)
	public void ReportsData() throws IOException, InterruptedException {
		
    	wd.findElement(By.xpath("//text()[contains(.,'Reports')]/ancestor::p[1]")).click();
    	wd.findElement(By.xpath("//p[text()='Lead Report']")).click();
    	wd.findElement(By.xpath("//p[text()='Attendance Report']")).click();
		 new Select(wd.findElement(By.id("txtPreferredYear"))).selectByValue("");
		 new Select(wd.findElement(By.id("selBatchGroup"))).selectByIndex(0);
		 new Select(wd.findElement(By.id("selBatchCode"))).selectByValue("");


    	//
    	wd.findElement(By.xpath("//p[text()='Payment Report']")).click();


	}
	}