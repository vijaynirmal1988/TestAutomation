package com.Selenium.testLccrm;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
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

		 @BeforeSuite
		    public void beforeSuite() {
		       
		        System.setProperty("webdriver.chrome.driver", "E:\\DOWNLOADS\\chromedriver.exe");
		        wd = new ChromeDriver();
		     // Enter url.
		     //	wd.get("http://localhost:80/lccrm/lccrmUI/user/userLogin.php");
		        wd.get("http://stagingenvironmentlccrm.cloudaccess.host/lccrm/lccrmUI/user/userLogin.php");
		     	wait = new WebDriverWait(wd,30);
				wd.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
				wd.manage().window().maximize();
		 }

	 @Test(priority=0)
	 public  void ReadData() throws IOException, InterruptedException
	 {
	 
	     // Import excel sheet.
		 File src=new File("E:\\DOWNLOADS\\Report 1.xlsx");
		 
		 // Load the file.
		 FileInputStream finput = new FileInputStream(src);
		 
		 // Load he workbook.
	    workbook = new XSSFWorkbook(finput);
	    
	    MissingCellPolicy blank = Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

		  // Load the sheet in which data is stored.
		 sheet= workbook.getSheetAt(0);

			int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();

			//for(int j=1 ; j < rowCount + 1 ; j++) {
			//	Row row = sheet.getRow(j);
			//	row.setRowNum(j);
				
				 // Import data for Email.
			       cell = sheet.getRow(1).getCell(0,blank);
			       cell.setCellType(CellType.STRING);
	         wd.findElement(By.id("txtemail")).sendKeys(cell.getStringCellValue());
			 
			 // Import data for password.
		    	cell = sheet.getRow(1).getCell(1,blank);
			    cell.setCellType(CellType.STRING);
			    wd.findElement(By.id("txtpwd")).sendKeys(cell.getStringCellValue());
			 //Enter to Login
			    wd.findElement(By.id("btnSave")).click();
			  //  Thread.sleep(2000);
			  //  Alert alert = wd.switchTo().alert();
			   // String alertMessage= wd.switchTo().alert().getText();
			   // alert.accept();
		//	}
		
		    Select center = new Select(wd.findElement(By.id("selCenter")));
			center.selectByIndex(3);
			
					for (int i = 1; i < rowCount+1; i++) {
						Row row = sheet.getRow(i);
						row.setRowNum(i);
						
						for (int k = 2; k < row.getLastCellNum(); k++) {

			
				 if(k < sheet.getLastRowNum()) {
				    	wd.findElement(By.xpath("//text()[contains(.,'Leads')]/ancestor::p[1]")).click();
				    	wd.findElement(By.xpath("//p[text()='Add New Lead']")).click();
				    }
		
	            	
	            wd.findElement(By.name("firstname")).sendKeys(row.getCell(k++,blank).getStringCellValue());
	           
				wd.findElement(By.name("middlename")).sendKeys(row.getCell(k++,blank).getStringCellValue());
	           
                wd.findElement(By.name("lastname")).sendKeys(row.getCell(k++,blank).getStringCellValue());
	           
                WebElement Date = wd.findElement(By.xpath("//input[@type='text'][@name='Dob']"));
			    Date.sendKeys(Keys.chord(Keys.CONTROL,"a"));
			    Date.sendKeys(Keys.BACK_SPACE);
			    Date.sendKeys(String.valueOf(row.getCell(k++,blank).getStringCellValue()));
			    Date.sendKeys(Keys.TAB);

			  
	            wd.findElement(By.xpath("//div[@class='form-check form-check-radio']//label[1]")).click();
			   
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
	          
				Select dd = new Select(wd.findElement(By.id("selCurrentEducationStatus")));
				dd.selectByIndex(10);
	           
                wd.findElement(By.id("txtInstitutionName")).sendKeys(row.getCell(k++,blank).getStringCellValue());
				
                Select dd1 = new Select(wd.findElement(By.id("selAnnualIncome")));
				dd1.selectByIndex(1);
				Select dd2 = new Select(wd.findElement(By.id("selStreamofEducation")));
				dd2.selectByIndex(1);
				Select dd3 = new Select(wd.findElement(By.id("selTypeofIndustry")));
				dd3.selectByIndex(1);
				
				wd.findElement(By.name("WorkExperience")).sendKeys(row.getCell(k++,blank).getStringCellValue());
	            
				wd.findElement(By.name("companyname")).sendKeys(row.getCell(k++,blank).getStringCellValue());
				wd.findElement(By.xpath("//*[@id=\"newform\"]//div[10]//div[1]//div[1]")).click();
				wd.findElement(By.xpath("//*[@id=\"newform\"]//div[10]//div[1]//div[2]")).click();
				wd.findElement(By.xpath("//*[@id=\"newform\"]//div[10]//div[1]//div[3]")).click();
				wd.findElement(By.xpath("//*[@id=\"newform\"]//div[10]//div[1]//div[4]")).click();
				wd.findElement(By.xpath("//*[@id=\"newform\"]//div[13]//div[1]//div[1]")).click();
				wd.findElement(By.xpath("//*[@id=\"newform\"]//div[13]//div[1]//div[2]")).click();
				wd.findElement(By.xpath("//*[@id=\"newform\"]//div[13]//div[1]//div[3]")).click();
	            wd.findElement(By.xpath("//div[@class='col-sm-4 col-sm-offset-1']//div[1]//label[1]")).click();
	            wd.findElement(By.xpath("//div[@class='col-sm-4 col-sm-offset-1']//div[3]//label[1]")).click();
	            wd.findElement(By.xpath("//div[@class='col-sm-4 col-sm-offset-1']//div[4]//label[1]")).click();
	            wd.findElement(By.xpath("//div[@class='col-sm-5']//div[1]//label[1]")).click();
	            wd.findElement(By.xpath("//div[@class='col-sm-5']//div[4]//label[1]")).click();
	           
				wd.findElement(By.id("txtAddFriendName1")).sendKeys(row.getCell(k++,blank).getStringCellValue());
	           
			    wd.findElement(By.id("txtAddFriendMobile1")).sendKeys(row.getCell(k++,blank).getStringCellValue());
	           
				wd.findElement(By.id("txtAddFriendEmail1")).sendKeys(row.getCell(k++,blank).getStringCellValue());
	         
				wd.findElement(By.id("txtAddFriendName2")).sendKeys(row.getCell(k++,blank).getStringCellValue());
	           
				wd.findElement(By.id("txtAddFriendMobile2")).sendKeys(row.getCell(k++,blank).getStringCellValue());
	           
				wd.findElement(By.id("txtAddFriendEmail2")).sendKeys(row.getCell(k++,blank).getStringCellValue());
	           
				Select dd4 = new Select(wd.findElement(By.id("selCourseEstimatedStart")));
				dd4.selectByIndex(2);
				Select dd5 = new Select(wd.findElement(By.id("selClassSchedulePref")));
				dd5.selectByIndex(2);
				Select dd6 = new Select(wd.findElement(By.id("selHasMagooshAccount")));
				dd6.selectByIndex(1);
				wd.findElement(By.xpath("//div[@class='col-sm-12']//div[1]//label[1]")).click();
				wd.findElement(By.xpath("//div[@class='col-sm-12']//div[5]//label[1]")).click();
			    wd.findElement(By.id("btnSave")).click();
			    Thread.sleep(2000);
            }
			    }
	 }
		// @Test(priority=1)
		 public  void BatchDataRead() throws IOException, InterruptedException
		 {
		 
			    
		         //Enter to Batch details
				 wd.findElement(By.xpath("//text()[contains(.,'Batch Details')]/ancestor::p[1]")).click();
				//Enter to Batch Group Master
		         wd.findElement(By.xpath("//p[text()='Batch Group Master']	")).click();
		         Thread.sleep(2000);
		         //add Batch Group
		         wd.findElement(By.xpath("//text()[.='Add']/ancestor::button[1]")).click();
		         Select Year = new Select(wd.findElement(By.id("selBatchGroupYear")));
				 Year.selectByIndex(1);
				 Select YearQuarter = new Select(wd.findElement(By.id("sltBatchGroupQuater")));
				 YearQuarter.selectByIndex(1);
				 wd.findElement(By.id("btnBGroupSave")).click();
				 //Enter to batch Master 
				 wd.findElement(By.xpath("//text()[contains(.,'Batch Details')]/ancestor::p[1]")).click();
		         wd.findElement(By.xpath("//p[text()='Batch Master']")).click();
		       //Enter to add batch details
		         wd.findElement(By.id("btnAddBatch")).click();
		         Select BatchGroupCode = new Select(wd.findElement(By.id("selBatchGroupCode")));
		         BatchGroupCode.selectByIndex(1);
				 Select CourseList = new Select(wd.findElement(By.id("selCourseList")));
				 CourseList.selectByIndex(1);
				 Select BatchType = new Select(wd.findElement(By.id("selBatchType")));
				 BatchType.selectByIndex(1);
				 Select BatchSession = new Select(wd.findElement(By.id("selBatchSession")));
				 BatchSession.selectByIndex(1);
				 WebElement StartDate = wd.findElement(By.id("txtBatchStartDate"));
				 StartDate.sendKeys("10/03/2019");
				 WebElement EndDate = wd.findElement(By.id("txtBatchEndDate"));
				 EndDate.sendKeys("12/25/2019");
		         wd.findElement(By.id("btnAdd")).click();
		         //Enter to add Course Details
		         Select facultyVerbal = new Select(wd.findElement(By.id("facultyVerbal")));
		         facultyVerbal.selectByIndex(1);
		         wd.findElement(By.id("noClassVerbal")).sendKeys("2");
		         wd.findElement(By.id("noTestVerbal")).sendKeys("2");
		         //Enter to add Class schedule 
		         WebElement ClassDate = wd.findElement(By.id("dtclass1"));
				 ClassDate.sendKeys("11/06/2019");
		         Select ClassSub = new Select(wd.findElement(By.id("classSubject1")));
		         ClassSub.selectByIndex(1);
		         wd.findElement(By.id("classTopic1")).sendKeys("verb");
		         wd.findElement(By.id("classContent1")).sendKeys("verb");
		         WebElement classStart1 = wd.findElement(By.id("classStart1"));
		         classStart1.sendKeys("05:30");
				 WebElement classEnd1 = wd.findElement(By.id("classEnd1"));
				 classEnd1.sendKeys("05:45");
				 //Enter to add Test Schedule
				 WebElement TestDate = wd.findElement(By.id("dtclass1"));
				 TestDate.sendKeys("11/06/2019");
			     Select testSubject1 = new Select(wd.findElement(By.id("testSubject1")));
			     testSubject1.selectByIndex(1);
			     wd.findElement(By.id("testTopic1")).sendKeys("verb");
			     wd.findElement(By.id("testContent1")).sendKeys("verb");
			     WebElement testStart1 = wd.findElement(By.id("classStart1"));
			     testStart1.sendKeys("05:30");
			     WebElement testEnd1 = wd.findElement(By.id("classEnd1"));
				 testEnd1.sendKeys("05:45");
		         wd.findElement(By.id("btnCourseAdd")).click();
		        
	        }
	    }
	  
		
