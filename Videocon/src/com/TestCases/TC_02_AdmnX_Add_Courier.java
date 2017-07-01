/**
 * 
 */
package com.TestCases;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.ITestResult;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.POM.Pages.POM_Login;
import com.POM.Pages.POM_Operations;
//import com.Snapdeal.POM.Pages.AdminX_Login_And_Add_View_Pages;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;



/**
 * @author Admin
 *
 */
public class TC_02_AdmnX_Add_Courier 
{
	WebDriver driver;
	WebDriverWait wait;
	ExtentReports reports;
	ExtentTest extent;
	Alert alert;
	POM_Login login;
	POM_Operations Ops;
	//Initializing the locators of the page.
	By firstName = By.id("ContentPlaceHolder1_uccourier_txtfname");
	By lastName  = By.id("ContentPlaceHolder1_uccourier_txtlname");
	By userName  = By.id("ContentPlaceHolder1_uccourier_txtagemail");
	By password  = By.id("ContentPlaceHolder1_uccourier_txtadpass");
	By address   = By.id("ContentPlaceHolder1_uccourier_txtaddress");
	By pincode   = By.id("ContentPlaceHolder1_uccourier_txtPinCode");
	By contactNo = By.id("ContentPlaceHolder1_uccourier_txtcontactNo");
	By mobileNo  = By.id("ContentPlaceHolder1_uccourier_txtmobileNo");
	By priority  =  By.id("ContentPlaceHolder1_uccourier_ddlpriority");
	By rpiPriority = By.id("ContentPlaceHolder1_uccourier_ddlrpiprio");
	By noOfChallans = By.id("ContentPlaceHolder1_uccourier_txtnoofchallan");
	By noOFRetailInvoices = By.id("ContentPlaceHolder1_uccourier_txtretailinvoice");
	By noOFDeclarations   = By.id("ContentPlaceHolder1_uccourier_txtdeclaration");
	By MF = By.id("ContentPlaceHolder1_uccourier_txtmultiply");
	By DF = By.id("ContentPlaceHolder1_uccourier_txtdivide");
	By category = By.id("ContentPlaceHolder1_uccourier_ddlcarriercat");
	By minWeight = By.id("ContentPlaceHolder1_uccourier_txtminWeight");
	By maxWeight = By.id("ContentPlaceHolder1_uccourier_txtmaxweight");
	By submit = By.id("ContentPlaceHolder1_uccourier_btnsubmit");
	By msgDisplayed = By.id("ContentPlaceHolder1_uccourier_lblmsg");
	By loadingimage = By.id("loading-image");
	@BeforeClass(alwaysRun = true)
	public void TestStart()
	{
		driver = new FirefoxDriver();
	    driver.get("http://192.168.1.201:7111");
	    driver.manage().window().maximize();
	    driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	    wait = new WebDriverWait(driver,30);
	    reports = new ExtentReports("C:\\Users\\Admin\\Desktop\\Videcon Test Report Extent\\TC_02_AdmnX_Add_Courier.html", true);
	    extent = reports.startTest("TC_025_AdmnX_Add_Courier", "To check the functionality of Adding courier with different data sets");
	    
	}
	@Test(priority = 0, groups = "test")
	public void Add_Courier_Test_Login()
	{
		try
		{
			POM_Login login = new POM_Login(driver);
			Ops = new POM_Operations(driver);
			//Initiate login process
			login.AdminXLogin();	
			extent.log(LogStatus.INFO, "Logged in successfully");
			//Open view coupons page.
			Ops.clickAdminSettings();
			wait.until(ExpectedConditions.elementToBeClickable(By.linkText("MASTERS")));
			//Open add courier page.
			Ops.OpenAddCourier();
			extent.log(LogStatus.INFO, "Add Courier page opened");
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		
	}
	@Test(priority = 1)
	public void TestRun() throws InterruptedException, IOException
	{
		File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_02_Add_Courier.xls");
		FileInputStream fileinput = new FileInputStream(src);

		HSSFWorkbook wb = new HSSFWorkbook(fileinput);
		HSSFSheet sheet = wb.getSheetAt(0);
		HSSFCell cell;
		CellStyle style = wb.createCellStyle();
		Font font = wb.createFont();
		
		
		//int n = sheet.getLastRowNum();
		System.out.println("Total number of rows = " +sheet.getLastRowNum());
		for(int i = 1; i<=sheet.getLastRowNum();i++)
		{
			try
			{
				
			//Enter data for first name.
			cell = sheet.getRow(i).getCell(0);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			//driver.findElement(firstName).sendKeys(cell.getStringCellValue());
			driver.findElement(firstName).sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
			extent.log(LogStatus.INFO, "First name = " +sheet.getRow(i).getCell(0).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell for first name");
				
			}
			
			try
			{
			//Enter data for last name.
			cell = sheet.getRow(i).getCell(1);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(lastName).sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
			extent.log(LogStatus.INFO,"Last name = " +sheet.getRow(i).getCell(1).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			
			try
			{
			//Enter data for user name.
			cell = sheet.getRow(i).getCell(2);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(userName).sendKeys(sheet.getRow(i).getCell(2).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
			}
			
			try
			{
			//Enter data for password
			cell = sheet.getRow(i).getCell(3);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(password).sendKeys(sheet.getRow(i).getCell(3).getStringCellValue());
			}
	
		   catch(Exception e)
		    {
			  System.out.println("Blank cell");
			 
		    }
			
			try
			{
			//Enter data for address.
			cell = sheet.getRow(i).getCell(4);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(address).sendKeys(sheet.getRow(i).getCell(4).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			//Enter data for pin code.
			try
			{
			cell = sheet.getRow(i).getCell(5);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(pincode).sendKeys(sheet.getRow(i).getCell(5).getStringCellValue());
			driver.findElement(contactNo).click();
			wait.until(ExpectedConditions.invisibilityOfElementLocated(loadingimage));
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			try
			{
			//Enter data for contact number.
			cell = sheet.getRow(i).getCell(6);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(contactNo).sendKeys(sheet.getRow(i).getCell(6).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			try
			{
			//Enter data for mobile number
			cell = sheet.getRow(i).getCell(7);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(mobileNo).sendKeys(sheet.getRow(i).getCell(7).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			try
			{
			//Select the priority.
			cell = sheet.getRow(i).getCell(8);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			Select priorityDD = new Select(driver.findElement(priority));
			String dataPriority = sheet.getRow(i).getCell(8).getStringCellValue();
			priorityDD.selectByVisibleText(dataPriority);
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			try
			{
			// Select RPI priority.
			cell = sheet.getRow(i).getCell(9);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			Select RPIPriorityDD = new Select(driver.findElement(rpiPriority));
			String dataRPIpriority = sheet.getRow(i).getCell(9).getStringCellValue();
			RPIPriorityDD.selectByVisibleText(dataRPIpriority);
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			try
			{
			//Enter data for number of challans.
			cell = sheet.getRow(i).getCell(10);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(noOfChallans).sendKeys(sheet.getRow(i).getCell(10).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			try
			{
			// Enter data for No. of retail invoices.
			cell = sheet.getRow(i).getCell(11);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(noOFRetailInvoices).sendKeys(sheet.getRow(i).getCell(11).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			try
			{
			//Enter data for Number of declaration.
			cell = sheet.getRow(i).getCell(12);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(noOFDeclarations).sendKeys(sheet.getRow(i).getCell(12).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			try
			{
			//Enter data for multiply factor.
			cell = sheet.getRow(i).getCell(13);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(MF).sendKeys(sheet.getRow(i).getCell(13).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			//Click on AWB Import bulk sheet.
			driver.findElement(By.id("ContentPlaceHolder1_uccourier_chkbarcode_0")).click();
			
			try
			{
			//Enter data for Dividing factor.
			cell = sheet.getRow(i).getCell(14);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(DF).sendKeys(sheet.getRow(i).getCell(14).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			try
			{
			//Enter data for category.
			cell = sheet.getRow(i).getCell(15);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			Select categoryDD = new Select(driver.findElement(category));
			String categoryData = sheet.getRow(i).getCell(15).getStringCellValue();
			categoryDD.selectByValue(categoryData);
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			try
			{
			//Enter data for Minimum weight.
			cell = sheet.getRow(i).getCell(16);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(minWeight).sendKeys(sheet.getRow(i).getCell(16).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			try
			{
			//Enter data for maximum weight.
			cell = sheet.getRow(i).getCell(17);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			driver.findElement(maxWeight).sendKeys(sheet.getRow(i).getCell(17).getStringCellValue());
			}
			catch(Exception e)
			{
				System.out.println("Blank cell");
				
			}
			
			
			//Click submit button.
		    driver.findElement(submit).click();
			Thread.sleep(3000);
			
			try
			{
			alert = driver.switchTo().alert();
			String alrtMsg = alert.getText();
			System.out.println("Message displayed for alert = " +alrtMsg);
			alert.accept();
			//Checking for submitting record.
			if(! alrtMsg.equalsIgnoreCase("Are you sure to Submit this record?"))
			{
				sheet.getRow(i).createCell(18).setCellValue(alrtMsg);
				sheet.getRow(i).createCell(20).setCellValue("Fail");
				FileOutputStream fout = new FileOutputStream(src);
				wb.write(fout);
				fout.close();
			}
			else
			{
				System.out.println("Message displayed for alert = " +alrtMsg);
			}
			//Writing the data for the alert displayed.
			
			
			
			//catch(Exception e)
			//{
			// wait for message to be displayed
			    wait.until(ExpectedConditions.visibilityOfElementLocated(msgDisplayed));
				String msg = driver.findElement(msgDisplayed).getText();
				if(!msg.contains("Saved Successfully!!!"))
				{
					sheet.getRow(i).createCell(18).setCellValue("No alert");
					//Set the status in the Main result cell.
					sheet.getRow(i).createCell(19).setCellValue(msg);
					sheet.getRow(i).createCell(20).setCellValue("Fail");
					FileOutputStream fout = new FileOutputStream(src);
					wb.write(fout);
					fout.close();
					driver.get(driver.getCurrentUrl());
				}
				else
				{
					sheet.getRow(i).createCell(18).setCellValue("No alert");
					//Set the status in the Main result cell.
					sheet.getRow(i).createCell(19).setCellValue(msg);
					sheet.getRow(i).createCell(20).setCellValue("Pass");
					FileOutputStream fout = new FileOutputStream(src);
					wb.write(fout);
					fout.close();
				}
		}
		catch(Exception e)
		{
			e.printStackTrace();
				
			}
			
				
				/*WritableWorkbook wb = Workbook.createWorkbook(new File("C:\\Users\\Admin\\Desktop\\DD_Add_Courier_Output.xls"));
		    	WritableSheet ws = wb.createSheet("customsheet",1);
		    	{
		    	Label label = new Label(0,0,"test");
		    	ws.addCell(label);
		    	}
		    	wb.write();
		    	wb.close();*/
				
				
				/*FileOutputStream fileoutput = new FileOutputStream("C:\\DD_Add_Courier.xls");
				Workbook workbook = Workbook.getWorkbook(new File("C:\\DD_Add_Courier.xls"));
				WritableWorkbook workbookCopy = Workbook.createWorkbook(new File("testSampleDataCopy.xls"), workbook);
				WritableSheet wSheet = workbookCopy.getSheet(0);
				Label label= new Label(18, i, "pass");
				wSheet.addCell(label);*/
			
			
			
		}
		
			
			
			
			
			
			
			
			/*//Data input for first name.
			driver.findElement(firstName).sendKeys(sheet.getCell(0, i).getContents());
			extent.log(LogStatus.INFO, "Data used for First name = " +sheet.getCell(0, i).getContents());
			//Click on submit button to check validity of data.
			driver.findElement(submit).click();
			try
			{
				alert = driver.switchTo().alert();
				if(alert.getText().equalsIgnoreCase("Only Characters are allowed in First Name!") | alert.getText().equalsIgnoreCase("Please Enter User Name"))
				{
					extent.log(LogStatus.INFO, "Message displayed in the alert box for First name field is = "+alert.getText());
					System.out.println("Message displayed in the alert box for First name field is = "+alert.getText());
					alert.accept();
					driver.findElement(firstName).sendKeys("AutoTest");
				}
				else
				{
					alert.accept();
				}
				
			}
			catch(Exception e)
			{
				System.out.println("No alert box found");
			}
			//Data input for last name.
			driver.findElement(lastName).sendKeys(sheet.getCell(1, i).getContents());
			extent.log(LogStatus.INFO, "Data used for First name = " +sheet.getCell(1, i).getContents());
			driver.findElement(submit).click();
			try
			{
				alert = driver.switchTo().alert();
				if(alert.getText().equalsIgnoreCase("Only Characters are allowed in Last Name!") | alert.getText().equalsIgnoreCase("Please Enter Last Name"))
				{
					extent.log(LogStatus.INFO, "Message displayed in the alert box for First name field is = "+alert.getText());
					System.out.println("Message displayed in the alert box for First name field is = "+alert.getText());
					alert.accept();
					driver.findElement(lastName).sendKeys("AutoTest");
				}
				else
				{
					alert.accept();
				}
				
			}
			catch(Exception e)
			{
				System.out.println("No alert box found");
			}
			// Data input for user name
			driver.findElement(userName).sendKeys(sheet.getCell(2, i).getContents());
			extent.log(LogStatus.INFO, "Data used for First name = " +sheet.getCell(2, i).getContents());
			driver.findElement(submit).click();
			try
			{
				alert = driver.switchTo().alert();
				if(alert.getText().equalsIgnoreCase("Please Enter User Name"))
				{
					extent.log(LogStatus.INFO, "Message displayed in the alert box for First name field is = "+alert.getText());
					System.out.println("Message displayed in the alert box for First name field is = "+alert.getText());
					alert.accept();
					driver.findElement(lastName).sendKeys("AutoTest");
				}
				else
				{
					alert.accept();
				}
				
			}
			catch(Exception e)
			{
				System.out.println("No alert box found");
			}
			// Data input for password
			driver.findElement(password).sendKeys(sheet.getCell(3, i).getContents());
			extent.log(LogStatus.INFO, "Data used for First name = " +sheet.getCell(3, i).getContents());
			driver.findElement(submit).click();
			try
			{
				alert = driver.switchTo().alert();
				if(alert.getText().equalsIgnoreCase("Please Enter Password"))
				{
					extent.log(LogStatus.INFO, "Message displayed in the alert box for First name field is = "+alert.getText());
					System.out.println("Message displayed in the alert box for First name field is = "+alert.getText());
					alert.accept();
					driver.findElement(lastName).sendKeys("123456");
				}
				else
				{
					alert.accept();
				}
				
			}
			catch(Exception e)
			{
				System.out.println("No alert box found");
			}
			// Data input for pin code.
			driver.findElement(pincode).sendKeys(sheet.getCell(4, i).getContents());
			extent.log(LogStatus.INFO, "Data used for First name = " +sheet.getCell(4, i).getContents());
			driver.findElement(submit).click();
			try
			{
				alert = driver.switchTo().alert();
				if(alert.getText().equalsIgnoreCase("Please Enter Pincode"))
				{
					extent.log(LogStatus.INFO, "Message displayed in the alert box for First name field is = "+alert.getText());
					System.out.println("Message displayed in the alert box for First name field is = "+alert.getText());
					alert.accept();
					driver.findElement(lastName).sendKeys("313324");
				}
				else
				{
					alert.accept();
				}
				If()
				{
					
				}
				
				
			}
			catch(Exception e)
			{
				System.out.println("No alert box found");
			}
			
			
		}*/
		
		}
	@AfterMethod
	public void teardown(ITestResult result) throws IOException
	{
		if(result.getStatus() == ITestResult.FAILURE)
		{
			extent.log(LogStatus.FAIL, "Test case which failed is = " +result.getName());
			System.out.println("Test case which failed is = " +result.getName());
			File src = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(src, new File("C:\\Selenium\\Screenshots\\TC_025\\Msg1.png"));
			extent.addScreenCapture("C:\\Selenium\\Screenshots\\TC_025\\Msg1.png");
		}
		if(result.getStatus() == ITestResult.SUCCESS)
		{
			extent.log(LogStatus.FAIL, "Test case which passed is = " +result.getName());
			System.out.println("Test case which passed is = " +result.getName());
		}
		if(result.getStatus() == ITestResult.SKIP)
		{
			extent.log(LogStatus.FAIL, "Test case which was Skipped is = " +result.getName());
			System.out.println("Test case which was Skipped is = " +result.getName());
		}
		
	}
	@AfterClass(alwaysRun = true)
	public void TestClosure()
	{
		//driver.quit();
		reports.endTest(extent);
		reports.flush();
	}
			

}
