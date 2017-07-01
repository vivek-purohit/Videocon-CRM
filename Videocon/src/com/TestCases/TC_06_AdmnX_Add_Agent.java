/**
 * 
 */
package com.TestCases;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.POM.Pages.POM_Login;
import com.POM.Pages.POM_Operations;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

/**
 * @author Admin
 *
 */
public class TC_06_AdmnX_Add_Agent 
{
	WebDriver driver;
	WebDriverWait wait;
	ExtentReports reports;
	ExtentTest extent;
	POM_Login login;
	POM_Operations Ops;
	HSSFWorkbook workbook;
	HSSFSheet sheet;
	HSSFCell cell;
	
	@BeforeClass(alwaysRun = true)
	public void TestSetup()
	{
		driver = new FirefoxDriver();
		login = new POM_Login(driver);
		Ops = new POM_Operations(driver);
		login.Openurl();
		driver.manage().window().maximize();
		wait = new WebDriverWait(driver,30);
		reports = new ExtentReports("C:\\Users\\Admin\\Desktop\\Videcon Test Report Extent\\TC_06_AdmnX_Add_Agent.html", true);
		extent = reports.startTest("TC_06_AdmnX_Add_Agent", "To check the functionality of adding agent in AdminX");
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		
	}
	/*
	 * Step-1] Login using adminx credentials.
	 * Step-2] Open Add agent page. 
	 */
	
	@Test(priority = 0, groups = "test")
	public void AdminXLogin()
	{
		// Step-1] Login using adminx credentials.
		login.AdminXLogin();
		extent.log(LogStatus.INFO, "Logged in successfully");
		
		// Step-2] Open Add agent page. 
		wait.until(ExpectedConditions.elementToBeClickable(By.linkText("MASTERS")));
		Ops.adminX_OpenAddAgent();
		
	}
	
	/*
	 * Step -1] Import data for callCentreDD.
	 * Step -2] Import data for floorManagerDD.
	 * Step -3] Import data for teamLeader.
	 * Step -4] Import data for firstName.
	 * Step -5] Import data for lastName.
	 * Step -6] Import data for userName.
	 * Step -7] Import data for password.
	 * Step -8] Import data for address.
	 * Step -9] Import data for pincode.
	 * Step -10] Import data for contactNo.
	 * Step -11] Import data for mobileNo.
	 * Step -12] Import data for assignShiftDD.
	 * Step -13] Import data for languageDD.
	 * Step -14] Import data for weekOffDD.
	 * Step -15] Import data for employeeCode.
	 * Step -16] Import data for manualChkBox.
	 * Step -17] Import data for contaqueUser.
	 * Step -18] Import data for submitButton.
	 */
	@Test(priority = 1, groups = "test")
	public void AddAgent() throws IOException, InterruptedException
	{
		// Initialize ID of the elements.
		By callCentreDD = By.id("ContentPlaceHolder1_ucaddagent_ddlCCM");
		By floorManagerDD = By.id("ContentPlaceHolder1_ucaddagent_drpAM");
	    By teamLeaderDD = By.id("ContentPlaceHolder1_ucaddagent_ddlleader");
	    By firstName = By.id("ContentPlaceHolder1_ucaddagent_txtfname");
	    By lastName = By.id("ContentPlaceHolder1_ucaddagent_txtlname");
	    By userName = By.id("ContentPlaceHolder1_ucaddagent_txtagemail");
	    By password = By.id("ContentPlaceHolder1_ucaddagent_txtadpass");
	    By address = By.id("ContentPlaceHolder1_ucaddagent_txtaddress");
	    By pincode = By.id("ContentPlaceHolder1_ucaddagent_txtPinCode");
	    By contactNo = By.id("ContentPlaceHolder1_ucaddagent_txtcontactNo");
	    By mobileNo = By.id("ContentPlaceHolder1_ucaddagent_txtmobileNo");
	    By assignShiftDD = By.id("ContentPlaceHolder1_ucaddagent_ddlShift");
	    By languageDD = By.id("ContentPlaceHolder1_ucaddagent_ddlLanguage");
	    By weekOffDD = By.id("ContentPlaceHolder1_ucaddagent_ddlWeek");
	    By employeeCode = By.id("ContentPlaceHolder1_ucaddagent_txtempcode");
	    By manualChkBox = By.id("ContentPlaceHolder1_ucaddagent_chkMan");
	    By contaqueUser = By.id("ContentPlaceHolder1_ucaddagent_txtcontid");
	    By submitButton = By.id("ContentPlaceHolder1_ucaddagent_btnsubmit");
	    
	    
	    //  Import excel
	    File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_06_AdmnX_Add_Agent.xls");
	    FileInputStream fin = new FileInputStream(src);
	    workbook = new HSSFWorkbook(fin);
	    sheet = workbook.getSheetAt(0);
	    
	    // Iterate the flow.
	    for(int i = 1; i<=sheet.getLastRowNum(); i++)
	    {
	    	try 
	    	{
				wait.until(ExpectedConditions.elementToBeClickable(submitButton));
				
				// Step -1] Import data for callCentreDD.
				cell = sheet.getRow(i).getCell(1);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select callCentre = new Select(driver.findElement(callCentreDD));
				callCentre.selectByVisibleText(cell.getStringCellValue());
				
				// Step -2] Import data for floorManagerDD.
				Ops.WaitforLoadingImageToDisappaer();
				cell = sheet.getRow(i).getCell(2);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select floorManager = new Select(driver.findElement(floorManagerDD));
				floorManager.selectByVisibleText(cell.getStringCellValue());
				
				// Step -3] Import data for teamLeaderDD.
				Ops.WaitforLoadingImageToDisappaer();
				cell = sheet.getRow(i).getCell(3);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select teamLeader = new Select(driver.findElement(teamLeaderDD));
				teamLeader.selectByVisibleText(cell.getStringCellValue());
				
				// Step -4] Import data for firstName.
				cell = sheet.getRow(i).getCell(4);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(firstName).sendKeys(cell.getStringCellValue());
				
				// Step -5] Import data for lastName.
				cell = sheet.getRow(i).getCell(5);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(lastName).sendKeys(cell.getStringCellValue());
				
				// Step -6] Import data for userName.
				cell = sheet.getRow(i).getCell(6);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(userName).sendKeys(cell.getStringCellValue());
				
				// Step -7] Import data for password.
				cell = sheet.getRow(i).getCell(7);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(password).sendKeys(cell.getStringCellValue());
				
				// Step -8] Import data for address.
				cell = sheet.getRow(i).getCell(8);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(address).sendKeys(cell.getStringCellValue());
				
				// Step -9] Import data for pincode.
				cell = sheet.getRow(i).getCell(9);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(pincode).sendKeys(cell.getStringCellValue());
				driver.findElement(pincode).sendKeys(Keys.ENTER);
				
				// Step -10] Import data for contactNo.
				Ops.WaitforLoadingImageToDisappaer();
				//Thread.sleep(2500);
				cell = sheet.getRow(i).getCell(10);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(contactNo).sendKeys(cell.getStringCellValue());
				
				// Step -11] Import data for mobileNo.
				cell = sheet.getRow(i).getCell(11);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(mobileNo).sendKeys(cell.getStringCellValue());
				
				// Step -12] Import data for assignShiftDD.
				cell = sheet.getRow(i).getCell(12);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select assignShift = new Select(driver.findElement(assignShiftDD));
				assignShift.selectByVisibleText(cell.getStringCellValue());
				
				// Step -13] Import data for languageDD.
				cell= sheet.getRow(i).getCell(13);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select language = new Select(driver.findElement(languageDD));
				language.selectByVisibleText(cell.getStringCellValue());
				
				// Step -14] Import data for weekOffDD.
				cell = sheet.getRow(i).getCell(14);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select weekOff = new Select(driver.findElement(weekOffDD));
				weekOff.selectByVisibleText(cell.getStringCellValue());
				
				// Step -15] Import data for employeeCode.
				cell = sheet.getRow(i).getCell(15);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(employeeCode).sendKeys(cell.getStringCellValue());
				
				
				// Step -16] Import data for manualChkBox.
				cell = sheet.getRow(i).getCell(16);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				String statusChkBox = cell.getStringCellValue();
				if(statusChkBox.equalsIgnoreCase("Yes"))
				{
					driver.findElement(manualChkBox).click();
				}
				
				// Step -17] Import data for contaqueUser.
				cell = sheet.getRow(i).getCell(17);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(contaqueUser).sendKeys(cell.getStringCellValue());
				
				// Step -18] Import data for submitButton.
				driver.findElement(submitButton).click();
				
				try
				{
					Alert alert = driver.switchTo().alert();
					alert.accept();
				}
				
				catch(Exception e)
				{
					System.out.println("No alert present");
				}
				
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_ucaddagent_lblmsg")));
				String msgDisplayed = driver.findElement(By.id("ContentPlaceHolder1_ucaddagent_lblmsg")).getText();
				sheet.getRow(i).createCell(18).setCellValue(msgDisplayed);
				System.out.println(msgDisplayed);
				extent.log(LogStatus.INFO, msgDisplayed);
				FileOutputStream fout = new FileOutputStream(src);
				workbook.write(fout);
				fout.close();
				
				driver.get(driver.getCurrentUrl());
			} 
	    	catch (Exception e)
	    	{
				// TODO Auto-generated catch block
				e.printStackTrace();
				driver.get(driver.getCurrentUrl());
			}
	    	
	    	
	    	
	    	
	    	
	    	
	    }
	}
	

}
