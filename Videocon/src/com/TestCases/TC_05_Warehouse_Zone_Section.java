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
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
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
public class TC_05_Warehouse_Zone_Section
{
	WebDriver driver;
	WebDriverWait wait;
	ExtentReports reports;
	ExtentTest extent;
	POM_Login login;
	POM_Operations Ops;
	HSSFWorkbook wb;
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
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		wait = new WebDriverWait(driver,30);
		reports = new ExtentReports("C:\\Users\\Admin\\Desktop\\Videcon Test Report Extent\\TC_05_Warehouse_Zone_Section.html", true);
	    extent = reports.startTest("TC_05_Warehouse_Zone_Section", "To check the functionality of adding Zone, Section, Shelf, Rack in the warehouse");
	}
	/*
	 * Step-1] Login using warehouse credentials.
	 * Step-2] Open View Zone page. 
	 */
	@Test(priority = 0, groups = "test")
	public void Login()
	{
		login.WarehouseLogin();
		wait.until(ExpectedConditions.elementToBeClickable(By.linkText("MASTERS")));
	    Ops.warehouse_OpenViewZone();
	    extent.log(LogStatus.INFO, "Logged in successfully.");
	}
	
	/*
	 * Step-1] Open Add Zone page.
	 * Step-2] Import data for Zone Name.
	 * Step-3] Import data for Capacity.
	 * Step-4] Click on submit button.
	 * Step-5] Accept alert.
	 * Step-6] Set message in the excel.
	 */
	@Test(priority = 1, groups = "test")
	public void AddZone() throws IOException
	{
		
			// Specify Ids.
			By addZoneBtn = By.id("ContentPlaceHolder1_btnCreate");
			By zoneNameTxtBox = By.id("ContentPlaceHolder1_txtname");
			By capacityTxtBox = By.id("ContentPlaceHolder1_txtcapacity");
			By submitButton = By.id("ContentPlaceHolder1_btnsubmit");
			By msgDisplayed = By.id("ContentPlaceHolder1_lblmsg");
			
			
			//  Step-1] Open Add Zone page.
			wait.until(ExpectedConditions.elementToBeClickable(addZoneBtn));
			driver.findElement(addZoneBtn).click();
			wait.until(ExpectedConditions.visibilityOfElementLocated(zoneNameTxtBox));
			File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_05_Warehouse_Zone_Section (2).xls");
			FileInputStream fin = new FileInputStream(src);
			wb = new HSSFWorkbook(fin);
			sheet = wb.getSheet("Zone");
			
			
			for(int i = 1; i<=sheet.getLastRowNum(); i++)
			{
				try 
				{
					
					// Step-2] Import data for Zone Name.
					cell = sheet.getRow(i).getCell(1);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					driver.findElement(zoneNameTxtBox).sendKeys(cell.getStringCellValue());
					
					// Step-3] Import data for capacity.
					cell = sheet.getRow(i).getCell(2);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					driver.findElement(capacityTxtBox).sendKeys(cell.getStringCellValue());
					
					// Step-4] Click on submit button.
					driver.findElement(submitButton).click();
					
					// Step-5] Accept alert.
					Alert alert = driver.switchTo().alert();
					alert.accept();
					
					wait.until(ExpectedConditions.visibilityOfElementLocated(msgDisplayed));
					// Step-6] Set message in the excel.
					String message = driver.findElement(msgDisplayed).getText();
					System.out.println(message);
					extent.log(LogStatus.INFO, message);
					sheet.getRow(i).createCell(3).setCellValue(message);
					FileOutputStream fout = new FileOutputStream(src);
					wb.write(fout);
					fout.close();
					// Reload URL.
					driver.get(driver.getCurrentUrl());
				} 
				catch (Exception e) 
				{
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				
			
			}
	     }
	
	
	/*
	 * Step-1] Open Add Zone page.
	 * Step-2] Import data for Zone drop down.
	 * Step-3] Import data for Capacity.
	 * Step-4] Click on submit button.
	 * Step-5] Accept alert.
	 * Step-6] Set message in the excel.
	 */
	@Test(priority = 2, groups = "test")
	public void AddSection() throws IOException
	{
		// Initialize IDs.
		By addSectionBtn = By.id("ContentPlaceHolder1_btnCreate");
		By zoneDD = By.id("ContentPlaceHolder1_ddlZone");
		By sectionName = By.id("ContentPlaceHolder1_txtname");
		By sectionCode = By.id("ContentPlaceHolder1_txtcode");
		By submitButton = By.id("ContentPlaceHolder1_btnsubmit");
		
		// Step-1] Open Add Section page.
		Ops.warehouse_OpenViewSection();
		wait.until(ExpectedConditions.elementToBeClickable(addSectionBtn));
		driver.findElement(addSectionBtn).click();
		wait.until(ExpectedConditions.elementToBeClickable(submitButton));
	    
		// Import Excel.
		File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_05_Warehouse_Zone_Section (2).xls");
		FileInputStream fin = new FileInputStream(src);
		wb = new HSSFWorkbook(fin);
	    
	}

}
