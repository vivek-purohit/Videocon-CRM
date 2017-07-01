/**
 * 
 */
package com.TestCases;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
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
import com.thoughtworks.selenium.webdriven.commands.IsAlertPresent;

/**
 * @author Admin
 *
 */
public class TC_04_AdmnX_Add_Scheme 
{
	WebDriver driver;
	ExtentReports reports;
	ExtentTest extent;
	WebDriverWait wait;
	POM_Login login;
	POM_Operations Ops;
	HSSFWorkbook wb;
	HSSFSheet sheet;
	HSSFCell cell;
	
	//Initialize IDs
	
	By addButton = By.id("ContentPlaceHolder1_schememaster_btnAdd");
	By categoryDD = By.id("ContentPlaceHolder1_schememaster_ddlcat");
	By subCategoryDD = By.id("ContentPlaceHolder1_schememaster_ddlsubcat");
	By name = By.id("ContentPlaceHolder1_schememaster_txtSchemeName");
	By code = By.id("ContentPlaceHolder1_schememaster_txtschemecodeValue");
	By startDate = By.id("ContentPlaceHolder1_schememaster_txtstart");
	By endDate = By.id("ContentPlaceHolder1_schememaster_txtend");
	By addProductGroup = By.id("ContentPlaceHolder1_schememaster_btnNewRow");
	//By productGroupDD = By.id("ContentPlaceHolder1_schememaster_Grdscheme_ddlScheme_0");
	By dimensionLength = By.id("ContentPlaceHolder1_schememaster_txtleng");
	By dimensionBreadth = By.id("ContentPlaceHolder1_schememaster_txtbread");
	By dimensionHeight = By.id("ContentPlaceHolder1_schememaster_Txtheig");
	By multiplyingFactor = By.id("ContentPlaceHolder1_schememaster_txtmultiply");
	By dividingFactor = By.id("ContentPlaceHolder1_schememaster_txtdivide");
	By weightTxtBox = By.id("ContentPlaceHolder1_schememaster_txtnetweight");
	By weightUnitDD = By.id("ContentPlaceHolder1_schememaster_ddlweightunit");
	By delChargesDD = By.id("ContentPlaceHolder1_schememaster_drpdelcharge");
	By loyaltyPoints = By.id("ContentPlaceHolder1_schememaster_txtloyalty");
	By submitButton = By.id("ContentPlaceHolder1_schememaster_btnsubmit");
	
	
	
	
	@BeforeClass(alwaysRun = true)
	public void TestSetup()
	{
		driver = new FirefoxDriver();
		login = new POM_Login(driver);
		Ops = new POM_Operations(driver);
		// Open URL.
		login.Openurl();
		reports = new ExtentReports("C:\\Users\\Admin\\Desktop\\Videcon Test Report Extent\\TC_04_AdmnX_Add_Scheme.html", true);
	    extent = reports.startTest("TC_04_AdmnX_Add_Scheme", "To check the functionality of adding scheme");
	    wait = new WebDriverWait(driver,30);
	    
	    
	    driver.manage().window().maximize();
	    driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	    
	}
	/*
	 * Step-1] Login using adminX credentials.
	 * Step-2] Open Add scheme page.
	 * Step-3] Import Excel.
	 */
	
	@Test(priority = 0, groups = "test")
	public void TestLogin()
	{
		// Step-1] Login using adminX credentials.
	    login.AdminXLogin();
	    extent.log(LogStatus.INFO, "LoggedIn successfully.");
	    wait.until(ExpectedConditions.elementToBeClickable(By.linkText("PRODUCT")));
	   // Step-2] Open Add scheme page.
	    Ops.OpenAddScheme();
	    
	}
	
	/*
	 * Step-1] Import data for Category.
	 * Step-2] Import data for Sub Category.
	 * Step-3] Import data for Name.
	 * Step-4] Input data for Start date.
	 * Step-5] Input data for End date.
	 * Step-6] Import  data for Product Group.
	 * Step-7] Import  data for CP.
	 * Step-8] Import  data for Del. Charge %.
	 * Step-9] Import  data for Discount %.
	 * Step-10] Iterate the process for the data retrieved in steps 7 to 9.
	 * Step-11] Import data for length.
	 * Step-12] Import data for breadth.
	 * Step-13] Import data for height.
	 * Step-14] Import data for Multiplying factor.
	 * Step-15] Import data for Dividing factor.
	 * Step-16] Import data for Weight.
	 * Step-17] Import data for Weight unit.
	 * Step-18] Import data for Delivery charges.
	 * Step-19] Import data for loyalty points.
	 * Step-20] Import data for mode.
	 * Step-21] Import data for  mode priority.
	 * Step-22] Click on submit button.
	 * 
	 */
	
	@Test(priority = 1, groups = "test")
	public void AddSchemeProcess() throws IOException, InterruptedException
	{
		File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_04_Add_Scheme.xls");
		FileInputStream fin = new FileInputStream(src);
		wb = new HSSFWorkbook(fin);
		sheet = wb.getSheetAt(0);
		
		for(int i = 1; i<=sheet.getLastRowNum(); i++)
		{
			try 
			{
				wait.until(ExpectedConditions.elementToBeClickable(addButton));
				driver.findElement(addButton).click();
				wait.until(ExpectedConditions.visibilityOfElementLocated(categoryDD));
				extent.log(LogStatus.INFO, "Scheme master page opened.");
				
				//Step-1] Import data for Category.
				cell = sheet.getRow(i).getCell(1);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select category = new Select(driver.findElement(categoryDD));
				category.selectByVisibleText(cell.getStringCellValue());
				
				//Step-2] Import data for Sub Category.
				cell = sheet.getRow(i).getCell(2);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select subCategory = new Select(driver.findElement(subCategoryDD));
				subCategory.selectByVisibleText(cell.getStringCellValue());
				
				//Step-3] Import data for Name.
				cell = sheet.getRow(i).getCell(3);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(name).sendKeys(cell.getStringCellValue());
				driver.findElement(name).sendKeys(Keys.ENTER);
				Thread.sleep(3500);
				//Step-3] Import data for Code.
				cell = sheet.getRow(i).getCell(4);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(code).sendKeys(cell.getStringCellValue());
				
				//Step-4] Input for Start date.
				driver.findElement(startDate).click();
				driver.findElement(By.xpath(".//*[@id='ContentPlaceHolder1_schememaster_startdate_today']")).click();
							
				// Step-5] Input data for End date.
				driver.findElement(endDate).sendKeys("2016-12-22");
				driver.findElement(endDate).sendKeys(Keys.ENTER);
				
				// Step-6] Import  data for Product group and Iterate.
				cell = sheet.getRow(i).getCell(5);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				String productGroupString = cell.getStringCellValue();
				List<String> productGrpArray = Arrays.asList(productGroupString.split(","));
				
				// Step-7] Import  data for CP.
				cell = sheet.getRow(i).getCell(6);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				String cpString = cell.getStringCellValue();
				List<String> cpArray = Arrays.asList(cpString.split(","));
				
				// Step-8] Import  data for Del charge %.
				cell = sheet.getRow(i).getCell(7);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				String delChargeString = cell.getStringCellValue();
				List<String> delChargeArray = Arrays.asList(delChargeString.split(","));
				
				// Step-9] Import  data for Discount %.
				cell = sheet.getRow(i).getCell(8);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				String discountString = cell.getStringCellValue();
				List<String> discountArray = Arrays.asList(discountString.split(","));
				
				
				// Step-10] Iterate the process for the data retrieved in steps 7 to 9.
				for(int j = 0; j<productGrpArray.size(); j++)
				{
					// Create ID for product group drop down.
					String firstPartIDprdctGrp = "ContentPlaceHolder1_schememaster_Grdscheme_ddlScheme_";
					String finalPartIDprdctGrp = firstPartIDprdctGrp+j;
					
					// Create ID for CP.
					String firstPartIdCP = "ContentPlaceHolder1_schememaster_Grdscheme_txttempcp_price_";
					String finalpartIdCP = firstPartIdCP+j;
					
					// Create ID for DelCharge %.
					String firstPartIdDelCharge = "ContentPlaceHolder1_schememaster_Grdscheme_txttemp_delcharge_";
					String finalpartIdDelCharge = firstPartIdDelCharge+j;
					
					// Create ID for Discount %.
					String firstPartIdDiscount = "ContentPlaceHolder1_schememaster_Grdscheme_txttemp_discount_";
					String finalpartIdDiscount = firstPartIdDiscount+j;
					
					
					
					// Select product group drop down.
					Select productGroup = new Select(driver.findElement(By.id(finalPartIDprdctGrp)));
					productGroup.selectByVisibleText(productGrpArray.get(j));
					Thread.sleep(3500);
					// Write data for CP.
					driver.findElement(By.id(finalpartIdCP)).sendKeys(cpArray.get(j));
					driver.findElement(By.id(finalpartIdCP)).sendKeys(Keys.ENTER);
					Thread.sleep(3500);
					// Write data for Del charge %.
					driver.findElement(By.id(finalpartIdDelCharge)).sendKeys(delChargeArray.get(j));
					
					// Write data for Discount %.
					driver.findElement(By.id(finalpartIdDiscount)).sendKeys(discountArray.get(j));
					
					// Click on "Add Product Group" if more than 1 Product group exists.
					int rowCount = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_schememaster_Grdscheme']/tbody/tr")).size();
				    
					if(productGrpArray.size() > 1 && rowCount <= productGrpArray.size())
					{
						driver.findElement(addProductGroup).click();
					}
				}
				
				// Step-11] Import data for length.
				cell = sheet.getRow(i).getCell(9);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(dimensionLength).sendKeys(cell.getStringCellValue());
				driver.findElement(dimensionLength).sendKeys(Keys.ENTER);
				Thread.sleep(3500);
				
				// Step-12] Import data for breadth.
				cell = sheet.getRow(i).getCell(10);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(dimensionBreadth).sendKeys(cell.getStringCellValue());
				driver.findElement(dimensionBreadth).sendKeys(Keys.ENTER);
				Thread.sleep(3500);
				
				// Step-13] Import data for height.
				cell = sheet.getRow(i).getCell(11);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(dimensionHeight).sendKeys(cell.getStringCellValue());
				driver.findElement(dimensionHeight).sendKeys(Keys.ENTER);
				Thread.sleep(3500);
				
				// Step-14] Import data for Multiplying factor.
				cell = sheet.getRow(i).getCell(12);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(multiplyingFactor).clear();
				driver.findElement(multiplyingFactor).sendKeys(cell.getStringCellValue());
				driver.findElement(multiplyingFactor).sendKeys(Keys.ENTER);
				Thread.sleep(3500);
				
				// Step-14] Import data for Dividing factor.
				cell = sheet.getRow(i).getCell(13);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(dividingFactor).clear();
				driver.findElement(dividingFactor).sendKeys(cell.getStringCellValue());
				driver.findElement(dividingFactor).sendKeys(Keys.ENTER);
				Thread.sleep(3500);
				
				// Step-16] Import data for Weight.
				cell = sheet.getRow(i).getCell(14);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(weightTxtBox).sendKeys(cell.getStringCellValue());
				
				 // Step-17] Import data for Weight unit.
				cell = sheet.getRow(i).getCell(15);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select weightUnit = new Select(driver.findElement(weightUnitDD));
				weightUnit.selectByVisibleText(cell.getStringCellValue());
				
				 // Step-18] Import data for Delivery charges.
				cell = sheet.getRow(i).getCell(16);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select deliveryCharges = new Select(driver.findElement(delChargesDD));
				deliveryCharges.selectByVisibleText(cell.getStringCellValue());
				
				// Step-19] Import data for Loyalty points.
				cell = sheet.getRow(i).getCell(17);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(loyaltyPoints).sendKeys(cell.getStringCellValue());
				
				// Step-20] Import data for mode.
				cell = sheet.getRow(i).getCell(18);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				String mode = cell.getStringCellValue();
				
				// Check for the retreived data and click accordingly.
				if(mode.equalsIgnoreCase("AIR"))
				{
					driver.findElement(By.id("ContentPlaceHolder1_schememaster_rdmode_0")).click();
					Thread.sleep(3500);
				}
				if(mode.equalsIgnoreCase("SURFACE"))
				{
					driver.findElement(By.id("ContentPlaceHolder1_schememaster_rdmode_1")).click();
					Thread.sleep(3500);
				}
				if(mode.equalsIgnoreCase("BOTH"))
				{
					driver.findElement(By.id("ContentPlaceHolder1_schememaster_rdmode_2")).click();
					Thread.sleep(3500);
					// Step-20] Import data for priority.
					cell = sheet.getRow(i).getCell(19);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_schememaster_ddlmodepri")));
					Select priority = new Select(driver.findElement(By.id("ContentPlaceHolder1_schememaster_ddlmodepri")));
					priority.selectByVisibleText(cell.getStringCellValue());
				}
				// Click on the packing box.
				driver.findElement(By.xpath(".//*[@id='ContentPlaceHolder1_schememaster_lstpacking']/option[3]")).click();
				
				
				// Step-22] Click on submit button.
				driver.findElement(submitButton).click();
				
				try 
				{
					Alert alert = driver.switchTo().alert();
					String alertMsg = alert.getText();
					if(alertMsg.equalsIgnoreCase("Are you sure to Submit this record?"))
					{
						alert.accept();
						wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//*[@id='ContentPlaceHolder1_schememaster_pnlwarehouse']/div/table/tbody/tr[1]/td[1]")));
						driver.findElement(By.xpath(".//*[@id='ContentPlaceHolder1_schememaster_pnlwarehouse']/div/table/tbody/tr[1]/td[1]")).sendKeys(Keys.ESCAPE);
					    wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_schememaster_lblmsg")));
					    String successMsg = driver.findElement(By.id("ContentPlaceHolder1_schememaster_lblmsg")).getText();
					    System.out.println(successMsg);
					    extent.log(LogStatus.PASS, successMsg);
					    sheet.getRow(i).createCell(21).setCellValue(alertMsg);
						FileOutputStream fout = new FileOutputStream(src);
						wb.write(fout);
						fout.close();
					}
					else
					{
						extent.log(LogStatus.ERROR, "Error message" +alertMsg);
						sheet.getRow(i).createCell(20).setCellValue(alertMsg);
						FileOutputStream fout = new FileOutputStream(src);
						wb.write(fout);
						fout.close();
						try 
						{
							alert.accept();
							Alert alert1 = driver.switchTo().alert();
							alert1.accept();
						} 
						catch (Exception e) 
						{
							// TODO Auto-generated catch block
							System.out.println("Second alert is not present");
						}
					}
				} 
				catch (Exception e)
				{
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				driver.get(driver.getCurrentUrl());
			} 
			catch (Exception e)
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			
		}
		reports.endTest(extent);
		reports.flush();
	} 

}
