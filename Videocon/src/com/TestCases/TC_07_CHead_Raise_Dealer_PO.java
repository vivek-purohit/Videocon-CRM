/**
 * 
 */
package com.TestCases;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
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
public class TC_07_CHead_Raise_Dealer_PO 
{
	WebDriver driver;
	WebDriverWait wait;
	ExtentReports reports;
	ExtentTest extent;
	POM_Login login;
	POM_Operations ops;
	JavascriptExecutor jse;
	// Excel Import.
	HSSFWorkbook workbook;
    HSSFSheet sheet;
    HSSFCell cell;
	
	
	@BeforeClass(alwaysRun = true)
	public void TestSetup()
	{
		driver = new FirefoxDriver();
		login = new POM_Login(driver);
		ops = new POM_Operations(driver);
		login.Openurl();
		driver.manage().window().maximize();
		wait = new WebDriverWait(driver,30);
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		reports = new ExtentReports("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_07_CHead_Raise_Dealer_PO.xls", true);
	    extent = reports.startTest("TC_07_CHead_Raise_Dealer_PO", "To check the functionality of raising dealer PO");
	    jse = (JavascriptExecutor)driver;
	    
	    
	    
	}
	
	@Test(priority = 0, groups = "test")
	public void TestLogin()
	{
		login.C_Head_Login();
		extent.log(LogStatus.INFO, "Logged In successfully !!");
		wait.until(ExpectedConditions.elementToBeClickable(By.linkText("APPROVAL AND PURCHASE")));
		
	}
	
	/*
	 * Step -1 ]Open Raise dealer PO page.
	 * Step -2 ]Select type like Warehouse or Vendor.
	 * Step -3 ]Select name of Warehouse/ Vendor.
	 * Step -4 ]Select Dealer.
	 * Step -5 ]Select Expected Date.
	 * Step -6 ]Enter name of Scheme.
	 * Step -7 ]Select the scheme from the result displayed in the table.
	 * Step -8 ]Enter PO quantity.
	 * Step -9 ]Select Payment Terms.
	 * Step -10 ]Enter Remarks.
	 * Step -11 ]Click on submit button.
	 */
	
	@Test(priority = 1, groups = "test")
	public void RaiseDealerPO() throws IOException, AWTException
	{
		// Generalize ID.
		By typeDD = By.id("ContentPlaceHolder1_ddltype");
		By typeNameDD = By.id("ContentPlaceHolder1_drpVWare");
		By dealerNameDD = By.id("ContentPlaceHolder1_drpDealer");
		By deliveryDate = By.id("ContentPlaceHolder1_txtWHExpectedDelDate");
		By srchSchemeNameTxtBox = By.id("ContentPlaceHolder1_SchemeWiseSearch1_txtSchemeSearch");
		By schemeNameResult = By.id("ContentPlaceHolder1_SchemeWiseSearch1_gvScheme_lnkSchemeName_0");
		By paymentTermsDD = By.id("ContentPlaceHolder1_drpPaymentTerms");
		By remarks = By.id("ContentPlaceHolder1_txtRemarks");
		By submitButton = By.id("ContentPlaceHolder1_btnSubmit");
		
		// Initialize excel.
		File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_07_CHead_Raise_Dealer_PO.xls");
		FileInputStream fin = new FileInputStream(src);
		workbook = new HSSFWorkbook(fin);
		sheet = workbook.getSheetAt(0);
		
		//workbook = new HSSFWorkbook(fin);
		/* Note : For message check url on which message 
		 * appears since ID of both message is same 
		 * but they are appearing at different place, success message
		 * on view dealer page and error message on raise dealer po page.
		*/
		By msgDisplayed = By.id("ContentPlaceHolder1_lblmsg");
		
		extent.log(LogStatus.INFO, "Raise Dealer PO page opened successfully.");
		
		
		for(int i = 1; i<=sheet.getLastRowNum(); i++)
		{
			try 
			{
				// Step -1 ] Open Raise dealer PO page.
				ops.chead_OpenRaiseDealerPO();
				wait.until(ExpectedConditions.visibilityOfElementLocated(typeDD));
				
				// Step -2 ] Select type like Warehouse or Vendor.
				cell = sheet.getRow(i).getCell(1);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select type = new Select(driver.findElement(typeDD));
				type.selectByVisibleText(cell.getStringCellValue());
				ops.WaitforLoadingImageToDisappaer();
				
				// Step -3 ] Select name of Warehouse/ Vendor.
				cell = sheet.getRow(i).getCell(2);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select typeName = new Select(driver.findElement(typeNameDD));
				typeName.selectByVisibleText(cell.getStringCellValue());
				
				// Step -4 ] Select Dealer.
				
				cell = sheet.getRow(i).getCell(3);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select dealer = new Select(driver.findElement(dealerNameDD));
				dealer.selectByVisibleText(cell.getStringCellValue());
				
				// Step -5 ] Select Expected Date.
				cell = sheet.getRow(i).getCell(4);
				if(DateUtil.isCellDateFormatted(cell))
				{
					try 
					{
						SimpleDateFormat df = new SimpleDateFormat("dd-MM-yyyy");
						driver.findElement(deliveryDate).sendKeys(df.format(cell.getDateCellValue()));
						System.out.println(df.format(cell.getDateCellValue()));
					} 
					catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
				
				// Step -6 ] Enter name of Scheme.
				cell = sheet.getRow(i).getCell(5);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(srchSchemeNameTxtBox).sendKeys(cell.getStringCellValue());
				ops.WaitforLoadingImageToDisappaer();
				
				// Step -7 ]Select the scheme from the result displayed in the table.
				wait.until(ExpectedConditions.elementToBeClickable(schemeNameResult));
				driver.findElement(schemeNameResult).click();
				ops.WaitforLoadingImageToDisappaer();
				
				// Step -8 ] Enter PO quantity.
				// Step-8.1] Retrieve the PO quantity from excel.
				cell = sheet.getRow(i).getCell(6);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				String POstring = cell.getStringCellValue();
				List<String> POarray;
				POarray = Arrays.asList(POstring.split(","));
				int rowCount = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_gvNew']/tbody/tr")).size();
				// Step-8.2] Check the table size and iterate the process of entering PO quantity.
				for(int j = 0; j<=rowCount-3; j++)
				{
					String firstpartID = "ContentPlaceHolder1_gvNew_txtPOQuantity_";
					String finalID = firstpartID+j;
					jse.executeScript("window.scrollBy(0,1000)", "");
					Thread.sleep(2500);
					wait.until(ExpectedConditions.elementToBeClickable(By.id(finalID)));
					/*driver.findElement(By.id(finalID)).clear();
					ops.WaitforLoadingImageToDisappaer();
					jse.executeScript("window.scrollBy(0,1000)", "");
					wait.until(ExpectedConditions.elementToBeClickable(By.id(finalID)));
					driver.findElement(By.id(finalID)).sendKeys(POarray.get(j));
					driver.findElement(By.id(finalID)).sendKeys(Keys.ENTER);
					ops.WaitforLoadingImageToDisappaer();
					jse.executeScript("window.scrollBy(0,1000)", "")*/
					
					driver.findElement(By.id(finalID)).click();
					Robot robot = new Robot();
					robot.keyPress(KeyEvent.VK_RIGHT);
					robot.keyRelease(KeyEvent.VK_RIGHT);
					robot.delay(1000);
					robot.keyPress(KeyEvent.VK_BACK_SPACE);
					robot.keyRelease(KeyEvent.VK_BACK_SPACE);
					Thread.sleep(1000);
					driver.findElement(By.id(finalID)).sendKeys(POarray.get(j));
					driver.findElement(By.id(finalID)).sendKeys(Keys.ENTER); 
					ops.WaitforLoadingImageToDisappaer();

				}
				
				// Step -9 ] Select Payment Terms.
				jse.executeScript("window.scrollBy(0,1000)", "");
				cell = sheet.getRow(i).getCell(7);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select paymentTerms = new Select(driver.findElement(paymentTermsDD));
				paymentTerms.selectByVisibleText(cell.getStringCellValue());
			
				// Step -10 ] Enter Remarks.
				cell = sheet.getRow(i).getCell(8);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(remarks).sendKeys(cell.getStringCellValue());
				
				// Step -11 ] Click on submit button.
				driver.findElement(submitButton).click();
				
				// Step -12 ] Accept alert.
				Alert alert = driver.switchTo().alert();
				alert.accept();
				
				try
				{
					Alert alert1 = driver.switchTo().alert();
					String alertMessage = alert.getText();
					alert.accept();
					sheet.getRow(i).createCell(9).setCellValue("alertMessage");
					FileOutputStream fout = new FileOutputStream(src);
					workbook.write(fout);
					fout.close();
				}
				catch(Exception e)
				{
					sheet.getRow(i).createCell(9).setCellValue("No error alert displayed");
					FileOutputStream fout = new FileOutputStream(src);
					workbook.write(fout);
					fout.close();
				}
				
				// Step -12 ] Check URL and based on that decide if error or success message.
				wait.until(ExpectedConditions.visibilityOfElementLocated(msgDisplayed));
				String currentURL = driver.getCurrentUrl();
				System.out.println("Current URL is = " +currentURL);
				
				if(currentURL.equalsIgnoreCase("http://192.168.1.201:7111/DealerPO/po.aspx"))
				{
					String message = driver.findElement(msgDisplayed).getText();
					sheet.getRow(i).createCell(10).setCellValue(message);
					FileOutputStream fout = new FileOutputStream(src);
					workbook.write(fout);
					fout.close();
				}
				
				if(currentURL.equalsIgnoreCase("http://192.168.1.201:7111/DealerPO/viewpo.aspx"))
				{
					String message = driver.findElement(msgDisplayed).getText();
					sheet.getRow(i).createCell(10).setCellValue(message);
					FileOutputStream fout = new FileOutputStream(src);
					workbook.write(fout);
					fout.close();
				}
			} 
			catch (Exception e) 
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			
		}
		
		
	}

}
