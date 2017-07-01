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

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
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

/*************************
 * @author Vivek Purohit *  
 *                       *
 *************************
 */

public class TC_01_GRN_Process 
{
	WebDriver driver;
	WebDriverWait wait;
	ExtentReports report;
	ExtentTest extent;
	POM_Login login;
	POM_Operations Ops;
	File src; 
	FileInputStream file;	
    HSSFWorkbook wb;
    HSSFSheet sheet;
    HSSFCell cell;
    JavascriptExecutor jse;
    
    
	
	
	//Initialize  Generalized IDs.
	By xpathIcon = By.xpath(".//*[@id='Li29']/a/div[1]/img");
	By loadingimage = By.id("loading-image");
	
	@BeforeClass(alwaysRun = true)
	public void TestSetup() 
	{
		driver = new FirefoxDriver();
		login = new POM_Login(driver);
		Ops = new POM_Operations(driver);
		wait = new WebDriverWait(driver,30);
		// Open URL. 
		login.Openurl();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		report = new ExtentReports("C:\\Users\\Admin\\Desktop\\Videcon Test Report Extent\\TC_01_GRN_Process.html", true);
	    extent = report.startTest("TC 01 GRN Process", "To check the working of GRN Process.");
	    jse = (JavascriptExecutor) driver;
	    
	    //JavascriptExecutor jse2 = (JavascriptExecutor) driver;
		
	}
	
	/*
	 * Create ASR (Advanced Stock Request).
	 * STEPS :
	 * Step-1] Login using C head.
	 * Step-2] Open Create ASR. Purchase --> ASR --> Create ASR.
	 * Step-3] Enter product name in the field box.
	 * Step-4] Click on the result displayed.
	 * Step-5] Enter the product quantity.
	 * Step-6] Click on submit button.
	 * Step-7] Accept alert displayed for submit prompt.
	 * Step-8] Retrieve the text displayed on the alert.
	 * */ 
	@Test(priority = 0, groups = "test")
	public void CreateASR() throws IOException, InterruptedException
	{
		try {
			//Specify IDs of the Create ASR page.
			By submitButton = By.id("ContentPlaceHolder1_btnsubmit");
			By prdctNameTxtBox = By.id("ContentPlaceHolder1_SchemeWiseSearch1_txtSchemeSearch");
			//By prdctCodeTxtBox = By.id("ContentPlaceHolder1_SchemeWiseSearch1_txtSchemeCodeSearch");
			
			By productNameOption = By.id("ContentPlaceHolder1_SchemeWiseSearch1_gvScheme_lnkSchemeName_0");
			//By productCodeOption = By.id("ContentPlaceHolder1_SchemeWiseSearch1_gvScheme_lblSchemeCode_0");
			By quantityTxtBox = By.id("ContentPlaceHolder1_Grdscheme_txtquantity_0");
			
			
			// Step-1] Login using C head.
			login.C_Head_Login();
			extent.log(LogStatus.INFO, "LoggedIn Successfully.");
			driver.findElement(xpathIcon).click();
			wait.until(ExpectedConditions.elementToBeClickable(By.linkText("PURCHASE")));
			
			// Step-2] Open Create ASR. Purchase --> ASR --> Create ASR.
			driver.findElement(By.linkText("PURCHASE")).click();
			driver.findElement(By.linkText("ASR")).click();
			driver.findElement(By.linkText("CREATE ASR")).click();
			wait.until(ExpectedConditions.elementToBeClickable(submitButton));
			extent.log(LogStatus.INFO, "Create ASR page opened.");
			// Step-3] Enter product name in the field box (executing through Apache POI)
			src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_01_GRN Process.xls");
            file = new FileInputStream(src);
			wb = new HSSFWorkbook(file);
			sheet = wb.getSheet("Create ASR");
			for(int i = 1; i<=sheet.getLastRowNum(); i++)
			{
				cell = sheet.getRow(i).getCell(1);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(prdctNameTxtBox).sendKeys(cell.getStringCellValue());
			    wait.until(ExpectedConditions.invisibilityOfElementLocated(loadingimage));
			    
			    //Step-4] Click on the result displayed.
			    driver.findElement(productNameOption).click();
			    wait.until(ExpectedConditions.invisibilityOfElementLocated(loadingimage));
			    System.out.println("Reached executor");
			    jse.executeScript("window.scrollBy(0,1000)", "");
			    System.out.println("executor executed");
			    //jse.executeScript("window.scrollBy(0,250)", "");
			    
			    // Step-5] Enter the product quantity.
			    wait.until(ExpectedConditions.visibilityOfElementLocated(quantityTxtBox));
			    cell = sheet.getRow(i).getCell(2);
			    cell.setCellType(Cell.CELL_TYPE_STRING);
			    driver.findElement(quantityTxtBox).sendKeys(cell.getStringCellValue());
			    
				driver.findElement(quantityTxtBox).sendKeys(Keys.ENTER);
				wait.until(ExpectedConditions.invisibilityOfElementLocated(loadingimage));
				//driver.findElement(By.id("ContentPlaceHolder1_btncancel")).click();
					
				//Step-6] Click on submit button.
				driver.findElement(submitButton).click();
				System.out.println("clicked on submit button");
				//Step-7] Accept alert displayed for submit prompt.
			    Alert alert = driver.switchTo().alert();
				alert.accept();
				System.out.println("Alert 1 accepted");
				Thread.sleep(1000);
			    // wait.until(ExpectedConditions.invisibilityOfElementLocated(loadingimage));
  
				//Step-8] Retrieve the text displayed on the alert.
					
				Alert alert1 = driver.switchTo().alert();
				String msgDisplayed = alert1.getText();
				alert1.accept();
				System.out.println("Alert 2 accepted");
				System.out.println("message displayed = "+msgDisplayed);
				
				//Step-9] Write the message displayed in the excel.
				sheet.getRow(i).getCell(3).setCellValue(msgDisplayed);
				FileOutputStream fout = new FileOutputStream(src);
				wb.write(fout);
				fout.close();	
				extent.log(LogStatus.INFO, "Message displayed after creating ASR = " +msgDisplayed);
			}
		} 
		catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
			File sc = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(sc, new File("C:\\Users\\Admin\\Desktop\\Videocon error screen shot\\E1.png"));
			extent.addScreenCapture("C:\\Users\\Admin\\Desktop\\Videocon error screen shot\\E1.png");
		}		
	}
	
	
	/*
	 * Steps for create PO.
	 * 
	 * Step-1]  Open Create PO.
	 * Step-2]  Click on search button.
	 * Step-3] Click on the CreatePO button.
	 * Step-4] Enter Expected delivery date.
	 * Step-5] Select payment terms.
	 * Step-6] Enter Remarks.
	 * Step-7] Click submit button.
	 * Step-8] Accept alert.
	 * 
	 */
	
	
     @Test(priority = 1, groups = "test1")
     public void CreatePO() throws IOException
     {
    	 try 
    	 {
			// Initialize Id of the create PO page.
			 By srchButton = By.id("ContentPlaceHolder1_btnserch");
			 By chckBox = By.id("ContentPlaceHolder1_gvPOReq_chkOne_0");
			 By createPOBtn = By.id("ContentPlaceHolder1_btnCreatePO");
			 By warehouseDD = By.id("ContentPlaceHolder1_drpVWare");
			 By expectedDate = By.id("ContentPlaceHolder1_txtWHExpectedDelDate");
			 By paymentTermsDD = By.id("ContentPlaceHolder1_drpPaymentTerms");
			 By remarksTxtBox = By.id("ContentPlaceHolder1_txtRemarks");
			 By submitButton = By.id("ContentPlaceHolder1_btnSubmit");
			
			 
			 login.C_Head_Login();
			extent.log(LogStatus.INFO, "LoggedIn Successfully.");
			driver.findElement(xpathIcon).click();
			wait.until(ExpectedConditions.elementToBeClickable(By.linkText("PURCHASE")));
			 
				
				//Step-1]  Open Create PO.
			 driver.findElement(By.linkText("PURCHASE")).click();
			 driver.findElement(By.linkText("PURCHASE ORDER")).click();
			 driver.findElement(By.linkText("CREATE PO")).click();
			 wait.until(ExpectedConditions.elementToBeClickable(srchButton));
			 
			 src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_01_GRN Process.xls");
	        file = new FileInputStream(src);
			wb = new HSSFWorkbook(file);
			sheet = wb.getSheet("Create ASR");
			 for(int i = 1; i<=sheet.getLastRowNum(); i++)
			 { 
				//Step-2]  Click on search button.
				 driver.findElement(srchButton).click();
				 Ops.WaitforLoadingImageToDisappaer();
				 
				 //Step-3] Click on the check box.
				 driver.findElement(chckBox).click();
				 Ops.WaitforLoadingImageToDisappaer();
				 
				 //Step-3] Click on the CreatePO button.
				 driver.findElement(createPOBtn).click();
				 Ops.WaitforLoadingImageToDisappaer();
				 
				//Step-3] Select warehouse.
				 Select warehouse = new Select(driver.findElement(warehouseDD));
				 cell = sheet.getRow(i).getCell(4);
				 cell.setCellType(Cell.CELL_TYPE_STRING);
				 warehouse.selectByVisibleText(cell.getStringCellValue());
				 Ops.WaitforLoadingImageToDisappaer();
				 
				//Step-4] Enter Expected delivery date.
				 cell = sheet.getRow(i).getCell(5);
				 cell.setCellType(Cell.CELL_TYPE_STRING);
				 driver.findElement(expectedDate).sendKeys(cell.getStringCellValue());
				 
				//Step-5] Select payment terms.
				 Select paymentTerms = new Select(driver.findElement(paymentTermsDD));
				 paymentTerms.selectByValue("12");
				 
				//Step-6] Enter remarks.
				 driver.findElement(remarksTxtBox).sendKeys("Test");
				 
				 //Step-7] Click submit button.
				 driver.findElement(submitButton).click();
				 
				//Step-7] Accept alert.
				 Alert alert = driver.switchTo().alert();
				 alert.accept();
				 
				//Step-8] Wait for message to be displayed and write in the excel.
				 Ops.WaitforLoadingImageToDisappaer();
				 wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_lblmsg")));
				 String msgDisplayed = driver.findElement(By.id("ContentPlaceHolder1_lblmsg")).getText();
				 sheet.getRow(i).getCell(6).setCellValue(msgDisplayed);
				 FileOutputStream fout = new FileOutputStream(src);
				 wb.write(fout);
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
