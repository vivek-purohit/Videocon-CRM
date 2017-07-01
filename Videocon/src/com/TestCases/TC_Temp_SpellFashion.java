/**
 * 
 */
package com.TestCases;

import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

/**
 * @author Admin
 *
 */
public class TC_Temp_SpellFashion 
{
	WebDriver driver;
	WebDriverWait wait;
	HSSFWorkbook workbook;
	HSSFSheet sheet;
	HSSFCell cell;
	
	@BeforeClass(alwaysRun = true)
	public void TestSetup()
	{
		driver = new FirefoxDriver();
		driver.get("http://spellfashionscamessi.com/");
		driver.manage().window().maximize();
		wait = new WebDriverWait(driver,30);
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		
	}
	@Test(priority = 0, groups = "test")
	public void TestLogin()
	{
		By userNameTXtBox = By.id("txtuser");
		By PasswordTxtBox = By.id("txtpass");
		By submitBtn = By.id("btnlogin");
		
		
		wait.until(ExpectedConditions.visibilityOfElementLocated(userNameTXtBox));
		driver.findElement(userNameTXtBox).sendKeys("harshit");
		driver.findElement(PasswordTxtBox).sendKeys("123456");
		driver.findElement(submitBtn).click();
		
	}
	
	@Test(priority = 1, groups = "test")
	public void CreateOrder() throws InterruptedException
	{
		try 
		{
			//int i = 1;
			By salesRepresentativeDD = By.id("ctl00_phnew_ddlSalesRep");
			By custFullNameDD = By.id("ctl00_phnew_ddlcustomer");
			By measurementMethodDD = By.id("ctl00_phnew_ddlmeasurementmethod");
			By baseSizeDD = By.id("ctl00_phnew_ddlSize");
			By loadingImage = By.xpath(".//*[@id='ctl00_phnew_UpdatePanel1prog']/img");
			By sleeveTypeDD = By.id("ctl00_phnew_ddlsleeve");
			By sleeveType_RadioBtn = By.id("ctl00_phnew_rblcuff_0");
			By defaultDD = By.id("ctl00_phnew_ddlDef");
			By btnStyleDD = By.id("ctl00_phnew_ddlBtnStyle");
			By collarBtnDD = By.id("ctl00_phnew_ddlColBtn");
			By sloppingShoulderDD = By.id("ctl00_phnew_ddlSlopingShoulder");
			By neckTypeDD = By.id("ctl00_phnew_ddlNeck");
			By priorityOrderDD = By.id("ctl00_phnew_ddlPrior");
			By browseButton = By.id("ctl00_phnew_FileUpload1");
			By handkerchiefDD = By.id("ctl00_phnew_ddlHanderchief");
			By createOrderLnk = By.xpath(".//*[@id='ctl00_dvdealer']/div[2]/p[8]/a");
			
			
			
			for(int i = 1; i< 100; i++)
			{
				
			wait.until(ExpectedConditions.elementToBeClickable(createOrderLnk));
			driver.findElement(createOrderLnk).click();
			// Initiate the process.
			wait.until(ExpectedConditions.visibilityOfElementLocated(salesRepresentativeDD));
			
			Select salesRepresentative = new Select(driver.findElement(salesRepresentativeDD));
			salesRepresentative.selectByVisibleText("1");
			
			Select customerFullName = new Select(driver.findElement(custFullNameDD));
			customerFullName.selectByVisibleText("06 Rehbergen");
			Thread.sleep(700);
			wait.until(ExpectedConditions.invisibilityOfElementLocated(loadingImage));
			
			Select measureMethod = new Select(driver.findElement(measurementMethodDD));
			measureMethod.selectByVisibleText("Base Shirt Measurement");
			Thread.sleep(700);
			wait.until(ExpectedConditions.invisibilityOfElementLocated(loadingImage));
			
			Select Base = new Select(driver.findElement(baseSizeDD));
			Base.selectByVisibleText("200test");
			Thread.sleep(700);
			wait.until(ExpectedConditions.invisibilityOfElementLocated(loadingImage));
			
			driver.findElement(By.id("ctl00_phnew_btnNewRow")).click();
			Thread.sleep(700);
			wait.until(ExpectedConditions.invisibilityOfElementLocated(loadingImage));
			Select fabricStyle = new Select(driver.findElement(By.id("ctl00_phnew_GridView1_ctl02_ddlFabricCode")));
			fabricStyle.selectByVisibleText("A-038");
			Thread.sleep(700);
			wait.until(ExpectedConditions.invisibilityOfElementLocated(loadingImage));
			
			
			
			Select sleeveType = new Select(driver.findElement(sleeveTypeDD));
			sleeveType.selectByVisibleText("Full Sleeve");
			driver.findElement(sleeveType_RadioBtn).click();
			
			Select default1 = new Select (driver.findElement(defaultDD));
			default1.selectByVisibleText("Yes");
			
			Select btnStyle = new Select(driver.findElement(btnStyleDD));
			btnStyle.selectByVisibleText("Thick Black");
			
			Select collarBtn = new Select(driver.findElement(collarBtnDD));
			collarBtn.selectByVisibleText("1");
			
			Select slopingShoulders = new Select(driver.findElement(sloppingShoulderDD));
			slopingShoulders.selectByVisibleText("Other");
			
			Select neckType = new Select(driver.findElement(neckTypeDD));
			neckType.selectByVisibleText("Normal Neck");
			
			Select handkerchief = new Select(driver.findElement(handkerchiefDD));
			handkerchief.selectByVisibleText("Yes");
			
			Select priorityOrder = new Select(driver.findElement(priorityOrderDD));
			priorityOrder.selectByVisibleText("No");
			Thread.sleep(700);
			wait.until(ExpectedConditions.invisibilityOfElementLocated(loadingImage));
			
			
			
			// Import the file path from the excel;
			File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Spellfashion.xls");
			FileInputStream fin = new FileInputStream(src);
			workbook = new HSSFWorkbook(fin);
			sheet = workbook.getSheetAt(0);
			
			cell = sheet.getRow(i).getCell(1);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			String imageApplicable = cell.getStringCellValue();
			System.out.println("If Image applicable is = " +imageApplicable);
			
			if(imageApplicable.equalsIgnoreCase("Yes"))
			{
				driver.findElement(browseButton).click();
				cell = sheet.getRow(i).getCell(2);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				String imagePath = cell.getStringCellValue();
				System.out.println("Imaage path retrieved as = " +imagePath);
				driver.findElement(browseButton).click();
				
				Thread.sleep(2500);
				StringSelection sel = new StringSelection(imagePath);
				Toolkit.getDefaultToolkit().getSystemClipboard().setContents(sel, null);
				
				Robot robot = new Robot();
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_V);
				robot.delay(1000);
				robot.keyRelease(KeyEvent.VK_CONTROL);
				robot.keyRelease(KeyEvent.VK_V);
				robot.delay(1000);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.delay(500);
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
			}
			
			driver.findElement(By.id("ctl00_phnew_btnUpdate")).click();
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_phnew_lblmsg")));
			String msgDisplayed = driver.findElement(By.id("ctl00_phnew_lblmsg")).getText();
			System.out.println(msgDisplayed);
			FileOutputStream fout = new FileOutputStream(src);
			sheet.getRow(i).createCell(3).setCellValue(msgDisplayed);
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
