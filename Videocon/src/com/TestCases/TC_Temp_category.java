/**
 * 
 */
package com.TestCases;

import java.io.File;
import java.io.FileInputStream;
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

/**
 * @author Admin
 *
 */
public class TC_Temp_category
{
	WebDriver driver;
	WebDriverWait wait;
	POM_Login login;
	HSSFWorkbook workbook;
	HSSFSheet sheet;
	HSSFCell cell;
	POM_Operations ops;
	
	
	@BeforeClass(alwaysRun = true)
	public void TestSetup()
	{
		driver = new FirefoxDriver();
		login = new POM_Login(driver);
		ops = new POM_Operations(driver);
		driver.get("http://192.168.1.201:7110");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		wait = new WebDriverWait(driver,30);
	}
	
	@Test(priority = 0)
	public void Login()
	{
		login.AdminXLogin();
	}
	
	@Test(priority = 1)
	public void TestCategory() throws IOException
	{
		try {
			wait.until(ExpectedConditions.elementToBeClickable(By.linkText("PRODUCT")));
			driver.findElement(By.linkText("PRODUCT")).click();
			driver.findElement(By.linkText("PRODUCT SETTING")).click();
			driver.findElement(By.linkText("ADD/VIEW CATEGORY")).click();
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_uccategory_btnsubmit")));
			File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\Temp_Create_Category.xls");
			FileInputStream fin = new FileInputStream(src);
			workbook = new HSSFWorkbook(fin);
			sheet = workbook.getSheetAt(0);
			for(int i = 1; i<=sheet.getLastRowNum(); i++)
			{
				cell = sheet.getRow(i).getCell(1);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(By.id("ContentPlaceHolder1_uccategory_txtschemecatname")).sendKeys(cell.getStringCellValue());
				driver.findElement(By.id("ContentPlaceHolder1_uccategory_btnsubmit")).click();
				Alert alert = driver.switchTo().alert();
				alert.accept();
				ops.WaitforLoadingImageToDisappaer();
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_uccategory_lblmsg")));
				String msgDisplayed = driver.findElement(By.id("ContentPlaceHolder1_uccategory_lblmsg")).getText();
				System.out.println(msgDisplayed);
				sheet.getRow(i).createCell(2).setCellValue(msgDisplayed);
				FileOutputStream fout = new FileOutputStream(src);
				workbook.write(fout);
				fout.close();
			}
		} 
		catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
