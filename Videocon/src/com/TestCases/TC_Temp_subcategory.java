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
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.POM.Pages.POM_Login;
import com.POM.Pages.POM_Operations;

/**
 * @author Admin
 *
 */
public class TC_Temp_subcategory 
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
	public void CreateSubcategory() throws IOException, InterruptedException
	{
		try {
			wait.until(ExpectedConditions.elementToBeClickable(By.linkText("PRODUCT")));
			 driver.findElement(By.linkText("PRODUCT")).click();
			    driver.findElement(By.linkText("PRODUCT SETTING")).click();
			    driver.findElement(By.linkText("VIEW SUB - CATEGORY")).click();
			    wait.until(ExpectedConditions.elementToBeClickable(By.id("ContentPlaceHolder1_schememaster_btnadd")));
			    driver.findElement(By.id("ContentPlaceHolder1_schememaster_btnadd")).click();
			    wait.until(ExpectedConditions.elementToBeClickable(By.id("ContentPlaceHolder1_ucsubcategory_btnsubmit")));
			    File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\Temp_CreateSubcategory.xls");
				FileInputStream fin = new FileInputStream(src);
				workbook = new HSSFWorkbook(fin);
				sheet = workbook.getSheetAt(0);
				for(int i = 1; i <= sheet.getLastRowNum(); i++)
				{
					wait.until(ExpectedConditions.elementToBeClickable(By.id("ContentPlaceHolder1_ucsubcategory_btnsubmit")));
					cell = sheet.getRow(i).getCell(1);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					Select categoryDD = new Select(driver.findElement(By.id("ContentPlaceHolder1_ucsubcategory_ddlSchemeCategory")));
					categoryDD.selectByVisibleText(cell.getStringCellValue());
					
					cell = sheet.getRow(i).getCell(2);
					cell.setCellType(Cell.CELL_TYPE_STRING);
				    driver.findElement(By.id("ContentPlaceHolder1_ucsubcategory_txtSubCatName")).sendKeys(cell.getStringCellValue());
				    
				    cell = sheet.getRow(i).getCell(3);
					cell.setCellType(Cell.CELL_TYPE_STRING);
				    driver.findElement(By.id("ContentPlaceHolder1_ucsubcategory_txtSubCatCode")).sendKeys(cell.getStringCellValue());
				    new Select(driver.findElement(By.id("ContentPlaceHolder1_ucsubcategory_GridView1_ddlstate_0"))).selectByVisibleText("Delhi");
				    Thread.sleep(1000);
				    cell = sheet.getRow(i).getCell(4);
					cell.setCellType(Cell.CELL_TYPE_STRING);
				    driver.findElement(By.id("ContentPlaceHolder1_ucsubcategory_GridView1_txtvat_0")).sendKeys(cell.getStringCellValue());
				    
				    cell = sheet.getRow(i).getCell(5);
					cell.setCellType(Cell.CELL_TYPE_STRING);
				    driver.findElement(By.id("ContentPlaceHolder1_ucsubcategory_GridView1_txtcst_0")).sendKeys(cell.getStringCellValue());
				    
				    driver.findElement(By.id("ContentPlaceHolder1_ucsubcategory_btnsubmit")).click();
				    
				    Alert alert = driver.switchTo().alert();
				    alert.accept();
				    Thread.sleep(1500);
				    wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_schememaster_lblmsg")));
				    String msgDisplayed = driver.findElement(By.id("ContentPlaceHolder1_schememaster_lblmsg")).getText();
				    System.out.println(msgDisplayed);
				    sheet.getRow(i).createCell(6).setCellValue(msgDisplayed);
				    FileOutputStream fout = new FileOutputStream(src);
				    workbook.write(fout);
				    fout.close();
				    driver.get(driver.getCurrentUrl());
				    wait.until(ExpectedConditions.elementToBeClickable(By.id("ContentPlaceHolder1_schememaster_btnadd")));
				    driver.findElement(By.id("ContentPlaceHolder1_schememaster_btnadd")).click();
				    wait.until(ExpectedConditions.elementToBeClickable(By.id("ContentPlaceHolder1_ucsubcategory_btnsubmit")));
				
				
				}
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
