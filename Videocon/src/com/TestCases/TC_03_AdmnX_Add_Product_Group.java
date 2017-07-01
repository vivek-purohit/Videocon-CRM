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
import java.util.Arrays;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
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
public class TC_03_AdmnX_Add_Product_Group 
{
	WebDriver driver;
	WebDriverWait wait;
	ExtentReports reports;
	ExtentTest extent;
	JavascriptExecutor jse;
	POM_Login login;
	POM_Operations Ops;
	File src;
	HSSFWorkbook wb;
	HSSFSheet sheet;
	HSSFCell cell;
	
	
	//Generalize IDs
	By addButton = By.id("ContentPlaceHolder1_ucviewproductgroup_btnadd");
	By prdctGrpName = By.id("ContentPlaceHolder1_ucaddproductgroup_txtpname");
	By prdctGrpCode = By.id("ContentPlaceHolder1_ucaddproductgroup_txtShort");
	By prdctCtgryDD = By.id("ContentPlaceHolder1_ucaddproductgroup_ddlcat");
	By prdctSbCtgryDD = By.id("ContentPlaceHolder1_ucaddproductgroup_ddlsubcat");
	By vendorDD = By.id("ContentPlaceHolder1_ucaddproductgroup_ddlVendor");
	By productPriceTxtBox = By.id("ContentPlaceHolder1_ucaddproductgroup_txtprice");
	By TP = By.id("ContentPlaceHolder1_ucaddproductgroup_txtTP");
	By attributeGrpDD = By.id("ContentPlaceHolder1_ucaddproductgroup_ddlAttributeGroup");
	By barcodeApplicable = By.id("ContentPlaceHolder1_ucaddproductgroup_chkbarcApp");
	By submitButton = By.id("ContentPlaceHolder1_ucaddproductgroup_btnsubmit");
	String attribute[];
	By attributeSaveButton = By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_btnAddAttributeDetailSubmit");
	By submitButtonProductGroup = By.id("ContentPlaceHolder1_ucaddproductgroup_btnsubmit");
	By successMessage = By.id("ContentPlaceHolder1_ucviewproductgroup_lblmsg");
	By errorMessage = By.id("ContentPlaceHolder1_ucaddproductgroup_lblmsg");
	
	@BeforeClass
	public void TestSetup()
	{
		driver = new FirefoxDriver();
		login = new POM_Login(driver);
		Ops = new POM_Operations(driver);
		driver.get("http://192.168.1.201:7110");
		driver.manage().window().maximize();
		wait = new WebDriverWait(driver,30);
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		reports = new ExtentReports("C:\\Users\\Admin\\Desktop\\Videcon Test Report Extent\\TC_03_AdmnX_Add_Product_Group.html", true);
		extent = reports.startTest("TC_03_AdmnX_Add_Product_Group", "To check the functionality of add product group in AdminX");
	    jse = (JavascriptExecutor) driver;
	}
	/*Steps]
	 * Step-1] Login using adminX credentials.
	 * Step-2] Open View product group. Product --> Product Setting --> View Product Group.
	 * Step-3] Open Add product group.
	 * Step-4] Import Excel file and data.
	 * Step-5] Import data for product group name.
	 * Step-6] Import data for product group code.
	 * Step-7] Import data for product category.
	 * Step-8] Import data for product sub category.
	 * Step-9] Import data for product Vendor.
	 * Step-10] Import data for product price.
	 * Step-11] Import data for Attribute Group.
	 * Step-12] Import data for Bar code Applicable.
	 * Step-13] Check for the attributes displayed.
	 * Step-14] Import data for color attribute.
	 * Step-15] Import data for size attribute.
	 * Step-16 ]Retrieve the data and match with the array.
	 * */
	
	
	@Test(priority = 0, groups = "test")
	public void AddProductGroup() throws IOException, InterruptedException
	{
		// Step-1] Login using adminX credentials.
		login.AdminXLogin();
		wait.until(ExpectedConditions.elementToBeClickable(By.linkText("PRODUCT")));
	    extent.log(LogStatus.INFO, "Logged in successfully");
		
	    //Step-2] Open View product group. Product --> Product Setting --> View Product Group.
	    Ops.OpenProductGroup();
	    wait.until(ExpectedConditions.elementToBeClickable(addButton));
	    
	    //Step-3] Open Add product group.
	   
	    // Step-4] Import Excel file and data.
	    src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_03_Add_Product_Group.xls");
	    FileInputStream fin = new FileInputStream(src);
	    wb = new HSSFWorkbook(fin);
	    sheet = wb.getSheetAt(0);
	    
	    for(int i = 1; i<= sheet.getLastRowNum(); i++)
	    {
	    	try {
				wait.until(ExpectedConditions.elementToBeClickable(addButton));
				 driver.findElement(addButton).click();
				wait.until(ExpectedConditions.visibilityOfElementLocated(prdctGrpName));
				extent.log(LogStatus.INFO, "Add product group page opened.");
				Thread.sleep(4500);
				// Step-5] Import data for product group name.
				cell = sheet.getRow(i).getCell(1);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(prdctGrpName).sendKeys(cell.getStringCellValue());
				
				// Step-6] Import data for product group code.
				cell = sheet.getRow(i).getCell(2);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(prdctGrpCode).sendKeys(cell.getStringCellValue());
   
				// Step-7] Import data for product category.
				cell = sheet.getRow(i).getCell(3);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select prdctCtgry = new Select(driver.findElement(prdctCtgryDD));
				prdctCtgry.selectByVisibleText(cell.getStringCellValue());
   
				// Step-8] Import data for product sub category.
				Select productSubCategory = new Select(driver.findElement(prdctSbCtgryDD));
				String waitForOptions = productSubCategory.getFirstSelectedOption().getText();
				System.out.println("displayed subcategory is = " +waitForOptions);
				cell = sheet.getRow(i).getCell(4);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				productSubCategory.selectByVisibleText(cell.getStringCellValue());
				
				// Step-9] Import data for Vendor.
				Select vendor = new Select(driver.findElement(vendorDD));
				cell = sheet.getRow(i).getCell(5);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				vendor.selectByVisibleText(cell.getStringCellValue());
				
				// Step-10] Import data for product price.
				cell = sheet.getRow(i).getCell(6);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(productPriceTxtBox).sendKeys(cell.getStringCellValue());;
				
				// Step-10] Import data for TP.
				cell = sheet.getRow(i).getCell(7);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(TP).sendKeys(cell.getStringCellValue());
				
				// Step-11] Import data for Attribute Group.
				Select attributeGroup = new Select(driver.findElement(attributeGrpDD));
				cell = sheet.getRow(i).getCell(8);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				attributeGroup.selectByVisibleText(cell.getStringCellValue());
				
				//Step-12] Import data for Bar code Applicable.
				cell = sheet.getRow(i).getCell(9);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				String barcodeStatus = cell.getStringCellValue();
				if(barcodeStatus.equalsIgnoreCase("Yes"))
				{
					driver.findElement(barcodeApplicable).click();
				}
				
				// Attribute automation.
				By attributeGrpLink = By.id("ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup_lblAttribute_0");
				By attributeGrpLink2 = By.id("ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup_lblAttribute_1");
				wait.until(ExpectedConditions.elementToBeClickable(attributeGrpLink));
			   
				// Retrieve the text of attribute group.
				cell = sheet.getRow(i).getCell(8);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				String attributeGrpText = cell.getStringCellValue();
				System.out.println("Attribute group is = " +attributeGrpText);
				
				//Retreive the attribute group option from the excel.
				cell = sheet.getRow(i).getCell(10);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				String attributeOption1 = cell.getStringCellValue();
				List<String> attributeOptionArray = Arrays.asList(attributeOption1.split(","));
				wait.until(ExpectedConditions.elementToBeClickable(attributeGrpLink));
				Thread.sleep(2000);
				driver.findElement(attributeGrpLink).click();
				System.out.println("Clicked on attribute link");
				// wait for attribute div to appear.
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_body")));
				int rowCountOptions = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail']/tbody/tr")).size();
				System.out.println("Number of options in the div for option 1 = " +rowCountOptions);
				String firstXpath = ".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_lblAttributeDetailName_";
				String thirdXpath = "']";
				for(int j = 0; j<rowCountOptions-1; j++)
				{
					String finalXpath = firstXpath+j+thirdXpath;
					System.out.println("Final xpath debugging = " +finalXpath);
					String attributeOptionText = driver.findElement(By.xpath(finalXpath)).getText();
					for(int k = 0; k<attributeOptionArray.size(); k++)
					{
						if(attributeOptionText.equalsIgnoreCase(attributeOptionArray.get(k)))
						{
							String chkFirstPart = "ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_chkAttributeDetail_";
				    		String chkFinalPart = chkFirstPart+k;
				    		driver.findElement(By.id(chkFinalPart)).click();
						}
					}
					// Click on save button.
					
				}
				driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_btnAddAttributeDetailSubmit")).click();
				if(attributeGrpText.equalsIgnoreCase("COLOR + SIZE"))
				{
					Thread.sleep(4500);
					wait.until(ExpectedConditions.elementToBeClickable(attributeGrpLink2));
					driver.findElement(attributeGrpLink2).click();
					cell = sheet.getRow(i).getCell(11);
					cell.setCellType(Cell.CELL_TYPE_STRING);
					String attributeOption2 = cell.getStringCellValue();
					List<String> attributeOptionArray2 = Arrays.asList(attributeOption2.split(","));
					//driver.findElement(attributeGrpLink2).click();
					
					// wait for div to appear.
					wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_pnlAddAttributeDetail")));
					int rowCountOptions2 = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail']/tbody/tr")).size();
					System.out.println("Number of options in the div for option 1 = " +rowCountOptions2);
					String firstXpath2 = ".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_lblAttributeDetailName_";
					String thirdXpath2 = "']";
					for(int j = 0; j<rowCountOptions2-1; j++)
					{
						String finalXpath2 = firstXpath2+j+thirdXpath2;
						System.out.println("Final xpath debugging = " +finalXpath2);
						String attributeOptionText2 = driver.findElement(By.xpath(finalXpath2)).getText();
						for(int k = 0; k<attributeOptionArray2.size(); k++)
						{
							if(attributeOptionText2.equalsIgnoreCase(attributeOptionArray2.get(k)))
							{
								String chkFirstPart = "ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_chkAttributeDetail_";
					    		String chkFinalPart = chkFirstPart+k;
					    		driver.findElement(By.id(chkFinalPart)).click();
							}
						}
						// Click on save button.
						
					
					
					
				}
					driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_btnAddAttributeDetailSubmit")).click();
				
				// Attribute automation.
				
				
				//***** Pending from the excel for attribute of color.
				
   /*  //Step-13] Take the size of the attribute table.
				int rowCount = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup']/tbody/tr")).size();
				System.out.println("Number of rows in the attribute table = "+rowCount);
      
				// *********Attribute Automation *********
     
     
     
     
     
				// Step-14] Check the size of the attribute table and get the text displayed.
				if(rowCount == 2)
				{
					String attributeName = driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup_lblAttribute_0")).getText();
				    System.out.println("Attribute name displayed = "+attributeName);
					//Step-15] Check if the attribute name is Color then perform operations.
					if(attributeName.equalsIgnoreCase("Color"))
					{
						// Click on the attribute.
						driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup_lblAttribute_0")).click();
						wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_body")));
						
						//Step-16] Import data for the color attribute.
						List<String> colorArray;
						cell = sheet.getRow(i).getCell(10);
						cell.setCellType(Cell.CELL_TYPE_STRING);
						String colorAttribute = cell.getStringCellValue();
						colorArray = Arrays.asList(colorAttribute.split(","));
					 
						//Step-17] Match with the attribute options.
						int rowCountAttrColor = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail']/tbody/tr")).size();
						for(int k = 0; k<rowCountAttrColor-1; k++)
						{
							String firstPart = ".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_lblAttributeDetailName_";
						    String secondPart = "']";
						    String finalPart = firstPart+k+secondPart;
						    //Check for each entry and retrieve the text.
						    String colorOption = driver.findElement(By.xpath(finalPart)).getText();
						    for(int m = 0; m< colorArray.size(); m++)
						    {
						    	if(colorOption.equalsIgnoreCase(colorArray.get(m)))
						    	{
						    		String chkFirstPart = "ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_chkAttributeDetail_";
						    		String chkFinalPart = chkFirstPart+k;
						    		driver.findElement(By.id(chkFinalPart)).click();
						    	}
						    }
						    //Click on save button
						}  
					
						driver.findElement(attributeSaveButton).click();
					}	
						// If attribute = Size.
						if(attributeName.equalsIgnoreCase("Size"))
				    	{
							//System.out.println("debug test");
				    		// Click on the attribute.
				    		driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup_lblAttribute_0")).click();
				    		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_body")));
				    		
				    		//Step-16] Import data for the color attribute.
				    		List<String> sizeArray;
				    		cell = sheet.getRow(i).getCell(11);
							cell.setCellType(Cell.CELL_TYPE_STRING);
							String sizeAttribute = cell.getStringCellValue();
							sizeArray = Arrays.asList(sizeAttribute.split(","));
				    	 
							//Step-17] Match with the attribute options.
							int rowCountAttrSize = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail']/tbody/tr")).size();
							//System.out.println("debug 1");
							for(int k = 0; k<rowCountAttrSize-1; k++)
							{
								String firstPart = ".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_lblAttributeDetailName_";
							    String secondPart = "']";
							    String finalPart = firstPart+k+secondPart;
							    //Check for each entry and retrieve the text.
							    String sizeOption = driver.findElement(By.xpath(finalPart)).getText();
							    for(int m = 0; m< sizeArray.size(); m++)
							    {
							    	if(sizeOption.equalsIgnoreCase(sizeArray.get(m)))
							    	{
							    		String chkFirstPart = "ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_chkAttributeDetail_";
							    		String chkFinalPart = chkFirstPart+k;
							    		driver.findElement(By.id(chkFinalPart)).click();
							    	}
							    }
							    //Click on save button
							}  
				    	
							driver.findElement(attributeSaveButton).click();
				    	}
						
						// IF attribute = Color Combo
						
						if(attributeName.equalsIgnoreCase("Color Combo"))
				    	{
							//System.out.println("debug test");
				    		// Click on the attribute.
				    		driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup_lblAttribute_0")).click();
				    		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_body")));
				    		
				    		//Step-16] Import data for the color attribute.
				    		List<String> ColorComboArray;
				    		cell = sheet.getRow(i).getCell(12);
							cell.setCellType(Cell.CELL_TYPE_STRING);
							String ColorComboAttribute = cell.getStringCellValue();
							ColorComboArray = Arrays.asList(ColorComboAttribute.split(","));
				    	 
							//Step-17] Match with the attribute options.
							int rowCountAttrColorCombo = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail']/tbody/tr")).size();
							//System.out.println("debug 1");
							for(int k = 0; k<rowCountAttrColorCombo-1; k++)
							{
								String firstPart = ".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_lblAttributeDetailName_";
							    String secondPart = "']";
							    String finalPart = firstPart+k+secondPart;
							    //Check for each entry and retrieve the text.
							    String ColorComboOption = driver.findElement(By.xpath(finalPart)).getText();
							    for(int m = 0; m< ColorComboArray.size(); m++)
							    {
							    	if(ColorComboOption.equalsIgnoreCase(ColorComboArray.get(m)))
							    	{
							    		String chkFirstPart = "ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_chkAttributeDetail_";
							    		String chkFinalPart = chkFirstPart+k;
							    		driver.findElement(By.id(chkFinalPart)).click();
							    	}
							    }
							    //Click on save button
							}  
				    	
							driver.findElement(attributeSaveButton).click();
				    	}
						
						
				          // IF attribute = Default
						
						if(attributeName.equalsIgnoreCase("Default"))
				    	{
							System.out.println("Default attribute");
							wait.until(ExpectedConditions.elementToBeClickable(By.id("ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup_lblAttribute_0")));
							driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup_lblAttribute_0")).click();
				    		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_body")));
							driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_chkAttributeDetail_0")).click();
							driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_btnAddAttributeDetailSubmit")).click();
				    	}
					  
				}
				
				if(rowCount == 3)
				{
					
					String attributeName = driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup_lblAttribute_0")).getText();
					String attributeName1 = driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup_lblAttribute_1")).getText();
				    System.out.println("Attribute name displayed = "+attributeName);
				    System.out.println("Attribute name displayed = "+attributeName1);
				    //Step-15] Check if the attribute name is Color then perform operations.
					if(attributeName.equalsIgnoreCase("Color"))
					{
						// Click on the attribute.
						driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup_lblAttribute_0")).click();
						wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_body")));
						
						//Step-16] Import data for the color attribute.
						List<String> colorArray;
						cell = sheet.getRow(i).getCell(10);
						cell.setCellType(Cell.CELL_TYPE_STRING);
						String colorAttribute = cell.getStringCellValue();
						colorArray = Arrays.asList(colorAttribute.split(","));
					 
						//Step-17] Match with the attribute options.
						int rowCountAttrColor = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail']/tbody/tr")).size();
						for(int k = 0; k<rowCountAttrColor-1; k++)
						{
							String firstPart = ".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_lblAttributeDetailName_";
						    String secondPart = "']";
						    String finalPart = firstPart+k+secondPart;
						    //Check for each entry and retrieve the text.
						    String colorOption = driver.findElement(By.xpath(finalPart)).getText();
						    for(int m = 0; m< colorArray.size(); m++)
						    {
						    	if(colorOption.equalsIgnoreCase(colorArray.get(m)))
						    	{
						    		String chkFirstPart = "ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_chkAttributeDetail_";
						    		String chkFinalPart = chkFirstPart+k;
						    		driver.findElement(By.id(chkFinalPart)).click();
						    	}
						    }
						    //Click on save button
						}  
					
						driver.findElement(attributeSaveButton).click();
					}	
					
					if(attributeName1.equalsIgnoreCase("Size"))
					{
						//System.out.println("debug test");
						// Click on the attribute.
						driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_gvAttributeGroup_lblAttribute_1")).click();
						wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_body")));
						
						//Step-16] Import data for the color attribute.
						List<String> sizeArray;
						cell = sheet.getRow(i).getCell(11);
						cell.setCellType(Cell.CELL_TYPE_STRING);
						String sizeAttribute = cell.getStringCellValue();
						sizeArray = Arrays.asList(sizeAttribute.split(","));
					 
						//Step-17] Match with the attribute options.
						int rowCountAttrSize = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail']/tbody/tr")).size();
						//System.out.println("debug 1");
						for(int k = 0; k<rowCountAttrSize-1; k++)
						{
							String firstPart = ".//*[@id='ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_lblAttributeDetailName_";
						    String secondPart = "']";
						    String finalPart = firstPart+k+secondPart;
						    //Check for each entry and retrieve the text.
						    String sizeOption = driver.findElement(By.xpath(finalPart)).getText();
						    for(int m = 0; m< sizeArray.size(); m++)
						    {
						    	if(sizeOption.equalsIgnoreCase(sizeArray.get(m)))
						    	{
						    		String chkFirstPart = "ContentPlaceHolder1_ucaddproductgroup_ucProject_TabAddAttributeDetail_TabPnlAddAttributeDetail_grdAttributeDetail_chkAttributeDetail_";
						    		String chkFinalPart = chkFirstPart+k;
						    		driver.findElement(By.id(chkFinalPart)).click();
						    	}
						    }
						    //Click on save button
						}  
					
						driver.findElement(attributeSaveButton).click();
					}
					
				}*/
				
				//Click on submit button.
				//driver.findElement(submitButtonProductGroup).click();
				Alert alert = driver.switchTo().alert();
				alert.accept();
				Thread.sleep(4000);
				// Wait for message to be displayed.
				String currentUrl = driver.getCurrentUrl();
				System.out.println("Current URl is = " +currentUrl);
				try {
					if(currentUrl.equalsIgnoreCase("http://192.168.1.201:7111/adminx/add_product_group.aspx"))
					{
						wait.until(ExpectedConditions.visibilityOfElementLocated(errorMessage));
						String errorMessageText = driver.findElement(errorMessage).getText();
						//sheet.getRow(i).getCell(13).setCellValue(errorMessageText);
						sheet.getRow(i).createCell(12).setCellValue(errorMessageText);
						FileOutputStream fout = new FileOutputStream(src);
						wb.write(fout);
						fout.close();
						driver.findElement(By.id("ContentPlaceHolder1_ucaddproductgroup_BtnBack")).click();
						
					   	
					}
					if(currentUrl.equalsIgnoreCase("http://192.168.1.201:7111/adminx/view_product_group.aspx"))
					{
						wait.until(ExpectedConditions.visibilityOfElementLocated(successMessage));
						String successMessageText = driver.findElement(successMessage).getText();
						sheet.getRow(i).createCell(12).setCellValue(successMessageText);
						FileOutputStream fout = new FileOutputStream(src);
						wb.write(fout);
						fout.close();
     
					}
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}} 
	    	catch (Exception e)
	    	{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	        
	        
	        
	        
	       
	        
	        
	        }
	   
	    
	    
	    }
	    
	    
}    
	    
	
	
	//}



