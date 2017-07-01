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
import java.lang.reflect.GenericArrayType;
import java.util.ArrayList;
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
public class TC_08_Picklist_Process 
{
	WebDriver driver;
	WebDriverWait wait;
	ExtentReports reports;
	ExtentTest extent;
	POM_Login login;
	POM_Operations ops;
	HSSFWorkbook workbook;
	HSSFSheet sheet;
	HSSFSheet sheet1;
	HSSFCell cell;
	 
	
	@BeforeClass(alwaysRun = true)
	public void TestSetup()
	{
		driver = new FirefoxDriver();
		login = new POM_Login(driver);
		ops = new POM_Operations(driver);
		driver.manage().window().maximize();
		login.Openurl();
		wait = new WebDriverWait(driver,30);
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
        reports = new ExtentReports("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_08_Picklist_Process.xls", true);
		extent = reports.startTest("TC_08_Picklist_Process", "To check the functionality of Picklist process");
	}
	
	
	/*
	 * Step -1 ] Login with warehouse login credentials.
	 * Step -2 ] Open Picklist page. (Shipping --> Picklist--> Picklist)

	 * */
	@Test(priority = 0, groups = "testing2")
	public void TestLogin()
	{
		try
		{
			login.WarehouseLogin();
			extent.log(LogStatus.INFO, "Logged In successfully");
			wait.until(ExpectedConditions.elementToBeClickable(By.linkText("SHIPPING")));
			ops.warehouse_OpenPicklist();
			extent.log(LogStatus.INFO, "Picklist page opened");
		}
		catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	/*
	 * Step -1 ] Click on "Create Manual" button.
	 * Step -2 ] Import excel.
	 * Step -3 ] Import data for Order number.
	 * Step -4 ] Click on search button.
	 * Step -5 ] Click on the checkbox.
	 * Step -6 ] Click on "Create Picklist" Button.
	 * Step -7 ] Accept alert.
	 * Step -8] Retrieve the message displayed.
	 * Step -9] Write the retrieved message in the excel.
	 */
	
	@Test(priority = 1, groups = "test")
	public void Test_CreatePicklist() throws IOException, InterruptedException
	{
		try 
		{
			// Initialize IDs.
			By createManualBtn = By.id("ContentPlaceHolder1_btnManual");
			By searchButton = By.id("ContentPlaceHolder1_btnSear");
			By orderNo_TxtBox = By.id("ContentPlaceHolder1_txtEnquiry");
			By chkBox = By.id("ContentPlaceHolder1_GridView1_cbHdrsel");
			By createPicklistBtn = By.id("ContentPlaceHolder1_btnPickList");
			
			
			
			// Step -1 ] Click on "Create Manual" button.
			wait.until(ExpectedConditions.elementToBeClickable(createManualBtn));
			driver.findElement(createManualBtn).click();
			wait.until(ExpectedConditions.elementToBeClickable(searchButton));
			
			// Step -2 ] Import excel.
			File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_08_Picklist_Process.xls");
			FileInputStream fin = new FileInputStream(src);
			workbook = new HSSFWorkbook(fin);
			sheet = workbook.getSheetAt(0);
			
			// Step -3 ] Import data for Order number.
			cell = sheet.getRow(1).getCell(1);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			String orderNumbers = cell.getStringCellValue();
			List<String> orderNumberArray;
			orderNumberArray = Arrays.asList(orderNumbers.split(", "));
			for(int i = 0; i<orderNumberArray.size(); i++)
			{
				driver.findElement(By.id("ContentPlaceHolder1_txtEnquiry")).sendKeys(orderNumberArray.get(i));
			    driver.findElement(By.id("ContentPlaceHolder1_txtEnquiry")).sendKeys(Keys.ENTER);
			}
			
			// Step -4 ] Click on search button.
			driver.findElement(searchButton).click();
			wait.until(ExpectedConditions.elementToBeClickable(chkBox));
			
			// Step -5 ] Click on the checkbox.
			driver.findElement(chkBox).click();
			
			// Step -6 ] Click on "Create Picklist" Button.
			driver.findElement(createPicklistBtn).click();
			
			try {
				// Step -7 ] Accept alert.
				Alert alert = driver.switchTo().alert();
				alert.accept();
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			// Step -8] Retrieve the message displayed.
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ContentPlaceHolder1_lblmsg")));
			String msgDisplayed = driver.findElement(By.id("ContentPlaceHolder1_lblmsg")).getText();
			System.out.println(msgDisplayed);
			extent.log(LogStatus.INFO, msgDisplayed);
			
			// Step -9] Write the retrieved message in the excel.
			sheet.getRow(1).createCell(2).setCellValue(msgDisplayed);
			//cell.setCellValue(msgDisplayed);
			FileOutputStream fout = new FileOutputStream(src);
			workbook.write(fout);
			fout.close();
		} catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	/*
	 * Step -1 ] Click on view picklist button.
	 * Step -2 ] Click on edit button.
	 * Step -3 ] Click on receive button.
	 * Step -4 ] Accept alert.
	 * Step -5] Check for message displayed.
	 * Step -6] Click on the checkbox.
	 * Step -7] Import data for assigned courier.
	 * Step -8] Click on assigned courier button.
	 * Step -9] Accept alert.
	 * Step -10 ] Retrieve message displayed.
	 * Step -11 ] Set message displayed in the excel.
	 *  
	 */
	@Test(priority = 2, groups = "test")
	public void viewPicklist() throws IOException
	{
		try 
		{
			// Initialize ID.
			By viewPicklist = By.id("ContentPlaceHolder1_btnViewPick");
			By editBtn = By.xpath(".//*[@id='ContentPlaceHolder1_grdPickList']/tbody/tr[2]/td[6]/a");
			By receiveBtn = By.id("ContentPlaceHolder1_btnrecieve");
			By assignedBtn = By.id("ContentPlaceHolder1_btnCourier");
			By msgLocation = By.id("ContentPlaceHolder1_lblmsg");
			By chkBox = By.id("ContentPlaceHolder1_grdManifestOrders_cbHdrsel");
			By courierDD = By.id("ContentPlaceHolder1_ddlCourier");
			By generateAWB = By.id("ContentPlaceHolder1_btnAWBGen");
			String productCode[];
			
			
			// Import excel
			
			File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_08_Picklist_Process.xls");
			FileInputStream fin = new FileInputStream(src);
			workbook = new HSSFWorkbook(fin);
			sheet = workbook.getSheetAt(0);
						
						
			// Step -1 ] Click on view picklist button.
			wait.until(ExpectedConditions.elementToBeClickable(viewPicklist));
			driver.findElement(viewPicklist).click();
			
			// Step -2 ] Click on edit button.
			wait.until(ExpectedConditions.elementToBeClickable(editBtn));
			driver.findElement(editBtn).click();
			
			// Step -3 ] Click on receive button.
			wait.until(ExpectedConditions.elementToBeClickable(receiveBtn));
			driver.findElement(receiveBtn).click();
			
			// Step -4 ] Accept alert.
			Alert alert = driver.switchTo().alert();
			alert.accept();
			ops.WaitforLoadingImageToDisappaer();
			
			// Step -5] Check for message displayed.
			wait.until(ExpectedConditions.visibilityOfElementLocated(msgLocation));
			
			
			
			// Step -6] Check the condition in excel, if yes then import else click on generate AWB.
			cell = sheet.getRow(1).getCell(3);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			String cond_AssignCourier = cell.getStringCellValue();
			if(cond_AssignCourier.equalsIgnoreCase("Yes"))
			{
				// Step -6] Click on the checkbox.
				driver.findElement(chkBox).click();
				
				// Step -7] Import data for assigned courier.
				cell = sheet.getRow(1).getCell(4);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				Select courier = new Select(driver.findElement(courierDD));
				courier.selectByVisibleText(cell.getStringCellValue());
				
			
			
			
			
			// Step -8] Click on assigned courier button.
			driver.findElement(assignedBtn).click();
			
			// Step -9] Accept alert.
			Alert alert1 = driver.switchTo().alert();
			alert1.accept();
			ops.WaitforLoadingImageToDisappaer();
			
 }
			wait.until(ExpectedConditions.visibilityOfElementLocated(msgLocation));
			
			// Step -10 ] retrieve message displayed.
			String msgDisplayed = driver.findElement(msgLocation).getText();
			System.out.println(msgDisplayed);
			extent.log(LogStatus.INFO, msgDisplayed);
			
			// Step -11] Set message in the excel.
			//File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_08_Picklist_Process.xls");
			FileOutputStream fout = new FileOutputStream(src);
			sheet.getRow(1).createCell(5).setCellValue(msgDisplayed);
			workbook.write(fout);
			fout.close();
			
			// Step -12 ] Click on generate AWB.
			driver.findElement(chkBox).click();
			driver.findElement(generateAWB).click();
			
			// Step -13] Accept alert.
			try
			{
				Alert alert3 = driver.switchTo().alert();
				alert3.accept();
			}
			catch(Exception e)
			{
				System.out.println("No alert present for generate AWB");
			}
			
			// Step -14 ] Wait for message to be displayed.
			ops.WaitforLoadingImageToDisappaer();
			wait.until(ExpectedConditions.visibilityOfElementLocated(msgLocation));
			
			// Step -15 ] Retrieve the message displayed.
			String msgDisplayedAWB = driver.findElement(msgLocation).getText();
			System.out.println(msgDisplayedAWB);
			
			// Step-16 ] Write the message displayed in the excel.
			sheet.getRow(1).createCell(6).setCellValue(msgDisplayedAWB);
			FileOutputStream fout1 = new FileOutputStream(src);
			workbook.write(fout1);
			fout1.close();
			
			// Step - 17] Check for the product code in the table.
			
			/*
			 *  Step - 17.1 ] Calculate the size of the table.
			 *  Step - 17.2 ] check for the product code row count.
			 *  Step - 17.3 ] Check for the size of the product code rows.
			 *  Step - 17.4 ] Retrieve the product code from the row.
			 *  
			 */
			
			// Step - 17.1 ] Calculate the size of the table.
			int rowCountTable = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_grdManifestOrders']/tbody/tr")).size();
			
			// Step - 17.2 ] Check for the product code row count.
			int j = 0;
			for(int i = 0; i<rowCountTable; i++ )
			{
				// xpath for product code rows.
				String firstXpathPrdctCode = ".//*[@id='ContentPlaceHolder1_grdManifestOrders_dataRto_";
				String secondXpathPrdctCode = "']/tbody/tr/td/table/tbody/tr/td[1]";
				String finalXpathPrdctCode = firstXpathPrdctCode+j+secondXpathPrdctCode;
				System.out.println("General Xpath ID of product code rows = " +finalXpathPrdctCode);
				// Step - 17.3 ] Check for the size of the product code rows.
				int rowCountPrdctCode = driver.findElements(By.xpath(finalXpathPrdctCode)).size();
				productCode = new String[rowCountPrdctCode+1];
				// Step - 17.4 ] Retrieve the product code from the row.
				for(int k = 3; k<= rowCountPrdctCode; k+=2)
				{
					String secondXpathPrdctCodeRow = "']/tbody/tr/td/table/tbody/tr[";
					String thirdXpathPrdctCodeRow  = "]/td[1]";
					String finalXpathPrdctCodeRow  = firstXpathPrdctCode+j+secondXpathPrdctCodeRow+k+thirdXpathPrdctCodeRow;
				    System.out.println("Xpath for product code sub rows = " +finalXpathPrdctCodeRow);
				    
				    // Get the code in text format.
				    productCode[k] = driver.findElement(By.xpath(finalXpathPrdctCodeRow)).getText();
				    System.out.println("Product code retrieved as = " +productCode[k]); 
				    // write the product code in the sheet.
				    sheet1 = workbook.getSheetAt(1);
				  // Check the count in the array.
				   int a = sheet1.getLastRowNum();
				   System.out.println("last row number = " +a);
				   sheet1.createRow(sheet1.getLastRowNum()+1).createCell(0).setCellValue(productCode[k]);
				   
				   //sheet1.getRow(m).createCell(0).setCellValue(productCode[k]);
				    
				    FileOutputStream fout2 = new FileOutputStream(src);
				    workbook.write(fout2);
				    fout2.close();
				    
				}
				j++;
				
			}
		} 
		catch (Exception e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
	@Test(priority = 3, groups = "testing1")
	public void ScanBarcode() throws InterruptedException, AWTException
	{
		try 
		{
			String productCode[];
			String Barcode[];
			By prdctSummaryPopUp = By.id("ContentPlaceHolder1_pnlProduct");
			By prdctSmryPopUp_CloseBtn = By.id("ContentPlaceHolder1_lnkProductClose");
			By editbutton = By.xpath(".//*[@id='ContentPlaceHolder1_grdPickList']/tbody/tr[2]/td[6]/a");
			By srchButton = By.id("ContentPlaceHolder1_btnserch");
			By barcodeString = By.xpath(".//*[@id='ContentPlaceHolder1_gridscheme_code']/tbody/tr[2]/td[3]");
			By barcodeScanTxtBox = By.id("ContentPlaceHolder1_txtbarcode");
			By barcodeScanPopUp = By.xpath(".//*[@id='ContentPlaceHolder1_pnlPerson']/fieldset");
			By barcodeScanTxtBoxPopUp = By.id("ContentPlaceHolder1_txtProductBarcode");
			
			
			/*wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//*[@id='ContentPlaceHolder1_grdPickList']/tbody/tr[2]/td[6]/a")));
			driver.findElement(editbutton).click();
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//*[@id='ContentPlaceHolder1_grdManifestOrders']/tbody/tr")));*/
			
			// Step - 17.1 ] Calculate the size of the table.
				int rowCountTable = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_grdManifestOrders']/tbody/tr")).size();	
			    System.out.println("Number of rows in the table = " +rowCountTable);
				// Iterate the process for the number of rows.
			    String firstID = "ContentPlaceHolder1_grdManifestOrders_lnkScheme_"; 
			    for(int i = 0; i<rowCountTable-1; i++)
			     {
			    	// Click on the scheme name link to get the product code and if Barcode clickable.
			    	 String finalID = firstID+i;
			    	 System.out.println("ID of the clickable scheme in the table = " +finalID);
			    	 
			    	 //Click on the scheme.
			    	 driver.findElement(By.id(finalID)).click();
			    	 ops.WaitforLoadingImageToDisappaer();
			    	 wait.until(ExpectedConditions.visibilityOfElementLocated(prdctSummaryPopUp));
			    	 if(driver.findElements(prdctSummaryPopUp).size() != 0)
			    	 {
			    		 // Get the size of the row in the pop ups.
			    		 int rowCountPopup_PrdctSmry = driver.findElements(By.xpath(".//*[@id='ContentPlaceHolder1_grdSchemeProduct']/tbody/tr")).size();
			    		 System.out.println("number of rows in product summary pop up = " +rowCountPopup_PrdctSmry);
			    		 // Iterate the process for getting the product code.
			    		 productCode = new String[rowCountPopup_PrdctSmry];
			    		 int m = 0;
			    		 for(int j = 2; j<=rowCountPopup_PrdctSmry; j++)
			    		 {
			    			 // creating custom xpaths
			    			 String firstXpath = ".//*[@id='ContentPlaceHolder1_grdSchemeProduct']/tbody/tr[";
			    		     String thirdXpath_barcode = "]/td[4]";
			    		     String thirdXpath_productCode = "]/td[2]";
			    		     String finalXpath_barcode = firstXpath+j+thirdXpath_barcode;
			    		     String finalXpath_productCode = firstXpath+j+thirdXpath_productCode;
			    		     System.out.println("xpath for barcode column = " +finalXpath_barcode);
			    		     System.out.println("xpath for product code column = " +finalXpath_productCode);
			    		     
			    		     // Getting the text for the barcode column if it is "Y" than it will scan for product code.
			    		     String barCodeApplicable = driver.findElement(By.xpath(finalXpath_barcode)).getText();
			    		     System.out.println("Check if barcode scannable = " +barCodeApplicable);
			    		     
			    		     if(barCodeApplicable.equalsIgnoreCase("Y"))
			    		     {
			    		    	 productCode[m] = driver.findElement(By.xpath(finalXpath_productCode)).getText();
			    		    	 System.out.println("Product code retreived as = " +productCode[m]);
			    		    	 m += 1;
			    		     }
			    		     	    		     
			    		 }
			               // Click on close button.
			    		 
			    		 driver.findElement(prdctSmryPopUp_CloseBtn).click();
			    		 
					     // check how many elements are present in the product code array.
					     int count = 0;
					     for(int x = 0; x<rowCountPopup_PrdctSmry; x++)
					     {
					    	 if(productCode[x] != null)
					    	 {
					    		 count++;
					    	 }
					    	 
					     }
					     System.out.println("Total number of scannable product code is equal to = " +count);
					     
					     // Scan barcode from the barcode page.
					     Barcode = new String[count+1];
					  // Open barcode page and retrieve the barcode.
					     // Get current url.
					     String currentURLPicklist = driver.getCurrentUrl();
			    		// Navigate to the barcode status url.
					     driver.get("http://192.168.1.201:7111/Warehouse/Barcode_status.aspx");
					     wait.until(ExpectedConditions.elementToBeClickable(By.id("ContentPlaceHolder1_radioButtonList_2")));
					     // Click on the radio button.
					     driver.findElement(By.id("ContentPlaceHolder1_radioButtonList_2")).click();
						 // Enter the product code in the text box.
					     Thread.sleep(2500);
					     
					     
					  // Enter the product code in the text box and get the barcode as text. 
					     for (int k = 0; k<count ; k++)
					     {
					    	 wait.until(ExpectedConditions.elementToBeClickable(By.id("ContentPlaceHolder1_radioButtonList_2")));
					    	 driver.findElement(By.id("ContentPlaceHolder1_radioButtonList_2")).click();
					    	 Thread.sleep(2500);
					    	 driver.findElement(By.id("ContentPlaceHolder1_txtschemebarcode")).sendKeys(productCode[k]);
					    	 System.out.println("product code entered = " +productCode[k]);
					    	 // Click on search button.
					    	 driver.findElement(srchButton).click();
					    	 wait.until(ExpectedConditions.visibilityOfElementLocated(barcodeString));
					    	 // Retrieve the barcode.
					    	 
					    	 Barcode[k] = driver.findElement(barcodeString).getText();
					    	 System.out.println("Barcode retrieved as = " +Barcode[k]);
					    	 
					    	 driver.get(driver.getCurrentUrl());
					     } 
					    	// Once barcode are retrieved get back to the main url.
						     driver.get(currentURLPicklist);
						     wait.until(ExpectedConditions.visibilityOfElementLocated(barcodeScanTxtBox));
						     // Enter the retrieved barcode in the barcode scan text box.
						     for(int l = 0; l<count; l++)
						     {
						    	 driver.findElement(barcodeScanTxtBox).sendKeys(Barcode[l]);
						    	 // Press enter key.
						    	 driver.findElement(barcodeScanTxtBox).sendKeys(Keys.ENTER);
						    	 ops.WaitforLoadingImageToDisappaer();
						    	 Thread.sleep(1500);
						    	 //WebElement ele = driver.findElement(By.xpath(".//*[@id='ContentPlaceHolder1_pnlPerson']/fieldset"));
						    	 // if pop up is displayed then enter the barcode in the pop up.
						    	 
						    	 if(driver.findElements(barcodeScanPopUp).size() != 0)
						    	 {
						    		 for(int x = l+1; x<count; x++ )
						    		 {
						    			 driver.findElement(barcodeScanTxtBoxPopUp).clear();
						    			 System.out.println("Pop up was displayed and script running fine");
						    			 driver.findElement(barcodeScanTxtBoxPopUp).sendKeys(Barcode[x]);
						    			 driver.findElement(barcodeScanTxtBoxPopUp).sendKeys(Keys.ENTER);
						    			 System.out.println("barcode entered is = "+Barcode[x]);
						    			 Thread.sleep(2000);
						    		 }
						    	 }	 
						    			 // Close the new opened pop up windows.
						    			 Robot robot = new Robot();
						    			 // press al + f4.
						    			 for(int y = 0; y<2; y++)
						    			 {
						    				 robot.keyPress(KeyEvent.VK_ALT);
						    				 robot.keyPress(KeyEvent.VK_F4);
						    				 robot.delay(1000);
						    				 robot.keyRelease(KeyEvent.VK_ALT);
						    				 robot.keyRelease(KeyEvent.VK_F4);
						    				 Thread.sleep(1000);
						    			 }
						    		 
						    	 
						    
					     }
					     
				     
			    		 
			    	 }
			    	 
			     }
		} 
		catch (Exception e) 
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		   
		
		 }
	
	/*
	 * Step -1] Open courier manifest.
	 * Step -2] Click on create Manifest.
	 * Step- 3] Select shipping provider.(Import data from excel)
	 * Step -4] Click on save and close.(Accept alert)
	 * Step -5] Click on edit button.
	 * Step -6] Enter order number. (Import data from excel sheet).
	 * Step -7] Write the message in 
	 * Step -7] Enter scan packing bar code. (Import data from excel sheet).
	 * Step -8] Click on complete manifest button.
	 * */
	@Test(priority = 4, groups = "testing2")
	public void CourierManifest() throws IOException, InterruptedException, AWTException
	{
		try 
		{
			By createManifestBtn = By.id("ContentPlaceHolder1_btnAutoMatic");
			By editBtn_Manifest = By.xpath(".//*[@id='ContentPlaceHolder1_grdPickList']/tbody/tr[2]/td[9]/a");
			By orderNumberTxtBox = By.id("ContentPlaceHolder1_txtOrder");
			By scanPckngBarcode = By.id("ContentPlaceHolder1_txtScanPacking");
			By completeManifestBtn = By.id("ContentPlaceHolder1_btnComplete");
			By shpngProvider_PopupDD = By.id("ContentPlaceHolder1_ddlShipment");
			By saveAndcloseBtn = By.id("ContentPlaceHolder1_btnManifest");
			By createManifest_Popup = By.id("ContentPlaceHolder1_pnlPerson");
			By msgDisplayed = By.id("ContentPlaceHolder1_lblmsg");
			
			// Import excel.
			File src = new File("C:\\Users\\Admin\\Desktop\\DD_FrmWrk\\Videocon\\TC_08_Picklist_Process.xls");
			FileInputStream fin = new FileInputStream(src);
			workbook = new HSSFWorkbook(fin);
			sheet = workbook.getSheetAt(0);
			
			// Step -1] Open courier manifest.
			wait.until(ExpectedConditions.elementToBeClickable(By.linkText("SHIPPING")));
			ops.warehouse_OPenCourierManifest();
			wait.until(ExpectedConditions.elementToBeClickable(createManifestBtn));
			
			// Step -2] Click on create Manifest.
			driver.findElement(createManifestBtn).click();
			wait.until(ExpectedConditions.visibilityOfElementLocated(createManifest_Popup));
			
			 // Step- 3] Select shipping provider.
			cell = sheet.getRow(1).getCell(7);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			Select shippingProvider = new Select(driver.findElement(shpngProvider_PopupDD));
			shippingProvider.selectByVisibleText(cell.getStringCellValue());
			wait.until(ExpectedConditions.visibilityOfElementLocated(createManifest_Popup));
			
			// Step -4] Click on save and close.(Accept alert).
			driver.findElement(saveAndcloseBtn).click();
			Alert alert = driver.switchTo().alert();
			alert.accept();
			wait.until(ExpectedConditions.visibilityOfElementLocated(msgDisplayed));
			
			// Write message in excel.
			String messageCrtmanifest = driver.findElement(msgDisplayed).getText();
			sheet.getRow(1).createCell(8).setCellValue(messageCrtmanifest);
			FileOutputStream fout = new FileOutputStream(src);
			workbook.write(fout);
			fout.close();
			
			// Step -5] Click on edit button.
			driver.findElement(editBtn_Manifest).click();
			wait.until(ExpectedConditions.visibilityOfElementLocated(orderNumberTxtBox));
			
			// Step -6] Enter order number. (Import data from excel sheet).
			cell = sheet.getRow(1).getCell(1);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			String orderNumbers = cell.getStringCellValue();
			List<String> orderNumberArray = Arrays.asList(orderNumbers.split(", "));
			System.out.println("Size of the order number array = " +orderNumberArray.size());
			String msgDisplayed_PckngBarcode;
			By msgDisplayed1 = By.id("ContentPlaceHolder1_lblmsg");
			for(int i = 0; i<orderNumberArray.size(); i++)
			{
				
				driver.findElement(orderNumberTxtBox).sendKeys(orderNumberArray.get(i));
				driver.findElement(orderNumberTxtBox).sendKeys(Keys.ENTER);
				// Message displayed after scanning orer number.
				wait.until(ExpectedConditions.visibilityOfElementLocated(msgDisplayed1));
				String msgDisplayed_OrderScan = driver.findElement(msgDisplayed1).getText();
				System.out.println("Message displayed after scanning order number = " +msgDisplayed_OrderScan);
				
				// Step -7] Enter scan packing bar code. (Import data from excel sheet).
				cell = sheet.getRow(1).getCell(9);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				driver.findElement(scanPckngBarcode).sendKeys(cell.getStringCellValue());
				driver.findElement(scanPckngBarcode).sendKeys(Keys.ENTER);
				Thread.sleep(4000);
				
			}	
			    msgDisplayed_PckngBarcode = driver.findElement(msgDisplayed1).getText();
				if(msgDisplayed_PckngBarcode.equalsIgnoreCase("Packing Code scanned successfully"))
				{
					driver.findElement(completeManifestBtn).click();
					Alert alert2 = driver.switchTo().alert();
					alert2.accept();
					Thread.sleep(5000);
					// Get the count of windows.
					
					/*IList<String> window = new IList<String> (driver.getWindowHandles());
					if(window.co)*/
					sheet.getRow(1).createCell(10).setCellValue("Process accompalished successfully");
					FileOutputStream fout2 = new FileOutputStream(src);
					workbook.write(fout2);
					fout2.close();
				}
		} catch (Exception e) 
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		}
		
		
		
		
		
		
	}
		 
		



	
	






















