/**
 * 
 */
package com.POM.Pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

/*************************
 *                       *
 * @author Vivek Purohit *
 *                       *
 *************************
 */

/*
 * Process-1] Open Add Courier.
 * Process-2] Open View Product group.
 * Process-3] Open Add Scheme.
 * Process-4] Warehouse --> Open View Zone.
 * Process-5] Warehouse --> Open View Section.
 * Process-6] Warehouse --> Open View Shelf.
 * Process-7] Warehouse --> Open View Rack.
 * Process-8] AdminX --> Open Add agent Page. Masters--> Call center --> Add agent.
 * Process-9] CHead ---> Open raise Dealer PO.
 * Process - 10] Warehouse --> Open Picklist.
 * Process -11] Warehouse --> Open Courier Manifest.
 */
public class POM_Operations 
{
	WebDriver driver;
	public POM_Operations(WebDriver driver)
	{
		this.driver = driver;
	}
	
	public void WaitforLoadingImageToDisappaer()
	{
		WebDriverWait wait1 = new WebDriverWait (driver,30);
		By loadingimage = By.id("loading-image");
		wait1.until(ExpectedConditions.invisibilityOfElementLocated(loadingimage));
		
	}
	
	public void clickAdminSettings()
	{
		driver.findElement(By.xpath(".//*[@id='Li29']/a/div[1]")).click();
	}
	
	//Process-1] Open Add Courier.
		public void OpenAddCourier() throws InterruptedException
		{
			driver.findElement(By.linkText("MASTERS")).click();
		    driver.findElement(By.linkText("DISTRIBUTION USERS SETTING")).click();
		    driver.findElement(By.linkText("VIEW COURIER")).click();
		    Thread.sleep(3000);
		    driver.findElement(By.id("ContentPlaceHolder1_viewcourier_btnadd")).click();
		    
		    
		}
		
	     // Process-2] Open View Product group.
		public void OpenProductGroup()
		{
			driver.findElement(By.linkText("PRODUCT")).click();
		    driver.findElement(By.linkText("PRODUCT SETTING")).click();
		    driver.findElement(By.linkText("VIEW PRODUCT GROUP")).click();	
		}
		
		// Process-3] Open Add Scheme.
		public void OpenAddScheme()
		{
			driver.findElement(By.linkText("PRODUCT")).click();
		    driver.findElement(By.linkText("PRODUCT SETTING")).click();
		    driver.findElement(By.linkText("SCHEME MASTER")).click();
		    
		}
		
		// Process-4] Warehouse --> Open View Zone.
		public void warehouse_OpenViewZone()
		{
			driver.findElement(By.linkText("MASTERS")).click();
		    driver.findElement(By.linkText("PUTAWAY")).click();
		    driver.findElement(By.linkText("VIEW ZONE")).click();
		}
		
		// Process-5] Warehouse --> Open View Section.
		public void warehouse_OpenViewSection()
		{
			driver.findElement(By.linkText("MASTERS")).click();
		    driver.findElement(By.linkText("PUTAWAY")).click();
		    driver.findElement(By.linkText("VIEW SECTION")).click();
		}
		
		// Process-8] AdminX --> Open Add agent Page. Masters--> Call center --> Add agent.
		public void adminX_OpenAddAgent()
		{
			driver.findElement(By.linkText("MASTERS")).click();
		    driver.findElement(By.linkText("CALLCENTER USER")).click();
		    driver.findElement(By.linkText("ADD AGENT")).click();
		}
		// Process-9] CHead ---> Open raise Dealer PO.
		public void chead_OpenRaiseDealerPO()
		{
			driver.findElement(By.linkText("PURCHASE")).click();
			driver.findElement(By.linkText("LP PO")).click();
		    driver.findElement(By.xpath("(//a[contains(text(),'LP PO')])[2]")).click();
		}
		
		// Process - 10] Warehouse --> Open Picklist.
		public void warehouse_OpenPicklist()
		{
			driver.findElement(By.linkText("SHIPPING")).click();
		    driver.findElement(By.linkText("PICKLIST")).click();
		    driver.findElement(By.xpath("(//a[contains(text(),'PICKLIST')])[2]")).click();
		}
		
		// Process -11] Warehouse --> Open Courier Manifest.
		public void warehouse_OPenCourierManifest()
		{
			 driver.findElement(By.linkText("SHIPPING")).click();
			 driver.findElement(By.linkText("COURIER MANIFEST")).click();
			 driver.findElement(By.xpath("(//a[contains(text(),'COURIER MANIFEST')])[2]")).click();
		}
		
		
		

}
