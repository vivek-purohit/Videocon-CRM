/**
 * 
 */
package com.POM.Pages;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

/***************************
 *                         *
 * @author Vivek Purohit   *
 *                         *
 ***************************
 */
public class POM_Login 
{
	WebDriver driver;

	
	
	// Initialize ID.
	By username = By.id("txtuser");
	By password = By.id("txtpass");
	By submitButton = By.id("savechk");
	By xpathIcon = By.xpath(".//*[@id='Li29']/a/div[1]/img");
	
	public POM_Login(WebDriver driver)
	{
		this.driver = driver;
	}
	
	public void Openurl()
	{
		driver.get("http://192.168.1.201:7111");
	}
	
	// C_Head Login
	public void C_Head_Login()
	{
		WebDriverWait wait = new WebDriverWait(driver,30);
		driver.findElement(username).sendKeys("chead");
		driver.findElement(password).sendKeys("123456");
		driver.findElement(submitButton).click();
		wait.until(ExpectedConditions.elementToBeClickable(xpathIcon));
		driver.findElement(xpathIcon).click();
	}
	
	//AdminX Login
	public void AdminXLogin()
	{
		WebDriverWait wait = new WebDriverWait(driver,30);
		driver.findElement(username).sendKeys("adminx");
		driver.findElement(password).sendKeys("123456");
		driver.findElement(submitButton).click();
		wait.until(ExpectedConditions.elementToBeClickable(xpathIcon));
		driver.findElement(xpathIcon).click();
	}
	
	// Warehouse login (Dhulsi.wam)
	public void WarehouseLogin()
	{
		WebDriverWait wait = new WebDriverWait(driver,30);
		driver.findElement(username).sendKeys("shopdirect.wam");
		driver.findElement(password).sendKeys("123456");
		driver.findElement(submitButton).click();
		wait.until(ExpectedConditions.elementToBeClickable(xpathIcon));
		driver.findElement(xpathIcon).click();
	
	}
	
	
	

}
