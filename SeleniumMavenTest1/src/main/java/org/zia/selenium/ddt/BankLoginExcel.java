package org.zia.selenium.ddt;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import org.zia.selenium.util.ExcelDataConfig;

public class BankLoginExcel {
	
	WebDriver driver;
	
//	@BeforeClass
//	  public void setUp() throws Exception {
//		System.setProperty("webdriver.chrome.driver", "C:/chromedriver_win32/chromedriver.exe");
//			
//		driver = new ChromeDriver();
//		driver.manage().window().maximize();
//	    baseUrl = "https://s1.demo.opensourcecms.com/wordpress/wp-login.php";
//	    driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
//	  }
	
	@Test(dataProvider="bank")
	public void loginToBank(String username, String password) throws InterruptedException
	{
		System.setProperty("webdriver.chrome.driver", "C:/chromedriver_win32/chromedriver.exe");
		
		driver = new ChromeDriver();
		//WebDriver driver = new FirefoxDriver();
		driver.manage().window().maximize();
	    driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.get("http://localhost:8185/index");
//		driver.findElement(By.id("user_login")).sendKeys(username);
		driver.findElement(By.name("username")).sendKeys(username);
//		driver.findElement(By.id("user_pass")).sendKeys(password);
		driver.findElement(By.name("password")).sendKeys(password);
		
		driver.findElement(By.xpath("/html/body/div[1]/div/div/form/button")).click();
		//driver.findElement(By.id("wp-submit")).click();
		Thread.sleep(5000);
		System.out.println(driver.getTitle());
		
		Assert.assertNotEquals(driver.getTitle().contains("Dashboard"), "User is not able to login. Invalid credentials");
		System.out.println("Page title verified. User is able to successfully login");
		
		
	}
	
	@AfterMethod
	public void teardown() {
		
		driver.quit();
		
	}

	@DataProvider(name="bank")
	public Object[][] passData() {
		
//		ExcelDataConfig config = new ExcelDataConfig("C:\\Users\\Tariq Ahsan\\Desktop\\ISB Java JEE Training\\Selenium\\TestData\\BankInputData.xls");
		ExcelDataConfig config = new ExcelDataConfig("TestData/BankInputData.xls");
		int rows = config.getRowCount(0);
		
		Object[][] data = new Object[rows][2]; 
		
		for(int i = 0; i < rows; i++) {
			data[i][0] = config.getData(0, i, 0); // username
			data[i][1] = config.getData(0, i, 1); // password
		}
		
//		Object[][] data = new Object[3][2];
//		data[0][0] = "admin1";
//		data[0][1] = "demo1";
//
//		data[1][0] = "opensourcecms";
//		data[1][1] = "opensourcecms";
//		
//		data[2][0] = "admin";
//		data[2][1] = "demo3";
		
		return data;
		
	
	}
}
