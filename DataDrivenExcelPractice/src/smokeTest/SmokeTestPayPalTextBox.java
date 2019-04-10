package smokeTest;

import java.util.concurrent.TimeUnit;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import poifiles.Utilitylib;

public class SmokeTestPayPalTextBox<Screenshot> {
	static WebDriver driver;
	static Utilitylib ut;
	

	@BeforeMethod
	public static void beforet() throws Exception {
			
	
	}

	@BeforeTest
	public void openBrowse() {

		ut = new Utilitylib("D:\\QA\\Automation\\practice\\practice template.xlsx");
		System.setProperty("webdriver.chrome.driver", "D:\\selenium\\driver\\chromedriver_win32 (1)\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);

	}
	@AfterTest
	public void closeBrowser()
	{
		//driver.close();
	}

	@AfterMethod
	public void tearDown(ITestResult result) {
		if (ITestResult.FAILURE == result.getStatus()) {
			Utilitylib.Screenshot(driver, "failure snap_" + result.getName());
		} else {
			Utilitylib.Screenshot(driver, "Pass snap_" + result.getName());
			

		}

	}

	// start write script from here

	@Test(priority = 1)
	 public void script1() throws Exception {


		Utilitylib.scriptexcel(driver, ut, 0, 2, 5, 1);
		
	
		
		
		

	 }
	
		 

	// End script here

}
