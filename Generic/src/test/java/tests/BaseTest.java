package tests;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;

public class BaseTest {

	public static WebDriver driver = null;

	@BeforeSuite
	public void init() {
		System.setProperty("webdriver.chrome.driver", "D:\\selenium\\driver\\chromedriver_win32 (1)\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);

		driver.get("http://generic.dmob.in/");
	}

	@AfterSuite
	public void teardown() {
		BaseTest.driver.close();
	}

}
