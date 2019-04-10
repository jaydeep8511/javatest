package pages;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;

public class HomePage {

	WebDriver driver;

	public HomePage(WebDriver driver) {
		this.driver = driver;

	}

	@FindBy(how = How.XPATH, using = "//div[@class='panel header']//a[contains(text(),'Sign In')]")
	WebElement signIn;

	public void clickOnSignInLink() {
		signIn.click();
	}

}
