package pages;

import static org.testng.Assert.assertEquals;
import static org.testng.Assert.assertNotEquals;
import static org.testng.Assert.assertNotSame;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.How;
import org.testng.Assert;

import junit.awtui.Logo;

public class CustomerSignInPage {

	WebDriver driver;

	public CustomerSignInPage(WebDriver driver) {
		this.driver = driver;

	}

	@FindBy(how = How.XPATH, using = "//input[@title='Email']")
	WebElement Username;
	@FindBy(how = How.XPATH, using = "//fieldset[@class='fieldset login']//input[@id='pass']")
	WebElement Password;
	@FindBy(how = How.XPATH, using = "(//span[contains(.,'Sign In')])[1]")
	WebElement signInButton;
	@FindBy(how= How.XPATH, using ="(//i[@class='fa fa-caret-down'])[2]") WebElement Dropdown;
	@FindBy(how = How.XPATH, using = "//a[contains(text(),'Logout')]") WebElement Logout;

	public void enterUsername(String username) {
		Username.clear();
		Username.sendKeys(username);
	}

	public void enterPassword(String password) {
		Password.clear();
		Password.sendKeys(password);
	}

	
	public void verifyTitle() {
		String actualTitle = driver.getTitle();
		String expectedTitle = "Dashboard / Vnecoms Marketplace";
		assertEquals(actualTitle, expectedTitle);

	}
	
	public void verifyTitleinvalid() {
		String actualTitle = driver.getTitle();
		String expectedTitle = "Customer Login";
		assertEquals(actualTitle, expectedTitle);

	}

	public void clickOnSignInButton() {
		// TODO Auto-generated method stub
		signInButton.click();
	}
	
	public void clickOnLogOutlink() {
		Dropdown.click();
		Logout.click();
	}

}
