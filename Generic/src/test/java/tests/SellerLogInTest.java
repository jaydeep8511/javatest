package tests;

import org.openqa.selenium.support.PageFactory;
import org.testng.annotations.Test;

import pages.CustomerSignInPage;
import pages.HomePage;

public class SellerLogInTest extends BaseTest {
	
	@Test
	public void sellerLoginWithCorrectCredentials() {
		
		HomePage homepage = PageFactory.initElements(driver, HomePage.class);
		homepage.clickOnSignInLink();
		
		CustomerSignInPage signInpage = PageFactory.initElements(driver, CustomerSignInPage.class);
		signInpage.enterUsername("seller02@gmail.com");
		signInpage.enterPassword("admin@123");
		signInpage.clickOnSignInButton();
		signInpage.verifyTitle();
		signInpage.clickOnLogOutlink();
	}
	
	@Test
	public void sellerLoginWithInCorrectCredentials() {
		
		HomePage homepage = PageFactory.initElements(driver, HomePage.class);
		homepage.clickOnSignInLink();
		
		CustomerSignInPage signInpage = PageFactory.initElements(driver, CustomerSignInPage.class);
		signInpage.enterUsername("sellerdfdf@gmail.com");
		signInpage.enterPassword("admin@123");
		signInpage.clickOnSignInButton();
		signInpage.verifyTitleinvalid();
		
	}
}
