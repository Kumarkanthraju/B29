package page;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.WebDriverWait;

import generic.WebGeneric;

public class EnterTimeTrackPage {

	@FindBy(id="logoutLink")
	private WebElement logoutLink;
	
	public EnterTimeTrackPage(WebDriver driver) {
		PageFactory.initElements(driver,this);
	}
	
	public boolean verifyHomePageIsDisplayed(WebDriverWait wait,String expectedTitle) {
//		return WebGeneric.verifyTitle(wait, expectedTitle);
//		return WebGeneric.verifyUrl(wait, expectedUrl);
		return WebGeneric.verifyElementDisplayed(wait, logoutLink);
	}
}
