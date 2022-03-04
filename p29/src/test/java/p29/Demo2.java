package p29;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Demo2 {

	@Test
	public void testA() {
		WebDriverManager.firefoxdriver().setup();

		// WebDriverManager.chromedriver().browserVersion("97").setup();
		WebDriver driver = new FirefoxDriver();
		driver.get("http://www.google.com");
		driver.close();

	}
}
