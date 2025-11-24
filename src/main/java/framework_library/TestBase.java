package framework_library;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.InetAddress;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;

import org.openqa.selenium.Alert;
import org.openqa.selenium.NoAlertPresentException;
import org.testng.ITestContext;
import org.testng.ITestResult;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeSuite;

public class TestBase {
	protected AppLibrary appLibrary; // Application Library instance
	Properties usersProperties = null;
	private String suite;
	protected String testName;
	ITestContext testContext;

	@BeforeSuite
	public void cleanUp() throws Exception {
//		AppLibrary.cleanDirectory(AppLibrary.getConfig().getoutputPath());
	}

	@BeforeClass(alwaysRun = true)
	public void setUp(ITestContext context) throws Exception {

		suite = context.getCurrentXmlTest().getSuite().getName();
		suite = ((suite != null && !(suite.equals("Default suite"))) ? suite
				: InetAddress.getLocalHost().getHostName());
		suite = (suite.contains("Neo") ? "Neo" : suite);
		testName = this.getClass().getSimpleName();
		testName = ((testName != null && !(testName.equals("Default test"))) ? testName
				: this.getClass().getSimpleName());

		Date date = new Date();
		SimpleDateFormat sdf = new SimpleDateFormat("MMMddyyyyhhmmssaz");
		String currentDate = sdf.format(date);

		if (System.getProperty("Build") == null) {
			System.setProperty("Build", suite + "_" + currentDate);
			System.setProperty("Suite", suite);
		}
		System.setProperty("Test", testName);
		testContext = context;
	}

	public Properties loadUserProperties() {
		if (usersProperties == null) {
			usersProperties = new Properties();
			try {
				File f = new File("Users.properties");
				if (!f.exists()) {
					f = new File("TestData/users.properties");
				}
				usersProperties.load(new FileInputStream(f));

			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return usersProperties;
	}

	public String getProviderUserID() {
		String email = null;

		if (appLibrary.baseUrlSolo.contains("test")) {
			email = loadUserProperties().getProperty("PROVIDER_USER_STAGE");
		} else if (appLibrary.baseUrlSolo.contains("preprod")) {
			System.out.println("Properties: " + loadUserProperties());
			email = loadUserProperties().getProperty("PROVIDER_USER_PREPROD");
		} else if (appLibrary.baseUrlSolo.contains("dev") || appLibrary.baseUrlSolo.contains("uat")) {
			email = loadUserProperties().getProperty("PROVIDER_USER_DEV");
		}
		return email;
	}
	
	public String getProviderUserIDDirectory() {
		String email = null;
			email = loadUserProperties().getProperty("USERNAMED");
			return email;
	}
	
	public String getProviderPasswordIDDirectory() {
		String email = null;
			email = loadUserProperties().getProperty("PASSWORDD");
			return email;
	}

	public String getProviderPassword() {
		String pass = null;

		System.out.println("URL: " + appLibrary.baseUrlSolo);
		if (appLibrary.baseUrlSolo.contains("test")) {
			pass = loadUserProperties().getProperty("PROVIDER_PASS_STAGE");
		} else if (appLibrary.baseUrlSolo.contains("preprod")) {
			pass = loadUserProperties().getProperty("PROVIDER_PASS_PREPROD");
		} else if (appLibrary.baseUrlSolo.contains("dev") || appLibrary.baseUrlSolo.contains("uat")) {
			pass = loadUserProperties().getProperty("PROVIDER_PASS_DEV");
		}
		return pass;
	}

	public String getUserID() {
		String email = null;
		email = loadUserProperties().getProperty("USERNAME");
		return email;
	}

	public String getPassword() {
		String pass = null;
		pass = loadUserProperties().getProperty("PASSWORD");
		return pass;
	}

	public String getUserType() {
		String user = null;
		if (appLibrary.baseUrlSolo.contains("uat")) {
			user = loadUserProperties().getProperty("USER_TYPE");
		}
		return user;
	}

	@AfterMethod
	public void checkAlerts(ITestResult result) throws Exception {
		if (result.getStatus() == ITestResult.FAILURE) {
			try {
				String screenshotName = result.getName() + "_" + appLibrary.browser + "_"
						+ appLibrary.getFormattedDate() + ".png";
				appLibrary.getScreenshot(screenshotName);
				AppLibrary.autoLogger("Failed at URL: " + appLibrary.getCurrentDriverInstance().getCurrentUrl());
				int paramsLength = result.getParameters().length;
				AppLibrary.autoLogger("ScreenShot for " + testName + " "
						+ ((paramsLength > 0) ? " with parameter " + result.getParameters()[1] : "") + "saved as "
						+ screenshotName + ".png");

			} catch (Exception e) {
				AppLibrary.autoLogger("Failed fetching URL and screenshot due to error:" + e.getMessage());
				e.printStackTrace();
			}

			if (appLibrary.getCurrentSessionID() != null) {
				AppLibrary.autoLogger("Session id for " + testName + " is " + appLibrary.getCurrentSessionID());
				AppLibrary.autoLogger("Session details for " + testName
						+ " can be found at https://www.browserstack.com/automate/sessions/"
						+ appLibrary.getCurrentSessionID() + ".json");
			}
		}

		try {
			Alert alert = appLibrary.getCurrentDriverInstance().switchTo().alert();
			String alertText = alert.getText();
			alert.accept();
			AppLibrary.autoLogger("Unexpected Alert opened: " + alertText);
		} catch (NoAlertPresentException e) {
			// Alert was not present
		}

		try {
			while (appLibrary.getCurrentDriverInstance().getWindowHandles().size() > 1) {
				appLibrary.getCurrentDriverInstance().switchTo()
						.window((String) appLibrary.getCurrentDriverInstance().getWindowHandles().toArray()[1]).close();
			}

		} catch (Exception e) {
			if (appLibrary.getCurrentSessionID() != null) {
				AppLibrary.autoLogger("Session id for " + testName + " is " + appLibrary.getCurrentSessionID());
				AppLibrary.autoLogger("Session details for " + testName
						+ " can be found at https://www.browserstack.com/automate/sessions/"
						+ appLibrary.getCurrentSessionID());
			}
			throw e;
		}

		appLibrary.closeBrowser();
	}

	@AfterClass(alwaysRun = true)
	public void quitBrowser() {
		appLibrary.closeBrowser();
		AppLibrary.autoLogger("Closing the Browser Successfully");
	}

}
