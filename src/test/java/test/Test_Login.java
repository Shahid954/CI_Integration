package test;

import org.testng.annotations.BeforeClass;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import framework_library.AppLibrary;
import framework_library.TestBase;
import pageObject.LoginPage;

public class Test_Login extends TestBase {
	
	@BeforeClass
	@Parameters("browser")
	public void setup() {
		appLibrary = new AppLibrary("Test_Login");
	}
	
	
	
	@Test
	public void teamConnectLogin() throws Exception {
		appLibrary.getDriverInstance();
		appLibrary.launchApp();
		LoginPage page = new LoginPage(appLibrary);
		page.login("am","am");
		
	}

}
