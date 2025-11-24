package pageObject;

import framework_library.AppLibrary;

public class LoginPage {
	
	private AppLibrary appLibrary;
	
	public String userNameL = "xpath:-://input[@id='username']";
	public String passwordL = "xpath:-://input[@id='password']";
	public String logInL = "name:-:submit";
	
	
	public LoginPage(AppLibrary appLibrary) {
		this.appLibrary = appLibrary;
	}
	
	public void login(String userName, String password) throws Exception {
		appLibrary.enterText(userNameL, userName);
		appLibrary.enterText(passwordL, password);
		appLibrary.clickElement(logInL);
	}

}
