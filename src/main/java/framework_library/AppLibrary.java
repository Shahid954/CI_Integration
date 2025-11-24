package framework_library;

import static org.testng.Assert.assertEquals;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Random;
import java.util.Scanner;
import java.util.Set;
import java.util.function.Function;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.RandomStringUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Platform;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.LocalFileDetector;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.safari.SafariDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Reporter;

import com.github.javafaker.Faker;

import io.github.bonigarcia.wdm.WebDriverManager;
import junit.framework.Assert;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class AppLibrary {

	public final long GLOBALTIMEOUT = 75;
	private WebDriver driver;
	private static Configuration config;
	public String baseUrlSolo;
	public String baseUrlDirectory;
	public String browser;
	public String device;
	private String currentTestName;
	private String currentSessionID;

	public AppLibrary(String testName) {
		this.currentTestName = testName;
		config = new Configuration();
	}

	public static Configuration getConfig() {
		if (config == null) {
			config = new Configuration();
		}

		return config;
	}

	public WebDriver getDriverInstance() throws MalformedURLException {
		DesiredCapabilities caps = new DesiredCapabilities();
		String browserVersion, os, browserStackOSVersion, remoteGridUrl, environment;

		this.browser = config.getBrowserName();
		baseUrlSolo = config.getURLSolo();
		baseUrlDirectory = config.getURLDirectory();
		environment = config.getExecutionEnvironment();

		switch (environment) {

		case "browserstack":
			browserStackOSVersion = config.getBrowserStackOSVersion();
			browserVersion = config.getBrowserVersion();
			os = config.getOS();

			if (config.getBrowserName().equalsIgnoreCase("IE")) {
				caps.setCapability("browser", "IE");
			} else if (config.getBrowserName().equalsIgnoreCase("GCH")
					|| config.getBrowserName().equalsIgnoreCase("chrome")) {
				caps.setCapability("browser", "Chrome");
			} else if (config.getBrowserName().equalsIgnoreCase("safari")) {
				caps.setCapability("browser", "Safari");
			} else {
				caps.setCapability("browser", "Firefox");
			}

			if (browserVersion != null && !browserVersion.equals("") && !browserVersion.equals("latest")) {
				caps.setCapability("browser_version", browserVersion);
			}

			if (browserStackOSVersion != null) {
				caps.setCapability("os", os);
				if (os.toLowerCase().startsWith("win")) {
					caps.setCapability("os", "Windows");
				} else if (os.toLowerCase().startsWith("mac-") || os.toLowerCase().startsWith("os x-")) {
					caps.setCapability("os", "OS X");
				}

				if (os.equalsIgnoreCase("win7")) {
					browserStackOSVersion = "7";
				} else if (os.equalsIgnoreCase("win8")) {
					browserStackOSVersion = "8";
				} else if (os.equalsIgnoreCase("win8.1") || os.equalsIgnoreCase("win8_1")) {
					browserStackOSVersion = "8.1";
				} else if (os.toLowerCase().startsWith("mac-") || os.toLowerCase().startsWith("os x-")) {
					browserStackOSVersion = os.split("-")[1];
				}
				caps.setCapability("os_version", browserStackOSVersion);
			}
			caps.setCapability("resolution", "1920x1080");
			caps.setCapability("browserstack.debug", "true");
			caps.setCapability("build", System.getProperty("Build"));
			caps.setCapability("project", System.getProperty("Suite"));
			caps.setCapability("name", currentTestName);

			try {
				driver = new RemoteWebDriver(new URL("http://" + config.getBrowserStackUserName() + ":"
						+ config.getBrowserStackAuthKey() + "@hub.browserstack.com/wd/hub"), caps);
				((RemoteWebDriver) driver).setFileDetector(new LocalFileDetector());
			} catch (Exception e) {
				autoLogger("Issue creating new driver instance due to following error: " + e.getMessage() + "\n"
						+ e.getStackTrace());
				throw e;
			}

			break;

		case "seleniumgrid":
			autoLogger("Remote execution set up on URL: " + config.getRemoteGridUrl());
			remoteGridUrl = config.getRemoteGridUrl();
			caps.setBrowserName("chrome");
			caps.setPlatform(Platform.LINUX);
			String url = "http://" + remoteGridUrl + ":4444/wd/hub";
			autoLogger("===================================" + "\n" + "URL:" + url);
			driver = new RemoteWebDriver(new URL(url), caps);
			((RemoteWebDriver) driver).setFileDetector(new LocalFileDetector());
			break;

		case "local":

			if (config.getBrowserName().equalsIgnoreCase("GCH") || config.getBrowserName().equalsIgnoreCase("chrome")) {
				WebDriverManager.chromedriver().setup();
				ChromeOptions options = new ChromeOptions();
				options.addArguments("--test-type");
				options.addArguments("--disable-extensions");
				options.addArguments("--start-maximized");
				options.addArguments("--remote-allow-origins=*");

				// options.addArguments("--headless");
				// options.addArguments("--verbose");
				// options.addArguments("--no-sandbox");
				// options.addArguments("--test-type");
				// options.addArguments("--disable-extensions");
				// options.addArguments("--start-maximized");
				// options.addArguments("--remote-allow-origins=*");

				driver = new ChromeDriver(options);
			} else if (config.getBrowserName().equalsIgnoreCase("firefox")) {
				WebDriverManager.firefoxdriver().setup();
				// System.setProperty("webdriver.firefox.profile", "default");
				driver = new FirefoxDriver();
			}

			else {
				WebDriverManager.safaridriver().setup();
				driver = new SafariDriver();

			}
			break;
		}

		// driver.manage().timeouts().implicitlyWait(GLOBALTIMEOUT, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(GLOBALTIMEOUT));
		driver.manage().window().maximize();
		return driver;
	}

	public void launchApp() {
		// Delete cookies and Launch the Application
		driver.manage().deleteAllCookies();
		baseUrlSolo = config.getURLSolo();
		driver.get(baseUrlSolo);

		// Maximize the browser
		driver.manage().window().maximize();
		waitForPageToLoad();
	}

	public void launchAppDirectory() {
		// Delete cookies and Launch the Application
		driver.manage().deleteAllCookies();
		baseUrlDirectory = config.getURLDirectory();
		driver.get(baseUrlDirectory);

		// Maximize the browser
		driver.manage().window().maximize();
		waitForPageToLoad();
	}

	public void setBaseUrlSolo(String baseUrlSolo) {
		this.baseUrlSolo = baseUrlSolo;
	}

	public void setBaseUrlDirectory(String baseUrlDirectory) {
		this.baseUrlDirectory = baseUrlDirectory;
	}

	public void launchApp(String url) {
		// Delete cookies and Launch the Application
		driver.manage().deleteAllCookies();

		driver.get(url);
		waitForPageToLoad();

		// Maximize the browser
		driver.manage().window().maximize();
	}

	public WebDriver getCurrentDriverInstance() {
		return driver;
	}

	public void closeBrowser() {
		if (driver != null)
			driver.quit();
	}

	public By getLocatorBy(String locator) {
		By locatorBy = null;
		String string = locator;
		String[] parts = string.split(":-:");
		String type = parts[0]; // 004
		String object = parts[1];

		if (type.equals("id")) {
			locatorBy = By.id(object);
		} else if (type.equals("name")) {
			locatorBy = By.name(object);
		} else if (type.equals("class")) {
			locatorBy = By.className(object);
		} else if (type.equals("link")) {
			locatorBy = By.linkText(object);
		} else if (type.equals("partiallink")) {
			locatorBy = By.partialLinkText(object);
		} else if (type.equals("css")) {
			locatorBy = By.cssSelector(object);
		} else if (type.equals("xpath")) {
			locatorBy = By.xpath(object);
		} else {
			autoLogger("Please provide correct element locating strategy" + locator);
			throw new RuntimeException("Please provide correct element locating strategy" + locator);
		}
		return locatorBy;
	}

	public WebElement getElement(String locatorString) {
		By locatorBy = null;
		int counter = 0;
		// driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(2));
		autoLogger("Finding element using: " + locatorString);
		locatorBy = getLocatorBy(locatorString);

		for (; counter < 5; counter++) {
			try {
				autoLogger("Finding element, try" + counter + " with locator:" + locatorString);
				return driver.findElement(locatorBy);
			} catch (Exception e) {
				System.out.println(e);
				sleep(2000);
			}
		}

		throw new RuntimeException("Element not found: " + locatorString);

	}

	public WebElement findElement(String locatorString) {
		WebElement element = getElement(locatorString);
		return element;
	}

	public WebElement getElementByXpath(String locatorString) {

		int counter = 0;
		// driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(2));
		autoLogger("Finding element using: xpath:" + locatorString);

		for (; counter < 20; counter++) {
			try {
				autoLogger("Finding element, try" + counter + " with locator: xpath:" + locatorString);
				return driver.findElement(By.xpath(locatorString));
			} catch (Exception e) {
				sleep(2000);
			}
		}

		throw new RuntimeException("Element not found: xpath:" + locatorString);

	}

	public WebElement findElementByXpath(String locatorString) {
		WebElement element = getElementByXpath(locatorString);
		return element;
	}

	public List<WebElement> findElements(String locatorString) {
		By locatorBy = null;
		List<WebElement> elements = null;
		locatorBy = getLocatorBy(locatorString);
		elements = driver.findElements(locatorBy);
		return elements;
	}

	public void selectElement(WebElement element, String option) throws Exception {
		Select select = new Select(element);
		select.selectByVisibleText(option);
	}

	public boolean syncProgress() throws Exception {

		int loadCounter = 10;

		// driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		while (loadCounter > 0) {

			try {
				driver.findElement(
						By.xpath("//div[@id='MainContent_panelUpdateProgress'][contains(@style, 'display: none')]"));
				// driver.manage().timeouts().implicitlyWait(GLOBALTIMEOUT, TimeUnit.SECONDS);
				driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(GLOBALTIMEOUT));
				return true;
			} catch (Exception e) {
				loadCounter--;
			}

		}

		// driver.manage().timeouts().implicitlyWait(GLOBALTIMEOUT, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(GLOBALTIMEOUT));
		throw new Exception("Progress was not completed withing specified time");

	}

	public void selectByPartOfVisibleText(WebElement element, String value) {
		boolean flag = true;
		List<WebElement> optionElements = element.findElements(By.tagName("option"));
		Select select = new Select(element);
		for (WebElement optionElement : optionElements) {
			if (optionElement.getText().contains(value)) {
				String optionIndex = optionElement.getAttribute("index");
				select.selectByIndex(Integer.parseInt(optionIndex));
				flag = false;
				break;
			}
		}
		if (flag) {
			Assert.assertTrue("Option " + value + " was not found in the select", false);
		}
	}

	public void verifyElement(String locatorString, boolean checkVisibility) throws Exception {
		if (checkVisibility) {
			boolean isDisplayed = (findElement(locatorString).isDisplayed());
			if (isDisplayed == false) {
				throw new Exception("Element Not Visible Exception");
			}
		} else {
			// driver.manage().timeouts().implicitlyWait(timeOutInSeconds,
			// TimeUnit.SECONDS);
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(GLOBALTIMEOUT));
			findElement(locatorString);
		}
		// driver.manage().timeouts().implicitlyWait(GLOBALTIMEOUT, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(GLOBALTIMEOUT));
	}

	public void sleep(long milliSeconds) {
		try {
			Thread.sleep(milliSeconds);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}

	public String getCurrentSessionID() {
		return currentSessionID;
	}

	public void waitForPageToLoad() {
		new WebDriverWait(driver, Duration.ofSeconds(GLOBALTIMEOUT)).until(webDriver -> ((JavascriptExecutor) webDriver)
				.executeScript("return document.readyState").equals("complete"));

	}

	public void selectDeselectCheckBox(String locator, boolean selectCheckBox) {

		if (selectCheckBox) {
			if (!findElement(locator).isSelected())
				findElement(locator).click();
		} else if (findElement(locator).isSelected())
			findElement(locator).click();
	}

	public void clickElement(String locator) throws Exception {

		int i = 0;
		do {
			try {
				findElement(locator).click();
				break;
			} catch (Exception e) {
				sleep(1000);
				i++;
				continue;
			}

		} while (i < 2);

		if (i >= 2) {
			throw new Exception("Failed to click element, Locator: " + locator);
		}
	}

	public void clickElementWithJs(String locator) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		By by = getLocatorBy(locator);
		WebElement element = driver.findElement(by);
		js.executeScript("arguments[0].click();", element);
	}

	public void enterTextByJs(String locator, String textToEnter) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		By by = getLocatorBy(locator);
		WebElement element = driver.findElement(by);

		sleep(2000);
		// Click on the element
		js.executeScript("arguments[0].click();", element);

		// Clear existing value
		js.executeScript("arguments[0].value='';", element);

		// Enter the new value
		js.executeScript("arguments[0].value='" + textToEnter + "';", element);
	}

	public void enterText(String locator, String text) throws Exception {
		WebElement element = findElement(locator);
		element.click();
		element.clear();
		element.sendKeys(text);
	}

	public boolean verifyCheckBox(String locator) {
		return findElement(locator).isSelected();
	}

	public void waitForNavigation(String url) {
		int counter = 10;
		for (; counter > 0; counter--) {
			if (driver.getCurrentUrl().contains(url)) {
				break;
			} else {
				sleep(10000);
			}
		}
	}

	// public static void staticmmmm(String url) {
	// int counter = 10;
	// for (; counter > 0; counter--) {
	// if (driver.getCurrentUrl().contains(url)) {
	// break;
	// }
	// }
	// }

	public void switchToWindow(int windowNo) {
		Set<String> set = driver.getWindowHandles();
		String windowHandle = null;
		autoLogger("Current no. of windows are: " + set.size());
		if (windowNo <= set.size()) {
			ArrayList<String> windows = new ArrayList<String>(set);
			windowHandle = windows.get(windowNo - 1);
		}

		if (windowHandle != null) {
			driver.switchTo().window(windowHandle);
		} else {
			throw new RuntimeException("Specified window not available");
		}
	}

	public String getFormattedDate() {
		return getDate().replaceAll("/", "_").replaceAll(":", "_").replaceAll(" ", "_");
	}

	public String getDate() {
		DateFormat dateFormat = new SimpleDateFormat("MM/dd HH:mm:ss");
		Date date = new Date();
		autoLogger(dateFormat.format(date));
		return dateFormat.format(date);
	}

	public boolean waitTillElementLoaded(String locator) {
		// driver.manage().timeouts().implicitlyWait(1, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(2));
		int counter = 5;
		do {
			try {
				if (findElement(locator) != null) {
					return true;
				} else {
					sleep(1000);
					counter--;
				}
			} catch (Exception e) {
				sleep(3000);
				counter--;
				continue;
			}
		} while (counter > 0);
		// driver.manage().timeouts().implicitlyWait(GLOBALTIMEOUT, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(GLOBALTIMEOUT));
		throw new RuntimeException("element was not loaded:" + locator);
	}

//	public boolean waitTillElementLoadedonUi(String locator) {
//		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(GLOBALTIMEOUT));
//		try {
//			WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(getLocatorBy(locator)));
//			return element != null && element.isEnabled();
//		} catch (TimeoutException e) {
//			throw new RuntimeException("Element was not loaded within timeout for locator: " + locator, e);
//		}
//	}

	public boolean waitTillElementLoadedonUi(String locator) {
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(GLOBALTIMEOUT));
		final int MAX_RETRY = 3;
		for (int attempt = 1; attempt <= MAX_RETRY; attempt++) {
			try {
				WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(getLocatorBy(locator)));
				if (element != null && element.isEnabled()) {
					return true;
				}
			} catch (StaleElementReferenceException e) {
				if (attempt == MAX_RETRY) {
					throw new RuntimeException("Element is not attached to the page document for locator: " + locator,
							e);
				}
				// Wait a bit before retrying
				try {
					Thread.sleep(500);
				} catch (InterruptedException ie) {
					Thread.currentThread().interrupt();
					throw new RuntimeException("Thread was interrupted during sleep between retries", ie);
				}
			} catch (TimeoutException e) {
				throw new RuntimeException("Element was not loaded within timeout for locator: " + locator, e);
			}
		}
		return false;
	}

	public static void autoLogger(String message) {
		Reporter.log(message, true);
	}

	public void getScreenshot(String name) throws IOException {
		driver = getCurrentDriverInstance();
		String path = "screenshots/" + name;
		File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		FileUtils.copyFile(src, new File(path));
		autoLogger("screenshot at :" + path);
		autoLogger("screenshot for " + name + " available at :" + path);
	}

	public String[][] readExcel(String excelFilePath, int sheetNo) throws BiffException, IOException {
		File file = new File(excelFilePath);
		String inputData[][] = null;
		Workbook w;

		try {
			w = Workbook.getWorkbook(file);

			// Get the sheet
			Sheet sheet = w.getSheet(sheetNo);

			int colcount = sheet.getColumns();
			int rowcount = sheet.getRows();
			int countYes = 0;

			for (int i = 0; i < rowcount; i++) {
				if (sheet.getCell(colcount - 1, i).getContents().equalsIgnoreCase("Yes")) {
					countYes = countYes + 1;

				}
			}

			inputData = new String[countYes][colcount];
			int k = 0;
			for (int i = 0; i < rowcount; i++) {
				if (sheet.getCell(colcount - 1, i).getContents().equalsIgnoreCase("Yes")) {

					for (int j = 0; j < colcount; j++) {
						Cell cell = sheet.getCell(j, i);
						inputData[k][j] = cell.getContents().trim();

					}
					k = k + 1;
				}

			}

		} catch (BiffException e) {
			e.printStackTrace();
			throw e;
		} catch (IOException e) {
			e.printStackTrace();
			throw e;
		}
		return inputData;
	}

	public void scroll(String locator) {

		JavascriptExecutor js = (JavascriptExecutor) driver;
		By by = getLocatorBy(locator);
		WebElement element = driver.findElement(by);
		// This will scroll the page till the element is found
		js.executeScript("arguments[0].scrollIntoView(true);", element);
	}

	public void scrollTop() {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollTo(document.body.scrollHeight,0)");
	}

	public void switchToFrame(String locatorString) {
		driver.switchTo().frame(locatorString);
	}

	public void switchToDefault() {
		driver.switchTo().defaultContent();
	}

	public void handleAlertBox() throws Exception {
		driver.switchTo().alert().accept();
	}

	public void createFile(String path, String filename) throws IOException {
		try {
			File neo = new File(path + File.separator + filename);
			System.out.println(path + File.separator + filename);
			neo.createNewFile();
			neo = null;
		} catch (IOException e) {
			System.out.println("===========File creation failed, Path: " + path + File.separator + filename);
			e.printStackTrace();
			throw e;
		}
	}

	public void writeToFile(String data, String path, String fileName) throws Exception {

		FileWriter fw = null;
		BufferedWriter bw = null;

		try {
			File neo = new File(path + File.separator + fileName);
			fw = new FileWriter(neo.getAbsoluteFile(), true);
			bw = new BufferedWriter(fw);
			bw.write(data);
			bw.newLine();
			bw.close();
			fw.close();
		} catch (Exception e) {
			System.out.println("===========Exception in writing data to file");
			e.printStackTrace();
			throw e;
		} finally {
			if (bw != null) {
				bw.close();
			}
			if (fw != null) {
				fw.close();
			}
		}
	}

	public static void cleanDirectory(String path) {

		try {
			FileUtils.cleanDirectory(new File(path));
		} catch (IOException e) {
			System.out.println("===========Exception in cleaning directory");
			e.printStackTrace();
		}

	}

	public Object[][] readText(String filePath) throws Exception {

		String inputData[][] = null;
		Scanner myReader = null;

		try {
			File myObj = new File(filePath);
			myReader = new Scanner(myObj);
			List<String> fileData = new ArrayList<>();
			String data = null;
			int row = 0;
			int col = 0;

			while (myReader.hasNextLine()) {
				data = myReader.nextLine();
				if (!data.equalsIgnoreCase("")) {
					row++;
					fileData.add(data);
				}
			}

			col = data.split("\\|\\|").length;

			inputData = new String[row][col];
			row = 0;
			for (String string : fileData) {

				String[] lineData = string.split("\\|\\|");
				if (col == lineData.length) {

					for (int j = 0; j < lineData.length; j++) {
						inputData[row][j] = lineData[j];
					}

				} else {
					System.out.println("========column count does not match text file data");
					throw new Exception("column count does not match text file data");
				}

				row++;
			}

			myReader.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} finally {
			myReader.close();
		}

		return inputData;

	}

	public static String[][] readExcel(String excelFilePath, String sheetName) throws BiffException, IOException {
		File file = new File(excelFilePath);
		String[][] inputData = null;
		Workbook workbook = null;

		try {
			workbook = Workbook.getWorkbook(file);
			Sheet sheet = workbook.getSheet(sheetName); // Modified to use sheet name

			int colcount = sheet.getColumns();
			int rowcount = sheet.getRows();
			int countYes = 0;

			// Count rows where last column is "Yes"
			for (int i = 0; i < rowcount; i++) {
				if (sheet.getCell(colcount - 1, i).getContents().equalsIgnoreCase("Yes")) {
					countYes++;
				}
			}

			inputData = new String[countYes][colcount];
			int k = 0;
			for (int i = 0; i < rowcount; i++) {
				if (sheet.getCell(colcount - 1, i).getContents().equalsIgnoreCase("Yes")) {
					for (int j = 0; j < colcount; j++) {
						Cell cell = sheet.getCell(j, i);
						inputData[k][j] = cell.getContents().trim();
					}
					k++;
				}
			}

		} catch (BiffException | IOException e) {
			e.printStackTrace();
			throw e;
		} finally {
			if (workbook != null) {
				workbook.close(); // Properly close the workbook
			}
		}
		return inputData;
	}

	public String generateRandomString(int length) {
		return RandomStringUtils.randomAlphabetic(length);
	}

	public String generateRandomNumber(int length) {
		return RandomStringUtils.randomNumeric(length);
	}

	public String generateRandomNumberEin(int length) {
		return RandomStringUtils.randomNumeric(length);
	}

	public String generateRandomAlphaNumeric(int length) {
		return RandomStringUtils.randomAlphanumeric(length);
	}

	public String generateStringWithAllobedSplChars(int length, String allowdSplChrs) {
		String allowedChars = "abcdefghijklmnopqrstuvwxyz" + // alphabets
				"1234567890" + // numbers
				allowdSplChrs;
		return RandomStringUtils.random(length, allowedChars);
	}

	public String generateEmail(int length) {
		String allowedChars = "abcdefghijklmnopqrstuvwxyz" + "1234567890";
		String email = "";
		String temp = RandomStringUtils.random(length, allowedChars);
		email = temp.substring(0, temp.length()) + "@mailinator.com";
		return email;
	}

	public String generateUrl(int length) {
		String allowedChars = "abcdefghijklmnopqrstuvwxyz" + // alphabets
				"1234567890" + // numbers
				"_-."; // special characters
		String url = "";
		String temp = RandomStringUtils.random(length, allowedChars);
		url = temp.substring(0, 3) + "." + temp.substring(4, temp.length() - 4) + "."
				+ temp.substring(temp.length() - 3);
		return url;
	}

	public boolean isElementPresent(String xpath) {
		try {
			WebElement element = driver.findElement(By.xpath(xpath));
			return element != null;
		} catch (NoSuchElementException e) {
			return false;
		}
	}

	public static LocalDate getNextDate(LocalDate currentDate) {
		return currentDate.plusDays(1); // Adding 1 day to get the next date
	}

	public String getDateXDaysFromNow(int x) {
		DateFormat date = new SimpleDateFormat("MM/dd/yyyy");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, x);
		return date.format(cal.getTime());
	}

	public void waitForElementClickable(WebElement element) {
		new WebDriverWait(driver, Duration.ofSeconds(30)).until(ExpectedConditions.elementToBeClickable(element));
	}

	public void waitForElementVisible(WebElement element) {
		try {
			new WebDriverWait(driver, Duration.ofSeconds(90)).until(ExpectedConditions.visibilityOf(element));
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public void navigateToBack() {
		driver.navigate().back();
	}

	public void refreshPage() {
		driver.navigate().refresh();
	}

	public void scrollIntoView(String locator) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		By by = getLocatorBy(locator);
		WebElement element = driver.findElement(by);

		// First attempt to scroll the element into view
		js.executeScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center', inline: 'nearest'});",
				element);
		System.out.println("Scrolled into View using center alignment");

		// Check if the element is actually in view
		boolean isInView = (Boolean) js.executeScript("var elem = arguments[0],                 "
				+ "  box = elem.getBoundingClientRect(),    " + "  cx = box.left + box.width / 2,         "
				+ "  cy = box.top + box.height / 2,         " + "  e = document.elementFromPoint(cx, cy); "
				+ "for (; e; e = e.parentElement) {         " + "  if (e === elem) return true;           "
				+ "}                                        " + "return false;                            ", element);

		if (!isInView) {
			// Scroll again with different alignment if not in view
			js.executeScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'nearest', inline: 'nearest'});",
					element);
			System.out.println("Adjusted scroll to nearest viewable area");
		}
	}

	public void clickByMouseActions(String locator) {
		try {
			Actions actions = new Actions(driver);
			WebElement element = findElement(locator);
			actions.moveToElement(element).click().build().perform();
			autoLogger("Hovered over element: " + locator);
		} catch (Exception e) {
			autoLogger("Failed to hover over element: " + locator + ". Error: " + e.getMessage());
			throw new RuntimeException(e);
		}
	}

	public String generateRandomTime() {
		Random random = new Random();
		int hour = random.nextInt(12) + 1;
		int minute = (random.nextInt(12)) * 5;
		String formattedMinute = String.format("%02d", minute);
		String amPm = random.nextBoolean() ? "AM" : "PM";
		return hour + ":" + formattedMinute + "-" + amPm;
	}

	public String getRandomFutureDate() {
		// Get the current date
		LocalDate today = LocalDate.now();

		// Generate a random number of days between 1 and 90 (approximately 3 months)
		Random random = new Random();
		int randomDays = random.nextInt(90) + 1;

		// Get the date after random number of days
		LocalDate futureDate = today.plusDays(randomDays);

		// Create a formatter for mm/dd/yyyy format
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");

		// Format the date as per the formatter
		String formattedDate = futureDate.format(formatter);

		return formattedDate;
	}

	public boolean waitTillElementAbsent(String locator) {
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(GLOBALTIMEOUT));
		try {
			return wait.until(ExpectedConditions.invisibilityOfElementLocated(getLocatorBy(locator)));
		} catch (TimeoutException e) {
			throw new RuntimeException("Element was not absent within timeout for locator: " + locator, e);
		}
	}

	public String getRandomData(String category) {
		Faker faker = new Faker();
		String data;

		switch (category.toLowerCase()) {
		case "name":
			do {
				data = faker.name().fullName();
			} while (data.contains("'"));
			return data;

		case "firstname":
			do {
				data = faker.name().firstName();
			} while (data.contains("'"));
			return data;

		case "lastname":
			do {
				data = faker.name().lastName();
			} while (data.contains("'"));
			return data;

		case "address":
			return faker.address().streetAddress();

		case "email":
			return faker.internet().emailAddress();

		case "phonenumber":
			return faker.phoneNumber().phoneNumber();

		case "country":
			return faker.country().name();

		case "currency":
			return faker.currency().name();

		case "phone":
			return faker.phoneNumber().cellPhone();

		case "city":
			return faker.address().cityName();

		case "zipcode":
			return faker.address().zipCode();

		default:
			return "Unknown category";
		}
	}

	public WebElement smartFluentWaitForElementInUi(String locator, int pollingInSeconds) {
		FluentWait<WebDriver> wait = new FluentWait<>(driver).withTimeout(Duration.ofSeconds(GLOBALTIMEOUT))
				.pollingEvery(Duration.ofSeconds(pollingInSeconds))
				.ignoring(NoSuchElementException.class, StaleElementReferenceException.class);

		WebElement element = wait.until(new Function<WebDriver, WebElement>() {
			public WebElement apply(WebDriver driver) {
				try {
					WebElement element = findElement(locator);
					if (element != null && ExpectedConditions.visibilityOf(element).apply(driver) != null
							&& element.isDisplayed()) {
						return element;
					}
				} catch (NoSuchElementException e) {
					return null;
				}
				return null;
			}
		});
		return element;
	}

	public WebElement smartFluentWaitForElementInDOM(String locator, int pollingInSeconds) {
		FluentWait<WebDriver> wait = new FluentWait<>(driver).withTimeout(Duration.ofSeconds(GLOBALTIMEOUT))
				.pollingEvery(Duration.ofSeconds(pollingInSeconds))
				.ignoring(NoSuchElementException.class, StaleElementReferenceException.class);

		WebElement element = wait.until(new Function<WebDriver, WebElement>() {
			public WebElement apply(WebDriver driver) {
				try {
					WebElement element = findElement(locator);
					if (element != null) {
						return element;
					}
				} catch (NoSuchElementException e) {
					return null;
				}
				return null;
			}
		});
		return element;
	}

	public void smartScrollIntoView(String locator, boolean smoothScroll) throws Exception {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		By by = getLocatorBy(locator);
		WebElement element = null;
		String scrollBehavior = smoothScroll ? "smooth" : "instant";

		// Function to safely convert Number objects to Double
		Function<Object, Double> safeDoubleCast = value -> {
			if (value instanceof Double) {
				return (Double) value;
			} else if (value instanceof Long) {
				return ((Long) value).doubleValue();
			} else {
				throw new IllegalArgumentException("Expected a numeric value but received: " + value);
			}
		};

		int attempts = 0;
		boolean isInView = false;

		while (attempts < 10) {
			try {
				element = driver.findElement(by); // Attempt to locate the element each time

				// Check if element is already visible in the viewport
				// isInView = (Boolean) js.executeScript("return
				// arguments[0].isIntersectingViewport()", element);
				isInView = (Boolean) js.executeScript("var elem = arguments[0], box = elem.getBoundingClientRect(), "
						+ "cx = box.left + box.width / 2, cy = box.top + box.height / 2, e = document.elementFromPoint(cx, cy); "
						+ "for (; e; e = e.parentElement) { if (e === elem) return true; } return false;", element);

				if (isInView) {
					System.out.println("Element is already visible in the viewport. No need to scroll.");
					return;
				}

				Double elementTop = safeDoubleCast
						.apply(js.executeScript("return arguments[0].getBoundingClientRect().top", element));
				Double elementLeft = safeDoubleCast
						.apply(js.executeScript("return arguments[0].getBoundingClientRect().left", element));
				Double elementBottom = safeDoubleCast
						.apply(js.executeScript("return arguments[0].getBoundingClientRect().bottom", element));
				Double elementRight = safeDoubleCast
						.apply(js.executeScript("return arguments[0].getBoundingClientRect().right", element));
				Double windowHeight = safeDoubleCast.apply(js.executeScript("return window.innerHeight"));
				Double windowWidth = safeDoubleCast.apply(js.executeScript("return window.innerWidth"));

				// Determine alignment
				String verticalAlign = ((elementTop + elementBottom) / 2 < windowHeight / 2) ? "start" : "end";
				String horizontalAlign = ((elementLeft + elementRight) / 2 < windowWidth / 2) ? "start" : "end";

				js.executeScript("arguments[0].scrollIntoView({behavior: '" + scrollBehavior + "', block: '"
						+ verticalAlign + "', inline: '" + horizontalAlign + "'});", element);
				System.out.println("Attempt " + (attempts + 1)
						+ ": Scrolled into view with dynamic alignment based on size and position");

			} catch (StaleElementReferenceException e) {
				System.out.println("StaleElementReferenceException, retrying... [" + (attempts + 1) + "/3]");
			} catch (Exception e) {
				System.out.println("Error during scroll attempt: " + e.getMessage());
				if (attempts == 2) {
					throw new Exception("Failed to scroll element into view after multiple attempts.", e);
				}
			}
			attempts++;
		}
		throw new Exception(
				"Element is not visible after maximum attempts of scrolling or element is no longer attached to the DOM.");
	}

	public void assertion(String actual, String expected) {
		assertEquals(actual, expected);
	}

	public void assertion(boolean actual, boolean expected) {
		assertEquals(actual, expected);
	}

}
