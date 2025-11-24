package framework_library;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public class Configuration {
	private String urlSolo;
	private String urlDirectory;
	private String browserName;
	private String driverAgent;
	private String executionEnvironment;
	private String browserStackUserName;
	private String browserStackAuthKey;
	private String browserStackBrowserVersion;
	private String browserStackOSVersion;
	private String os;
	private String remoteGridUrl;
	private String browserVersion;
	private String outputPath;

	public Configuration() {
		final Properties prop = new Properties();

		try {
			prop.load(new FileInputStream(new File("config.properties")));
		} catch (IOException e) {
			e.printStackTrace();
		}

		urlSolo = System.getProperty("instanceUrl") != null ? System.getProperty("instanceUrl")
				: prop.getProperty("instanceUrl");
		
		urlDirectory = System.getProperty("instanceUrlDirectory") != null ? System.getProperty("instanceUrlDirectory")
				: prop.getProperty("instanceUrlDirectory");

		browserName = System.getProperty("browserName") != null ? System.getProperty("browserName")
				: prop.getProperty("browserName");

		driverAgent = System.getProperty("driverAgent") != null ? System.getProperty("driverAgent")
				: prop.getProperty("driverAgent");

		executionEnvironment = System.getProperty("executionEnvironment") != null
				? System.getProperty("executionEnvironment")
				: prop.getProperty("executionEnvironment");

		browserStackUserName = System.getProperty("browserstackUsername") != null
				? System.getProperty("browserstackUsername")
				: prop.getProperty("browserstackUsername");

		browserStackAuthKey = System.getProperty("browserstackAuthkey") != null
				? System.getProperty("browserstackAuthkey")
				: prop.getProperty("browserstackAuthkey");

		browserStackBrowserVersion = System.getProperty("browserstackBrowserversion") != null
				? System.getProperty("browserstackBrowserversion")
				: prop.getProperty("browserstackBrowserversion");

		browserStackOSVersion = System.getProperty("browserstackOsversion") != null
				? System.getProperty("browserstackOsversion")
				: prop.getProperty("browserstackOsversion");

		os = System.getProperty("os") != null ? System.getProperty("os") : prop.getProperty("os");

		remoteGridUrl = System.getProperty("remoteGridUrl") != null ? System.getProperty("remoteGridUrl")
				: prop.getProperty("remoteGridUrl");

		outputPath = System.getProperty("outputPath") != null ? System.getProperty("outputPath")
				: prop.getProperty("outputPath");

		AppLibrary.autoLogger("urlSolo:" + urlSolo);
		AppLibrary.autoLogger("urlDirectory:" + urlDirectory);
		AppLibrary.autoLogger("browserName:" + browserName);
		AppLibrary.autoLogger("driverAgent:" + driverAgent);
		AppLibrary.autoLogger("executionEnvironment:" + executionEnvironment);
		AppLibrary.autoLogger("browserStackUserName:" + browserStackUserName);
		AppLibrary.autoLogger("browserStackAuthKey:" + browserStackAuthKey);
		AppLibrary.autoLogger("browserStackBrowserVersion:" + browserStackBrowserVersion);
		AppLibrary.autoLogger("browserStackOSVersion:" + browserStackOSVersion);
		AppLibrary.autoLogger("os:" + os);
		AppLibrary.autoLogger("outputPath:" + outputPath);
	}

	public String getURLSolo() {
		return urlSolo;
	}

	public String getURLDirectory() {
		return urlDirectory;
	}
	public String getBrowserName() {
		return browserName;
	}

	public String getExecutionEnvironment() {
		return executionEnvironment;
	}

	public String getBrowserStackUserName() {
		return browserStackUserName;
	}

	public String getBrowserStackAuthKey() {
		return browserStackAuthKey;
	}

	public String getBrowserStackBrowserVersion() {
		return browserStackBrowserVersion;
	}

	public String getBrowserStackOSVersion() {
		return browserStackOSVersion;
	}

	public String getRemoteGridUrl() {
		return remoteGridUrl;
	}

	public String getOS() {
		return os;
	}

	public String getBrowserVersion() {
		return browserVersion;
	}

	public String getDriverAgent() {
		return driverAgent;
	}

	public String getoutputPath() {
		return outputPath;
	}

}