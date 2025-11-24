package framework_library;

import org.testng.IRetryAnalyzer;
import org.testng.ITestResult;

public class RetryAnalyzer implements IRetryAnalyzer {
    private int retryCount = 0;
    private static final int MAX_RETRY_LIMIT = 3; // Maximum retries

    @Override
    public boolean retry(ITestResult result) {
        if (retryCount < MAX_RETRY_LIMIT) {
            retryCount++;
            return true; // Tell TestNG to rerun the test
        }
        return false;
    }
}
