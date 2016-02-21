package utilities;

import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.mysql.jdbc.Connection;
import com.mysql.jdbc.Statement;

import core.DriverScript;


public class Utils extends DriverScript {

	WebDriver driver = null;
	// WebDriver driverRef = null;
	private Logger appLogs = null;
	WebDriverWait wait = null;
	Actions actions = null;

	public Utils() {
		// TODO Auto-generated constructor stub
		org.apache.log4j.PropertyConfigurator.configure(System
				.getProperty("user.dir") + "/log4j.properties");
		appLogs = Logger.getRootLogger();

	}

	/*
	 * Method: executor launches the browser Browser name should be provided as
	 * the parameter
	 */

	public int  executor(String[] sheetInfo,ArrayList<ArrayList<String>> sheetMatrix,
			HashMap<String, String> fileVariables) {

		//Utils utils= obj;
		//Utils utils = new Utils();
		int status = -1;
		String row = "";
		/*
		 * appLogs.info(""); appLogs.info("");
		 * appLogs.info("Current Executing File - " + sheetInfo[0]);
		 * appLogs.info("Current Executing Sheet - " + sheetInfo[1]);
		 * appLogs.info("Current Executing browser - " + sheetInfo[2]);
		 * appLogs.info("Current Executing File variables -" +fileVariables);
		 * appLogs.info("Current Executing Sheet Matrix - " +sheetMatrix);
		 * appLogs.info(""); appLogs.info("");
		 */
		appLogs.info("Called by thread : "  +Thread.currentThread().getName());
		
		appLogs.info("Launching " + sheetInfo[2] + " browser .....");
		
		testMethod();
		launch_browser(sheetInfo[2]);

		appLogs.info("");
		appLogs.info("< ------ Starting Execution for sheet : " + sheetInfo[1]
				+ " in File : " + sheetInfo[0] + " ------ >");

		for (int i = 0; i < sheetMatrix.size(); i++) {
			row = sheetMatrix.get(i).get(6);
			String action = sheetMatrix.get(i).get(2);
			if (sheetMatrix.get(i).get(0).equals("Main Sheet")) {
				appLogs.info("Currently Executing Row " + row + " : "+ sheetMatrix.get(i).get(1) + " -- "
						+ " in Main Sheet : " + sheetInfo[1] + " -- in File "
						+ " : " + sheetInfo[0]);
			} else {
				appLogs.info("Currently Executing Row " + row + " : "+ sheetMatrix.get(i).get(1) + " -- "
						+ " in Common Sheet : " + sheetMatrix.get(i).get(0)
						+ " in Main Sheet : " + sheetInfo[1] + " in File "
						+ sheetInfo[0]);
			}

			switch (action) {

			case "goto_url":
				status = 0;
				status = goto_url(sheetInfo[0], sheetMatrix.get(i), fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}

				break;

			case "enter_text":
				status = 0;
				status = enter_text(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			case "click":
				status = 0;
				status = click(sheetInfo[0], sheetMatrix.get(i), fileVariables,
						row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			case "click_wait":
				status = 0;
				status = click_wait(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			case "clear_text":
				status = 0;
				status = clear_text(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			case "assert_presence":
				status = 0;
				status = assert_presence(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			case "assert_no_presence":
				status = 0;
				status = assert_no_presence(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			case "assert_text":

				status = 0;
				status = assert_text(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			// not tested yet
			case "mouse_over":
				status = 0;
				status = mouse_over(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}

				break;

			case "click_button":
				break;

			case "click_using_button_name":
				break;

			case "click_using_button_value":
				break;

			case "click_alert_box_ok":
				status = 0;
				status = click_alert_box_ok(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			case "click_if_exists":
				break;

			case "click_if_exists_button_value":
				status = 0;
				status = click_if_exists_button_value(sheetInfo[0],
						sheetMatrix.get(i), fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			case "click_button_wait":
				break;

			case "select_checkbox_wait":
				break;

			case "select_check_box":
				break;

			case "select_radion_button":
				break;

			case "select_from_dropdown":
				status = 0;
				status = select_from_dropdown(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			// not tested yet
			case "store_text":
				status = 0;
				status = store_text(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			// Not tested yet
			case "store_value":
				status = 0;
				status = store_value(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			// Not tested yet
			case "store_element_title":
				status = 0;
				status = store_element_title(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			case "upload_file":
				status = 0;
				status = upload_file(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				 if (status == 0) {
					return 0;
				}
				break;

			case "assert_query":
				status = 0;
				status = assert_query(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			case "quit":
				quit(sheetInfo[0], sheetInfo[1], 1);

				break;

			default:
				appLogs.error("Execution Failed in Row : " + row
						+ " --> Method " + action + " in sheet " + sheetInfo[1]
						+ " is not valid.");
				createResultSet(sheetInfo[0], sheetInfo[1], "FAIL", "Method "
						+ action + " in step : " + row + " in sheet "
						+ sheetInfo[1] + " is not valid.", "N/A");
				quit(sheetInfo[0], sheetInfo[1], 0);
				return 0;

			}

		}
		return 1;
	}

	/*
	 * Method: launch_browser launches the browser Browser name should be
	 * provided as the parameter
	 */
	public void launch_browser(String browser) {
		//WebDriver driver = null;
		
        try{
        	Thread.sleep(10000);
        	appLogs.info("Inside Launch browser " + Thread.currentThread().getName());
    		if (browser.equalsIgnoreCase("chrome")) {
    			appLogs.info("Running Chrome browser" + Thread.currentThread().getName());
    			System.setProperty("webdriver.chrome.driver",System.getProperty("user.dir")+ "//Input//Drivers//chromedriver.exe");
    			appLogs.info("Browser launched by " + Thread.currentThread().getName());
    			
    			driver = new ChromeDriver();
    			
    		}

    		if (browser.equalsIgnoreCase("firefox")) {
    			 driver = new FirefoxDriver();
    		}

    		driver.manage().window().maximize();
    		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
    		wait = new WebDriverWait(driver, 30);
    		
        }
		catch(Exception e){
	     appLogs.info("Error is lauching by " + Thread.currentThread().getName() + e.getMessage());
		}
		//actions = new Actions(driver);

	}

	/*
	 * Method: goto_urlArguments: HasMap of all the data in one row, Hashmap of
	 * all the variables in the config sheet, current executing row number,
	 * current sheet nameNavigates to the given URL that is present in the Test
	 * Data Options
	 */
	public int goto_url(String fileName, ArrayList<String> data,
			HashMap<String, String> variables, String stepNumber,
			String currentSheet) {
		String dataParam = "";
		try {

			if (data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			dataParam = getDataParam(variables, data.get(5));
			driver.get(dataParam);
			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,
					stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}

	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int enter_text(String fileName, ArrayList<String> data,HashMap<String, String> variables, String stepNumber,String currentSheet) {

		String dataParam = "";
		String locatorValue = "";
		try {

			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			Thread.sleep(1000);
			dataParam = getDataParam(variables, data.get(5));
			locatorValue = getDataParam(variables, data.get(4));

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				driver.findElement(By.xpath(locatorValue)).sendKeys(dataParam);
				break;

			case "id":
				driver.findElement(By.id(locatorValue)).sendKeys(dataParam);
				break;

			case "class_name":
				driver.findElement(By.className(locatorValue)).sendKeys(
						dataParam);
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
			}// switch ends here

			return 1;

		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,
					stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}
	

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int clear_text(String fileName, ArrayList<String> data,HashMap<String, String> variables, String stepNumber,String currentSheet) {
		String locatorValue = "";
		try {

			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(variables, data.get(4));
			Thread.sleep(1000);
			switch (data.get(3)) { // switch starts here.

			case "xpath":
				driver.findElement(By.xpath(locatorValue)).clear();
				;
				break;

			case "id":
				driver.findElement(By.id(locatorValue)).clear();
				break;

			case "class_name":
				driver.findElement(By.className(locatorValue)).clear();
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
			}// switch ends here

			return 1;

		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,	stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int click(String fileName, ArrayList<String> data,
			HashMap<String, String> variables, String stepNumber,
			String currentSheet) {

		String locatorValue = "";
		try {
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(variables, data.get(4));
			Thread.sleep(1000);
			switch (data.get(3)) { // switch starts here.

			case "xpath":
				driver.findElement(By.xpath(locatorValue)).click();
				break;

			case "id":
				driver.findElement(By.id(locatorValue)).click();
				break;

			case "class_name":
				driver.findElement(By.className(locatorValue)).click();
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
			}// switch ends here

			return 1;
		} catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,
					stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}

	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int click_if_exists_button_value(String fileName,
			ArrayList<String> data, HashMap<String, String> variables,
			String stepNumber, String currentSheet) {

		String locatorValue = "";
		try {

			if (data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(variables, data.get(4));
			Thread.sleep(1000);
			if (!driver.findElements(
					By.xpath("//*[@value = '" + locatorValue + "']")).isEmpty()) {
				((WebElement) driver.findElement(By.xpath("//*[@value = '"
						+ locatorValue + "']"))).click();
			}

			else {
				return 1;
			}

			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,
					stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int click_wait(String fileName, ArrayList<String> data,
			HashMap<String, String> variables, String stepNumber,
			String currentSheet) {
		String locatorValue = " ";
		try {
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(variables, data.get(4));

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By
						.xpath(locatorValue)));
				// driver.findElement(By.xpath(data.get("Locator Value"))).click();
				// actions.moveToElement(driver.findElement(By.xpath(data.get("Locator Value")))).click().perform();
				actions.moveToElement(
						driver.findElement(By.xpath(locatorValue))).click()
						.perform();
				break;

			case "id":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By
						.id(locatorValue)));
				// driver.findElement(By.id(data.get("Locator Value"))).click();
				actions.moveToElement(driver.findElement(By.id(locatorValue)))
						.click().perform();
				break;

			case "class_name":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By
						.className(locatorValue)));
				// driver.findElement(By.className(data.get("Locator Value"))).click();
				actions.moveToElement(
						driver.findElement(By.className(locatorValue))).click()
						.perform();
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
			}// switch ends here

			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,
					stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int assert_text(String fileName, ArrayList<String> data,
			HashMap<String, String> variables, String stepNumber,
			String currentSheet) {

		String actualValue = "";
		String dataParam = "";
		String locatorValue = "";
		try {
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")
					|| data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(variables, data.get(4));
			dataParam = getDataParam(variables, data.get(5));

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				actualValue = driver.findElement(By.xpath(locatorValue))
						.getText();
				break;

			case "id":
				actualValue = driver.findElement(By.id(locatorValue)).getText();
				break;

			case "class_name":

				actualValue = driver.findElement(By.className(locatorValue))
						.getText();
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
			}// switch ends here

			if (!(actualValue.equals(dataParam))) {
				appLogs.error("Execution Failed: Actual text on platform does not match with Argument Provided for Action "
						+ data.get(2)+ " in Step number "+ stepNumber+ " of Sheet " + currentSheet);
				appLogs.info("Actual Value :" + actualValue + "-----" + " Exceted Value : " + dataParam);

				getScreenshot(fileName, currentSheet, stepNumber,"Execution Failed at Step " + stepNumber+ " due to data mismatch - "
								+ " Actual Value :" + actualValue + " ----- "+ " Expected Value : " + dataParam);
				return 0;

			}
			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int assert_presence(String fileName, ArrayList<String> data,
			HashMap<String, String> variables, String stepNumber,
			String currentSheet) {
		String locatorValue = "";
		try {
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(variables, data.get(4));

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				driver.findElement(By.xpath(locatorValue));
				break;

			case "id":
				driver.findElement(By.id(locatorValue));
				break;

			case "class_name":
				driver.findElement(By.className(locatorValue));
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
			}// switch ends here

			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,
					stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int assert_no_presence(String fileName, ArrayList<String> data,HashMap<String, String> variables, String stepNumber,String currentSheet) {
		String locatorValue = "";
		List numberOfElements = new ArrayList();
		try {
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(variables, data.get(4));
			switch (data.get(3)) { // switch starts here.

			case "xpath":
				numberOfElements = driver.findElements(By.xpath(locatorValue));
				break;

			case "id":
				numberOfElements = driver.findElements(By.id(locatorValue));
				break;

			case "class_name":
				numberOfElements = driver.findElements(By.className(locatorValue));
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
			}// switch ends here

			if (numberOfElements.size() > 0) {

				appLogs.error("Execution Failed :  For Action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet "
						+ currentSheet);
				getScreenshot(fileName, currentSheet, stepNumber, "Execution Failed :  For Action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet "
						+ currentSheet );
				return 0;
			} 
			else 
				return 1;

		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int click_alert_box_ok(String fileName, ArrayList<String> data,
			HashMap<String, String> variables, String stepNumber,
			String currentSheet) {

		Alert alert = driver.switchTo().alert();

		String dataParam = "";
		String locatorValue = "";

		try {
			if (data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			dataParam = getDataParam(variables, data.get(5));
			Thread.sleep(1000);

			if (!(alert.getText().equals(dataParam))) {
				appLogs.error("Execution Failed: Alert Text does not match with Argument Provided for Action "
						+ data.get(2)+ " in Step number "+ stepNumber+ " of Sheet " + currentSheet);
				appLogs.info("Actual Value :" + alert.getText() + "-----"+ " Exceted Value : " + dataParam);

				getScreenshot(fileName, currentSheet, stepNumber,"Execution Failed at Step " + stepNumber
								+ " due to in correct alert pop up message - "+ " Actual Value :" + alert.getText()
								+ " ----- " + " Expected Value : " + dataParam);
				return 0;
			}

			alert.accept(); // Clicks on OK button in the Alert window.

			return 1;

		} catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int select_from_dropdown(String fileName, ArrayList<String> data,
			HashMap<String, String> variables, String stepNumber,
			String currentSheet) {

		String dataParam = "";
		String locatorValue = "";
		try {
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			Thread.sleep(1000);
			dataParam = getDataParam(variables, data.get(5));
			locatorValue = getDataParam(variables, data.get(4));

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				new Select(driver.findElement(By.xpath(locatorValue))).selectByVisibleText(dataParam);
				break;

			case "id":
				new Select(driver.findElement(By.id(locatorValue))).selectByVisibleText(dataParam);
				break;

			case "class_name":
				new Select(driver.findElement(By.className(locatorValue))).selectByVisibleText(dataParam);
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
			}// switch ends here

			return 1;
		} 
		
		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int mouse_over(String fileName, ArrayList<String> data,HashMap<String, String> variables, String stepNumber, String currentSheet) {

		String locatorValue = "";
		try {
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(variables, data.get(4));
			switch (data.get(3)) { // switch starts here.

			case "xpath":
				actions.moveToElement(driver.findElement(By.xpath(locatorValue))).build().perform();
				break;

			case "id":
				actions.moveToElement(driver.findElement(By.id(locatorValue))).build().perform();
				break;

			case "class_name":
				actions.moveToElement(driver.findElement(By.className(locatorValue))).build().perform();
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
			}// switch ends here

			return 1;

		} catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}

	
	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int store_text(String fileName, ArrayList<String> data,HashMap<String, String> variables, String stepNumber,String currentSheet) {
		String key = null;
		String value = null;
		String locatorValue = "";
		try {

			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(variables, data.get(4));
			key = data.get(5).toString().substring(3).trim();

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				value = driver.findElement(By.xpath(locatorValue)).getText();
				break;

			case "id":
				value = driver.findElement(By.id(locatorValue)).getText();
				break;

			case "class_name":
				value = driver.findElement(By.className(locatorValue)).getText();
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
			}// switch ends here

			variables.put(key, value);

			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}

	
	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int store_value(String fileName, ArrayList<String> data,HashMap<String, String> variables, String stepNumber,String currentSheet) {
		String key = null;
		String value = null;
		String locatorValue = "";
		try {

			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(variables, data.get(4));
			key = data.get(5).toString().substring(3).trim();

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				value = driver.findElement(By.xpath(locatorValue)).getAttribute("value");
				break;

			case "id":
				value = driver.findElement(By.id(locatorValue)).getAttribute("value");
				break;

			case "class_name":
				value = driver.findElement(By.className(locatorValue)).getAttribute("value");
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
			}// switch ends here

			variables.put(key, value);

			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int store_element_title(String fileName, ArrayList<String> data, HashMap<String, String> variables, String stepNumber,String currentSheet) {
	
		String key = null;
		String value = null;
		String locatorValue = "";
		try {

			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(variables, data.get(4));
			key = data.get(5).toString().substring(3).trim();

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				value = driver.findElement(By.xpath(locatorValue)).getAttribute("title");
				break;

			case "id":
				value = driver.findElement(By.id(locatorValue)).getAttribute("title");
				break;

			case "class_name":
				value = driver.findElement(By.className(locatorValue)).getAttribute("title");
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
			}// switch ends here

			variables.put(key, value);

			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int upload_file(String fileName, ArrayList<String> data,	HashMap<String, String> variables, String stepNumber,String currentSheet) {

		String dataParam = "";
		String locatorValue = " ";
		
		if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
			missingArgs(fileName, data, stepNumber, currentSheet);
			return 0;
			
		}

		dataParam = getDataParam(variables, data.get(5));
		locatorValue = getDataParam(variables, data.get(4));
		JavascriptExecutor js = (JavascriptExecutor) driver;
		
		try{
			
			switch (data.get(3)) { // switch starts here.

			case "xpath":
				js.executeScript("document.getElementByXpath('"+locatorValue+"').style.display= 'block';");
				driver.findElement(By.xpath(locatorValue)).sendKeys(dataParam);
				//js.executeScript("document.getElementByXpath('"+locatorValue+"').style.display = 'none';");

				break;

			case "id":
				js.executeScript("document.getElementById('"+locatorValue+"').style.display= 'block';");
				driver.findElement(By.id(locatorValue)).sendKeys(dataParam);
				//js.executeScript("document.getElementById('"+locatorValue+"').style.display = 'none';");

				break;

			case "class_name":
				js.executeScript("document.getElementByClassName('"+locatorValue+"').style.display= 'block';");
				driver.findElement(By.className(locatorValue)).sendKeys(dataParam);
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(fileName, data, stepNumber, currentSheet);
				return 0;
				
			}// switch ends here
		
			return 1;
			
		}
		catch(Exception e){
			String exceptionMessage = exceptionalCondition(fileName, data,stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;				
		}
		
	}

	/**
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int assert_query(String fileName, ArrayList<String> data,HashMap<String, String> variables, String stepNumber,String currentSheet) {

		String URL = "jdbc:mysql://mysql-proxy.brightedge.com:13700/optiweber2";
		String userName = "readyonly";
		String password = "granted";
		String query = "";
		String dataParam = "";
		String locatorValue = "";

		try {

			if (data.get(4).equals("N/A") || data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(variables, data.get(4));
			dataParam = getDataParam(variables, data.get(5));
			query = data.get(4).toString();
			appLogs.info("Executing query -  " + query);

			/* Creating a new connection */
			Connection con = (Connection) DriverManager.getConnection(URL,
					userName, password);

			/* Creating a new statement object */
			Statement stmt = (Statement) con.createStatement();

			/* Execute the SQL Query to generate the result set */
			ResultSet resultSet = stmt.executeQuery(query);

			while (resultSet.next()) {
				if (!(resultSet.getString(1).equals(dataParam))) {
					appLogs.error("Execution Failed: " + dataParam
							+ " is not present in the data base in "
							+ stepNumber + " of Sheet " + currentSheet);
					// main.resultDataRow.add("Execution Failed at Step " +
					// stepNumber + " due to missing arguments.");
					// main.resultDataRow.add("N/A");
					return 0;
				}

				else {
					appLogs.info("Data Matched successfully");
				}

			}

			return 1;
		}

		catch (SQLException e) {
			appLogs.error("Execution Failed :  For Action " + data.get(2)
					+ " in Step number " + stepNumber + " of Sheet "
					+ currentSheet);
			appLogs.error(e.getMessage());
			getScreenshot(fileName, currentSheet, stepNumber, e.getMessage()
					.substring(0, 50) + "...");

			return 0;
		} catch (Exception e) {
			String exceptionMessage = exceptionalCondition(fileName, data,
					stepNumber, currentSheet, e);
			getScreenshot(fileName, currentSheet, stepNumber, exceptionMessage);
			return 0;
		}

	}

	/**
	 * 
	 * @param variables - HashMap of the all the variables defined in the config sheet of the file
	 * @param data - contains the String provided in 'Locator Value' or 'Test Data/Options' column of a row.
	 * @return - Returns the updated string value
	 */	
	public String getDataParam(HashMap<String, String> variables, String data) {

		for (String key : variables.keySet()) {
			appLogs.debug("Finding data parameter for : " + key + " : "+ variables.get(key));

			if (data.indexOf(key) != -1)
				data = data.replace(key, variables.get(key));
		}
		appLogs.debug("Acutal Data parameter is " + data);

		return data.trim();

	}

	/**
	 * 
	 * @param fileName - Name of currently Executing File
	 * @param data - Cell data from the currently executing row 
	 * @param stepNumber - Currently executing row number 
	 * @param currentSheet - Currently Executing Sheet
	 */	
	public void missingArgs(String fileName, ArrayList data, String stepNumber,
			String currentSheet) {
		appLogs.debug("Inside the missingArgs Method");
		appLogs.error("Execution Failed: Missing Arugments for the Action "
		        + data.get(2) + " in Step number " + stepNumber + " of Sheet "
				+ currentSheet);
		createResultSet(fileName, currentSheet, "FAIL","Missing Arugments for the Action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet "
						+ currentSheet, "N/A");
		quit(fileName, currentSheet, 0); //Quit the browser due to Execution Failure
	}
	
	
	/**
	 * 
	 * @param fileName - Name of currently Executing File
	 * @param data - Cell data from the currently executing row 
	 * @param stepNumber - Currently executing row number 
	 * @param currentSheet - Currently Executing Sheet
	 */	
	public void incorrectLocatorType(String fileName, ArrayList data, String stepNumber,String currentSheet){
		appLogs.debug("Inside the incorrectLocatorType Method");
		appLogs.error("Execution Failed: Incorrect Locator Type for the Action "
				+ data.get(2) + " in Step number " + stepNumber + " of Sheet "
				+ currentSheet);
		createResultSet(fileName, currentSheet, "FAIL","Incorrect Locator Type for the Action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet "+ currentSheet, "N/A");
		
		quit(fileName, currentSheet, 0);  //Quit the browser due to Execution Failure
		
	}

	/**
	 * 
	 * @param fileName - Name of currently Executing File
	 * @param data - Cell data from the currently executing Sheet
	 * @param stepNumber - Currently executing row number 
	 * @param currentSheet - Currently Executing Sheet
	 * @param e - Exception class object
	 * @return - returns to the Exception message
	 */

	public String exceptionalCondition(String fileName, ArrayList<String> data,String stepNumber, String currentSheet, Exception e) {
		appLogs.debug("Inside the exceptionalCondition method");
		appLogs.error("Exception Raised :  For Action " + data.get(2)+ " in Step number " + stepNumber + " of Sheet" + currentSheet);
		
		appLogs.error(e.getMessage());
		return e.getMessage();
	}
	

	/**
	 * 
	 * @param fileName - Name of currently Executing File
	 * @param stepNumber - Currently executing row number 
	 * @param sheetName - Currently Executing Sheet
	 * @param status - Status of the currently executing sheet for which quit is called.(0-FAIL , 1- PASS)
	 * @return - N/A
	 */
	public void quit(String fileName, String sheetName, int status) {

		if (status == 1) {
			appLogs.info("< ------ Successfully Completed the execution for sheet : "+ sheetName + " in File : " + fileName + "------ >");
			createResultSet(fileName, sheetName, "PASS", "N/A", "N/A");
		} else
			appLogs.error("< ------ Executon failed for sheet :" + sheetName + " ------ >");

		driver.quit();
	}

	
	/**
	 * 
	 * @param fileName - Name of currently Executing File
	 * @param stepNumber - Currently executing row number 
	 * @param sheetName - Currently Executing Sheet
	 * @param exceptionMessage - Exception message that was raised in the current stepNumber
	 * @return - N/A
	 */
	public void getScreenshot(String fileName, String sheetName,String stepNumber, String exceptionMessage) {

		String screenshotName = fileName + "-" + sheetName + "-" + "Step"+ stepNumber;
		File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		try {
			FileUtils.copyFile(scrFile, new File(System.getProperty("user.dir")+ "/Results/Screenshots/" + screenshotName + ".png"));
		} 
		catch (IOException e) {// TODO Auto-generated catch block
			e.printStackTrace();
		}

		createResultSet(fileName, sheetName, "FAIL", exceptionMessage,screenshotName);
		quit(fileName, sheetName, 0);
	}
	
	public void testMethod(){
		appLogs.info("This is test" + Thread.currentThread().getName());
	}

}
