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
import java.util.LinkedHashMap;
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

	WebDriver driver = null;       //declaring WebDriver object.
	private Logger appLogs = null; //declaring Logger class Object for logging purpose
	WebDriverWait wait = null;     //
	Actions actions = null;        //declaring an Actions API class object
	
	LinkedHashMap<String, HashMap<String, String>> allCommonSheetVars= null;    // declaring a linked HashMap variable
	//public static int parallelCounter = 0;  //counter used to control the parallel processing
	

	/**
	 * Class Constructor
	 * 
	 */
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
			HashMap<String, String> fileVariables , LinkedHashMap<String,HashMap<String, String>> allCommonSheetVariables) {

		//parallelCounter++;
		allCommonSheetVars = allCommonSheetVariables;
		//System.out.println(parallelCounter);
		int status = -1;
		String row = "";
		
		appLogs.debug("Called by thread : "  +Thread.currentThread().getName());
		
		appLogs.info("Launching " + sheetInfo[2] + " browser .....");
		
		try {
		/* lanuch the browser */ 
		launch_browser(sheetInfo[2]);

		appLogs.info("");
		appLogs.info("< ----------- Starting Execution for sheet : " + sheetInfo[1]
				+ " in File : " + sheetInfo[0] + " ----------- >");

		/* for loop to parse through each row in the sheet*/
		for (int i = 0; i < sheetMatrix.size(); i++) {
			
			row = sheetMatrix.get(i).get(6);   //current running row number taken as a string.
			String action = sheetMatrix.get(i).get(2);  //get the utility method form each row.
			
			/* Print the log message based on the currently executing sheet - Main or Common */
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

			
			switch (action) { //Switch to call a particular utility method starts here.

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

			//Not working
			case "assert_query":
				status = 0;
				status = assert_query(sheetInfo[0], sheetMatrix.get(i),
						fileVariables, row, sheetInfo[1]);
				if (status == 0) {
					return 0;
				}
				break;

			case "quit":
				this.quit(sheetInfo[0], sheetInfo[1], 1);
				return 1;
			
			default:
				appLogs.error("Execution Failed in Row : " + row + " --> Method " + action + " in sheet " + sheetInfo[1]
						+ " is not valid.");
				createResultSet(sheetInfo[0], sheetInfo[1], "FAIL", "Method "+ action + " in step : " + row + " in sheet "
						+ sheetInfo[1] + " is not valid.", "N/A");
				quit(sheetInfo[0], sheetInfo[1], 0);
				return 0;
			} //Switch to call each utility method ends here. 
		}
		
		return 1;
		} 
		catch(Exception e){
			//appLogs.info(parallelCounter);
			appLogs.info("Exception Raised by Thread - " + Thread.currentThread().getName());
			createResultSet(sheetInfo[0], sheetInfo[1], "FAIL", e.getMessage().toString(),"N/A");
			return 0;			
		}
		
		finally{
			//parallelCounter--;
		}
		
	}

	
	/**
	 * Launches the brwoser to perform the test.
	 * 
	 * @param browser - the name of the browser to be launched.
	 */
	public void launch_browser(String browser) {
		//WebDriver driver = null;
		
        try{
        	      	
        	//appLogs.info("Inside Launch browser " + Thread.currentThread().getName());
    		if (browser.equalsIgnoreCase("chrome")) {
    			//appLogs.info("Running Chrome browser" + Thread.currentThread().getName());
    			File browserFilePath = new File(System.getProperty("user.dir")+ "//Input//Drivers//chromedriver.exe");
    			System.setProperty("webdriver.chrome.driver", browserFilePath.getAbsolutePath());
    			//System.setProperty("webdriver.chrome.driver",System.getProperty("user.dir")+ "//Input//Drivers//chromedriver.exe");
    			//appLogs.info("Browser launched by " + Thread.currentThread().getName());			
    			driver = new ChromeDriver();
    			
    		}

    		if (browser.equalsIgnoreCase("firefox")) {
    			 driver = new FirefoxDriver();
    		}

    		driver.manage().window().maximize();  //maximize the browser window.
    		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS); // provide implicit wait 
    		wait = new WebDriverWait(driver, 30);
    		actions = new Actions(driver);  //create object for Actions call for each driver instance
    		
        }
		catch(Exception e){
	     appLogs.info("Error is lauching by " + Thread.currentThread().getName() + e.getMessage());
		}
		

	}

	/**
	 * Navigates to a particular URL using WebDriver.get(URL) method
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int goto_url(String fileName, ArrayList<String> data,HashMap<String, String> variables, String stepNumber,
			String currentSheet) {
		String dataParam = "";
		try {
             /* check if there is any missing parameter in current row  */ 
			if (data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}
			
			dataParam = getDataParam(fileName, currentSheet, data, variables, data.get(5));
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
	 * Enters text in a field using WebDriver.sendKeys() method
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of SuccessfullExecution
	 */

	public int enter_text(String fileName, ArrayList<String> data,HashMap<String, String> variables, String stepNumber,String currentSheet) {

		String dataParam = "";
		String locatorValue = "";
		try {
			/* check if there is any missing parameter in current row  */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			Thread.sleep(1000);
			dataParam = getDataParam(fileName, currentSheet, data, variables, data.get(5));
			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));

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
	 * Clears the user input from a web element using WebDriver.clear() method.
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
			
			/* check if there is any missing parameter in current row  */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));
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
	 * Clicks on a particular element using WebDriver.click() method.
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int click(String fileName, ArrayList<String> data,HashMap<String, String> variables, String stepNumber,
			String currentSheet) {

		String locatorValue = "";
		try {
			
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));
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
	 * Clicks on an element using its value attribute, only if that element is present on the page.
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

			/* check if required parameters for this action are missing or not */ 
			if (data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));
			
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
	 * Wait a particular element to appear on the page then click it. 
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int click_wait(String fileName, ArrayList<String> data, HashMap<String, String> variables, String stepNumber,
			String currentSheet) {
		
		String locatorValue = " ";
		
		try {
			
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By
						.xpath(locatorValue)));
				// driver.findElement(By.xpath(data.get("Locator Value"))).click();
				// actions.moveToElement(driver.findElement(By.xpath(data.get("Locator Value")))).click().perform();
				actions.moveToElement(driver.findElement(By.xpath(locatorValue))).click().perform();
				break;

			case "id":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorValue)));
				// driver.findElement(By.id(data.get("Locator Value"))).click();
				actions.moveToElement(driver.findElement(By.id(locatorValue))).click().perform();
				break;

			case "class_name":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.className(locatorValue)));
				// driver.findElement(By.className(data.get("Locator Value"))).click();
				actions.moveToElement(driver.findElement(By.className(locatorValue))).click().perform();
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
	 * Assert that a particular element is displayed on the page based on its text value.
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int assert_text(String fileName, ArrayList<String> data, HashMap<String, String> variables, String stepNumber,
			String currentSheet) {

		String actualValue = "";
		String dataParam = "";
		String locatorValue = "";
		
		try {
			
			/* check if required parameters for this action are missing or not */ 
			
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue =getDataParam(fileName, currentSheet, data, variables, data.get(4));
			dataParam = getDataParam(fileName, currentSheet, data, variables, data.get(5));

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				actualValue = driver.findElement(By.xpath(locatorValue)).getText();
				break;

			case "id":
				actualValue = driver.findElement(By.id(locatorValue)).getText();
				break;

			case "class_name":

				actualValue = driver.findElement(By.className(locatorValue)).getText();
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
	 * assert the presence of an element on the page based on its locator value.
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int assert_presence(String fileName, ArrayList<String> data, HashMap<String, String> variables, String stepNumber,
			String currentSheet) {
		
		String locatorValue = "";
		
		try {
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));

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
	 * Assert that an element is not present on the page based on its locator value 
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
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));
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
						+ " in Step number " + stepNumber + " of Sheet " + currentSheet );
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
	 * Verify the browser alert box on the page and click on 'OK' button. 
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int click_alert_box_ok(String fileName, ArrayList<String> data, HashMap<String, String> variables, String stepNumber,
			String currentSheet) {

		Alert alert = driver.switchTo().alert();

		String dataParam = "";
		String locatorValue = "";

		try {
			/* check if required parameters for this action are missing or not */ 
			if (data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			dataParam = getDataParam(fileName, currentSheet, data, variables, data.get(5));
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
	 * Select a element from the drop down (to be used only with the web elements that have 
	 *  'Select' tag)
	 *  
	 * @param fileName -  Name of the currently executing file
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int select_from_dropdown(String fileName, ArrayList<String> data,HashMap<String, String> variables, String stepNumber,
			String currentSheet) {

		String dataParam = "";
		String locatorValue = "";
		try {
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			Thread.sleep(1000);
			
			dataParam = getDataParam(fileName, currentSheet, data, variables, data.get(5));
			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));

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
	 * Hover the mouse over a web element.
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

			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));
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
	 * Store the text of an element in a variable and add this variable to 'variables' hashmap for further usage.
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
			
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));
			if(data.get(5).toString().substring(0, 5).equals("var <")){
				key = data.get(5).toString().substring(3).trim();
			}
			else{
				incorrectVariableDeclaration(fileName, data, stepNumber, currentSheet);
				return 0;
			}

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
	 * Store the value of an element in a variable and add it to the 'Variables' hashmap for further usage.
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
            
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));
			if(data.get(5).toString().substring(0, 5).equals("var <")){
				key = data.get(5).toString().substring(3).trim();
			}
			else{
				incorrectVariableDeclaration(fileName, data, stepNumber, currentSheet);
				return 0;
			}

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
	 * Store the title attribute of a web Element in a variable and store this in 'Variables'  hashMap for further usage
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
 
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));
			if(data.get(5).toString().substring(0, 5).equals("var <")){
				key = data.get(5).toString().substring(3).trim();
			}
			else{
				incorrectVariableDeclaration(fileName, data, stepNumber, currentSheet);
				return 0;
			}
			

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
	 * Upload a file 
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
		
		/* check if required parameters for this action are missing or not */ 
		if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
			missingArgs(fileName, data, stepNumber, currentSheet);
			return 0;
			
		}

		dataParam = getDataParam(fileName, currentSheet, data, variables, data.get(5));
		locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));
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
			/* check if required parameters for this action are missing or not */ 
			if (data.get(4).equals("N/A") || data.get(5).equals("N/A")) {
				missingArgs(fileName, data, stepNumber, currentSheet);
				return 0;
			}

			locatorValue = getDataParam(fileName, currentSheet, data, variables, data.get(4));
			dataParam = getDataParam(fileName, currentSheet, data, variables, data.get(5));
			
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
	 * Takes the data from a 'Locator Value' and 'Test Data/Options' column and adds and updates them for 
	 * any 'variable' provided in the 'variables' hashmap
	 * 
	 * @param variables - HashMap of the all the variables defined in the config sheet of the file
	 * @param data - contains the String provided in 'Locator Value' or 'Test Data/Options' column of a row.
	 * @return - Returns the updated string value
	 */	
	public String getDataParam(String fileName, String currentSheet, ArrayList<String> data,
			HashMap<String, String> variables, String param) {
		
		String commonSheetKey = fileName+":"+currentSheet+":"+data.get(0)+":"+data.get(7);
		appLogs.debug("Common Sheet keys is " + commonSheetKey);
		

		if(  allCommonSheetVars.get(commonSheetKey)!=null){
			for(String key: allCommonSheetVars.get(commonSheetKey).keySet()){
				
				if (param.indexOf(key) != -1){
					param= param.replace(key, allCommonSheetVars.get(commonSheetKey).get(key));	
					   appLogs.debug("Acutal Data parameter is in allCommonSheetVars Map " + param);
    				    return param.trim();
				}				
			}
		}
			
		for (String key : variables.keySet()) {
			appLogs.debug("Finding data parameter for : " + key + " : "+ variables.get(key));

			if (param.indexOf(key) != -1){
				param= param.replace(key, variables.get(key));
			    appLogs.debug("Acutal Data parameter is in Variables " + param);
			    return param.trim();	
			}
				
		}
		
		appLogs.debug("Acutal Data parameter is " + param);
		return param.trim();

	}

	/**
	 * Adds error message, updates final Result Set in case of any missing arguments in a row.
	 * 
	 * @param fileName - Name of currently Executing File
	 * @param data - Cell data from the currently executing row 
	 * @param stepNumber - Currently executing row number 
	 * @param currentSheet - Currently Executing Sheet
	 */	
	public void missingArgs(String fileName, ArrayList data, String stepNumber,String currentSheet) {
		appLogs.debug("Inside the missingArgs Method");
		
		appLogs.error("Execution Failed: Missing Arugments for the Action " + data.get(2) 
				    + " in Step number " + stepNumber + " of Sheet "+ currentSheet);
		
		createResultSet(fileName, currentSheet, "FAIL","Missing Arugments for the Action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet "+ currentSheet, "N/A");
		
		quit(fileName, currentSheet, 0); //Quit the browser due to Execution Failure
	}
	
	
	/**
	 * Adds error message, updates final Result Set in case of any incorrect Locator Type in a row.
	 * 
	 * @param fileName - Name of currently Executing File
	 * @param data - Cell data from the currently executing row 
	 * @param stepNumber - Currently executing row number 
	 * @param currentSheet - Currently Executing Sheet
	 */	
	public void incorrectLocatorType(String fileName, ArrayList data, String stepNumber,String currentSheet){
		appLogs.debug("Inside the incorrectLocatorType Method");
		
		appLogs.error("Execution Failed: Incorrect Locator Type for the Action "+ data.get(2) 
				    + " in Step number " + stepNumber + " of Sheet "+ currentSheet);
		
		createResultSet(fileName, currentSheet, "FAIL","Incorrect Locator Type for the Action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet "+ currentSheet, "N/A");
		
		quit(fileName, currentSheet, 0);  //Quit the browser due to Execution Failure
		
	}

	/**
	 * Adds error message, updates final Result Set in case of incorrect variable declaration in the 
	 * Test Data/Options Column
	 * 
	 * @param fileName - Name of currently Executing File
	 * @param data - Cell data from the currently executing row 
	 * @param stepNumber - Currently executing row number 
	 * @param currentSheet - Currently Executing Sheet
	 */	
	public void incorrectVariableDeclaration(String fileName, ArrayList data, String stepNumber,String currentSheet){
		
        appLogs.debug("Inside the incorrectVariableAssgnment Method");
		
		appLogs.error("Execution Failed: Incorrect Variable Declaration for the Action "+ data.get(2) 
				    + " in Step number " + stepNumber + " of Sheet "+ currentSheet  + "USE format -  var <variable name>");
		
		createResultSet(fileName, currentSheet, "FAIL","Incorrect Variable Declaration for the Action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet "+ currentSheet + " USE format - var <variable name>", "N/A");
		
		quit(fileName, currentSheet, 0);  //Quit the browser due to Execution Failure
		
	}
	
	/**
	 * Logs error message Set in case of any exceptional condition. 
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
		return e.getMessage().toString();
	}
	

	/**
	 * Quit the browser 
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

		driver.close();
		//driver.quit();
	}

	
	/**
	 * Takes screenshot in case of any failure
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

