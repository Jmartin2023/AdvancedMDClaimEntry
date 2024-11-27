package rpa;

import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Hashtable;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.io.FileUtils;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.ITestResult;
import org.testng.SkipException;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import org.xml.sax.SAXException;



import objects.ExcelOperations;
import objects.SeleniumUtils;
import objects.Utility;

import objects.ExcelReader;




public class AdvancedMDClaimEntry {
	Logger logger = LogManager.getLogger(AdvancedMDClaimEntry.class);

	String projDirPath,NPI, status, claimNo ,claimNumAvaility, AvailityDOS, denialReason,DOB ,serviceDate ,firstName, lastName,memberID, maximusStatus,DOS, claimStatus,dateofbirth, npivalue, charges,currency, error, originalTab, checkNum,checkDate,paidAmount,paymentDate, receivedDate, allowedAmount, processedDate,finalizedDate;
	Set<String> allWindowHandles;
	String LoginWindow ;
	String secondWindow;
	String chargeWindow;
	

	public static ExcelReader excel, excel1; 
	public static String sheetName = "Sheet1";
	int rowNum = 1;
	boolean skipFlag =false;
	WebDriver driver;
	SimpleDateFormat parser  = new SimpleDateFormat("M/dd/yyyy");
	// output format: yyyy-MM-dd
	
	// output format: yyyy-MM-dd
	SimpleDateFormat formatter = new SimpleDateFormat("MM/dd/yyyy");
	//JavascriptExecutor js;
	SeleniumUtils sel;
	Utility utility;
Boolean firstRun;
	ExcelOperations excelFile;
	WebDriverWait waitExplicit ;
	WebDriverWait wait10;
	static String excelFileName, payer;

	@BeforeTest
	public void preRec() throws InterruptedException, SAXException, IOException, ParserConfigurationException {


		sel = new SeleniumUtils(projDirPath);

		driver = sel.getDriver();
		waitExplicit	= new WebDriverWait(driver, Duration.ofSeconds(50));
		wait10	= new WebDriverWait(driver, Duration.ofSeconds(10));
		//js = (JavascriptExecutor) driver;

		utility = new Utility();

		String[] params = new String[]{"url", "username", "password","excelName"};
		HashMap<String, String> configs = utility.getConfig("config.xml", params);

		String url = configs.get("url"), 
				username = configs.get("username"), 

				password = configs.get("password");

		excelFileName = configs.get("excelName");
		System.out.println(excelFileName);

		driver.get(url);
		logger.info("Open url: " + url);


		driver.switchTo().frame("frame-login");
		sel.pauseClick(driver.findElement(By.id("loginName")), 20);
		driver.findElement(By.id("loginName")).sendKeys(username);
		logger.info("Username entered as: "+username);
		driver.findElement(By.id("password")).sendKeys(password);
		logger.info("password entered");
		driver.findElement(By.id("officeKey")).sendKeys("125092");
		logger.info("officeKey entered");

		LoginWindow = driver.getWindowHandle();
		driver.findElement(By.xpath("//button[contains(text(),'Log in')]")).click();
		logger.info("Login button clicked");

		Thread.sleep(16000);
		allWindowHandles = driver.getWindowHandles();
		//driver.switchTo().defaultContent();
		// Iterate through window handles to find the new window
		for (String windowHandle : allWindowHandles) {
			if (!windowHandle.equals(LoginWindow)) {
				// Switch to the new window
				driver.switchTo().window(windowHandle);
				logger.info("Switched to main window");
				break;
			}
		}
		Thread.sleep(3000);
		waitExplicit.until(ExpectedConditions.elementToBeClickable(By.id("mnuPatientInfo"))).click();
		//   driver.findElement(By.id("mnuPatientInfo")).click();
		logger.info("Clicked on patient info");
		//   Thread.sleep(10000);


		firstRun=true;
	}
	
	
	
	@Test(dataProvider= "getData") 
	public void claimEntry(Hashtable<String,String> data) throws InterruptedException, ParseException {
rowNum++;
		String status  = data.get("Bot Status");
		String name = data.get("Patient Name").trim();
		String DOS= data.get("DOS").trim();
		String admissionDate= data.get("Admission Date").trim();
		
		String DOB = data.get("DOB");
		String providerCode = data.get("Provide Code");
		if(providerCode.isBlank()|| providerCode.isEmpty() || providerCode.contains("-")) {
			excel.setCellData(sheetName, "Bot Status", rowNum, "Fail. Provider Code not present");
			logger.info("CPT not present");
			throw new SkipException("CPT not present");
		}else {
			providerCode = providerCode.trim();
		}
		
		String chartNum = data.get("Chart Number");
		String phone = data.get("Phone");
		String renderingProvider = data.get("Rendering Provider").trim();
		String refProvider = data.get("Referring Provider").trim();
		String facility = data.get("Hospital Name").trim();
		String cpts = data.get("CPTs").replace(".0", "");
		String diagnosis = data.get("DX");
		String[] diagArray;
		if(diagnosis.contains(",")) {
			diagArray = diagnosis.split(",");	
		}
		else{
			 diagArray = diagnosis.split(" ");	
		}
		
		String lastNameProvider= 	removeSingleLetters(refProvider);
		if(DOB.contains("-")) {
			DOB="";
		}else {
			DOB=DOB.trim();
		DOB=	formatter.format(parser.parse(DOB));
		}
		
		if(chartNum.isBlank()|| chartNum.isBlank()|| chartNum.contains("-")) {
			excel.setCellData(sheetName, "Bot Status", rowNum, "Fail. Chart Number not present.");
			logger.info("Chart Number not present");
			throw new SkipException("Chart Number not present");
			
		}else {
			
			chartNum =chartNum.trim().replace(".0", "").trim();
			
		}
		if(admissionDate.isBlank()|| admissionDate.isBlank()|| admissionDate.contains("-")) {
			excel.setCellData(sheetName, "Bot Status", rowNum, "Fail. Admission Date not present.");
			logger.info("Admission Date not present.");
			throw new SkipException("Admission Date not present.");
			
		}
		
		
		
	/*	 firstName = null ;
		 lastName= null ;
		
		if(name.contains(",")) {
			firstName = name.split(",")[0].trim();
			lastName = name.split(",")[1].trim();	
			
		}
		firstName = 	WordUtils.capitalizeFully(firstName);
		lastName = 	WordUtils.capitalizeFully(lastName);
*/
		logger.info("New record");
		
		if(status.isBlank() || status.isEmpty()) {
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(35));
			wait.until(webDriver -> "complete".equals(((JavascriptExecutor) webDriver)
					.executeScript("return document.readyState")));


try {
			waitExplicit.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt("frmPatientInfo1"));
}catch(Exception e) {}
			Thread.sleep(5000);
			


			if(firstRun==true) {
				Thread.sleep(30000);
			}
			firstRun=false;
			
			waitExplicit.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[contains(@id,'mat-input')]")));

			driver.findElement(By.xpath("//input[contains(@id,'mat-input')]")).clear();
			driver.findElement(By.xpath("//input[contains(@id,'mat-input')]")).sendKeys(chartNum);
			
			
			logger.info("patient name entered as: "+name +" | DOB: "+DOB+" | Chart Number: "+chartNum);

	
		
			
try {
	
	logger.info("patient searching in try block");
	wait10.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[contains(text(),'"+DOB+"')]/ancestor::div//div[contains(text(),'"+chartNum+"')]")));

			//span[text()='Wood']/following-sibling::span[contains(text(),'William')]/ancestor::div[@class='row-item']/following-sibling::div[1]/div[contains(text(),'07/27/1965')]
			
			driver.findElement(By.xpath("//div[contains(text(),'"+DOB+"')]/ancestor::div//div[contains(text(),'"+chartNum+"')]")).click();
			logger.info("patient selected");
			
}catch(Exception e) {
	logger.info("patient searching in catch block");
	
	try {
		wait10.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[contains(text(),'"+DOB+"')]/ancestor::div//div/span[contains(text(),'"+chartNum+"')]")));

		//span[text()='Wood']/following-sibling::span[contains(text(),'William')]/ancestor::div[@class='row-item']/following-sibling::div[1]/div[contains(text(),'07/27/1965')]
		
		driver.findElement(By.xpath("//div[contains(text(),'"+DOB+"')]/ancestor::div//div/span[contains(text(),'"+chartNum+"')]")).click();
		logger.info("patient selected");
		
		
	}catch(Exception e1)
	{
		excel.setCellData(sheetName, "Bot Status", rowNum, "Patient not found");
		logger.info("Patient not found");
		throw new SkipException("Skipping this record");
	}
	

}
			
			Thread.sleep(5000);
			//div[@class='additional-item dob-item mr-medium' and text()='12/12/1975']
			
try {
				
				allWindowHandles = driver.getWindowHandles();
				secondWindow = driver.getWindowHandle();
				
				for (String windowHandle : allWindowHandles) {
					if (!windowHandle.equals(LoginWindow) && !windowHandle.equals(secondWindow)) {
						// Switch to the new window
						driver.switchTo().window(windowHandle);
						String title = driver.getTitle();
						driver.close();
						logger.info(title+ " window closed");
						
					}
				}
				
				
			}catch(Exception e) {}
			
driver.switchTo().window(secondWindow);
			driver.switchTo().defaultContent();
			waitExplicit.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt("frmPatientInfo"));
			logger.info("Switched to frame frmPaitentInfo");
			driver.findElement(By.xpath("//span[text()='Transaction Entr']")).click();
			logger.info("Clicked on Transaction entry");

			driver.manage().window().maximize();
			Thread.sleep(3000);
			waitExplicit.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt("frm"));
			logger.info("Switched to frame frm");
			Thread.sleep(3000);
			wait.until(ExpectedConditions.elementToBeClickable(By.id("btnTxCharge")));
			
			
			
			

			WebElement chargeButton = driver.findElement(By.id("btnTxCharge"));
			((JavascriptExecutor) driver).executeScript("arguments[0].click();", chargeButton);
			logger.info("Clicked on charge button using JavaScript");
			Thread.sleep(5000);
			allWindowHandles = driver.getWindowHandles();

			for (String windowHandle : allWindowHandles) {
				if (!windowHandle.equals(LoginWindow) && !windowHandle.equals(secondWindow)&& driver.switchTo().window(windowHandle).getTitle().contains("Charge Entry")) {
					// Switch to the new window
					driver.switchTo().window(windowHandle);
					logger.info("Switched to charge entry window");
					break;
				}
			}

			//driver.switchTo().defaultContent();

			Thread.sleep(2000);
			chargeWindow = driver.getWindowHandle();
			System.out.println(driver.getTitle());
			driver.manage().window().maximize();


			/*
// Print details of each element
List<WebElement> labels = driver.findElements(By.tagName("label"));

// Loop through the list and print the text of each <label>
for (WebElement label : labels) {
    System.out.println(label.getText().trim());
}
			        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='ellProvider']//span"))).click();
			        logger.info("Clicked on provider options");

			
 */
			
			
			
			
			
		Thread.sleep(2000);
		try {
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@id='ellFacility']/input")));
		}catch(Exception e) {
			
			try {
				wait.until(ExpectedConditions.elementToBeClickable(By.id("btnNewBatch"))).click();
				logger.info("Clicked on Begin New Batch");
				Thread.sleep(1500);
				driver.findElement(By.id("btnOK")).click();
				logger.info("Clicked on OK");
			}catch(Exception e1) {}
			
		}
		Thread.sleep(2000);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@id='ellFacility']/input")));
		
		
		((JavascriptExecutor) driver).executeScript("arguments[0].value = arguments[1];", driver.findElement(By.xpath("//div[@id='ellFacility']/input")), facility.split("-")[0].trim());
		driver.findElement(By.xpath("//div[@id='ellFacility']/input")).sendKeys(Keys.TAB);
		
		//	driver.findElement(By.xpath("//div[@id='ellFacility']/input")).sendKeys(facility.split("-")[0].trim()+Keys.TAB)	;
		logger.info("Facility entered as "+facility.split("-")[0].trim());
						Thread.sleep(3500);
		
			
			wait.until(ExpectedConditions.elementToBeClickable(By.id("selEpisode")));
			Select select = new Select(driver.findElement(By.id("selEpisode")));
			Boolean flagAdmissionDate =false;
			for (WebElement option : select.getOptions()) {
	            if (option.getText().contains(admissionDate) || option.getText().contains(admissionDate.replace("/", ""))) {
	            	System.out.println(option.getText().contains(admissionDate) );
	            	System.out.println(option.getText().contains(admissionDate) );
	                option.click(); // Select the option
	                driver.findElement(By.id("selEpisode")).sendKeys(Keys.TAB);
	                System.out.println("Option selected: " + option.getText());
	                flagAdmissionDate=true;
	                break; // Exit loop once the desired option is selected
	            }
			}
			
			if(flagAdmissionDate==false) {
				logger.info("Admission Date Not Available in Application");
				excel.setCellData(sheetName, "Admission Date Comments", rowNum, "Admission Date Not Available in Application");
			}
			

			Thread.sleep(2000);
		try {	
			driver.findElement(By.id("txtBeginDate")).clear();
			driver.findElement(By.id("txtEndDate")).clear();
			
			((JavascriptExecutor) driver).executeScript("arguments[0].value = arguments[1];", driver.findElement(By.id("txtBeginDate")), DOS);
	
			logger.info("Starting DOS entered as "+DOS);
			((JavascriptExecutor) driver).executeScript("arguments[0].value = arguments[1];", driver.findElement(By.id("txtEndDate")), DOS);
			logger.info("Ending DOS entered as "+DOS);
		}catch(Exception e) {
			excel.setCellData(sheetName, "Bot Status", rowNum, "DOS option disabled on Portal");
			driver.close();
			driver.switchTo().window(secondWindow);
			logger.info("Skipping this record, DOS option disabled on Portal");
			throw new SkipException("Skipping this record");
		}
		
			
		
			
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='ellProvider']//span"))).click();
			logger.info("Clicked on provider options");

			Thread.sleep(3000);
			allWindowHandles = driver.getWindowHandles();

			for (String windowHandle : allWindowHandles) {
				if (!windowHandle.equals(LoginWindow) && !windowHandle.equals(secondWindow)&& !windowHandle.equals(chargeWindow)&& driver.switchTo().window(windowHandle).getTitle().contains("Select - Provider")) {
					// Switch to the new window
					driver.switchTo().window(windowHandle);
					logger.info("Switched to provider window");
					break;
				}
			}
			System.out.println(driver.getTitle());
			driver.manage().window().maximize();

			Thread.sleep(4000);
			
			
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//button[@id='btnSearch']/preceding-sibling::input[1]")));
			driver.findElement(By.xpath("//button[@id='btnSearch']/preceding-sibling::input[1]")).clear();
			//driver.findElement(By.xpath("//button[@id='btnSearch']/preceding-sibling::input[1]")).sendKeys(renderingProvider);
			
			
			((JavascriptExecutor) driver).executeScript("arguments[0].value = arguments[1];", driver.findElement(By.xpath("//button[@id='btnSearch']/preceding-sibling::input[1]")), providerCode);
		
			driver.findElement(By.id("rbSearchCriteriaProvCode")).click();
			logger.info("Clicked on Code Option");
			

			logger.info("Provider entered as "+providerCode);
			driver.findElement(By.id("btnSearch")).click();
			driver.findElement(By.id("btnSearch")).click();
			driver.findElement(By.id("btnSearch")).click();
			driver.findElement(By.id("btnSearch")).click();
			logger.info("Search button clicked");
Thread.sleep(2000);
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//nobr[contains(text(),'"+providerCode+"')]")));
			
			driver.findElement(By.xpath("//nobr[contains(text(),'"+providerCode+"')]")).click();
			logger.info("Provider selected");
			driver.findElement(By.id("btnOK")).click();
			logger.info("OK Button clicked");

			driver.switchTo().window(chargeWindow);
			logger.info("Switched to charge window");
			
			Thread.sleep(2000);
			if(!admissionDate.contains("-") && flagAdmissionDate==true) {
				wait.until(ExpectedConditions.elementToBeClickable(By.id("btnExtraInfo")));
				driver.findElement(By.id("btnExtraInfo")).click();
				logger.info("Extra info button clicked");
				
				Thread.sleep(3000);
				allWindowHandles = driver.getWindowHandles();

				for (String windowHandle : allWindowHandles) {
					if (!windowHandle.equals(LoginWindow) && !windowHandle.equals(secondWindow)&&  !windowHandle.equals(chargeWindow)&& driver.switchTo().window(windowHandle).getTitle().contains("Extra Information")) {
						// Switch to the new window
						driver.switchTo().window(windowHandle);
						logger.info("Switched to extra info window");
						break;
					}
				}
				
				String extraInfoWindow = driver.getWindowHandle();
				wait.until(ExpectedConditions.elementToBeClickable(By.id("hospitalizationfrom")));
				((JavascriptExecutor) driver).executeScript("arguments[0].value = arguments[1];", driver.findElement(By.id("hospitalizationfrom")), admissionDate);
				driver.findElement(By.id("hospitalizationfrom")).sendKeys(Keys.TAB);
				
				logger.info("Admission from date entered as: "+ admissionDate);
				
				driver.findElement(By.id("btnSaveClose")).click();
				logger.info("Saven and close button clicked");
				
				try {
					Thread.sleep(2000);
					String alertMsg = driver.switchTo().alert().getText();
					excel.setCellData(sheetName, "Admission Date Comments", rowNum, alertMsg);
					driver.switchTo().alert().dismiss();
					logger.info("Alert Dismissed "+alertMsg);
					driver.switchTo().window(extraInfoWindow);
					logger.info("Switched back to extra info window");
					driver.close();
					logger.info("Closed extra info window");
				}catch(Exception e) {
					logger.info("Admission Date Added");
					excel.setCellData(sheetName, "Admission Date Comments", rowNum, "Admission Date Added");
				}
				
				} 
			
	/*		allWindowHandles = driver.getWindowHandles();

			for (String windowHandle : allWindowHandles) {
				if (!windowHandle.equals(LoginWindow) && !windowHandle.equals(secondWindow)&& !windowHandle.equals(providerWindow)) {
					// Switch to the new window
					driver.switchTo().window(windowHandle);
					laogger.info("Switched to main window");
					break;
				}
			}
			*/
			driver.switchTo().window(chargeWindow);
			logger.info("Switched to charge window");
			
			
			

				
				 wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@id='ellProccode']//input"))).sendKeys(cpts.trim()+Keys.TAB);
					logger.info("CPT entered "+cpts);
					for(int i=0; i<diagArray.length && i<4;i++) {
						
						((JavascriptExecutor) driver).executeScript("arguments[0].value = arguments[1];", driver.findElement(By.xpath("//div[@id='ellDiag10code"+String.valueOf(i+1)+"']//input")), diagArray[i].trim());
						if(i<3) {
							driver.findElement(By.xpath("//div[@id='ellDiag10code"+String.valueOf(i+1)+"']//input")).sendKeys(Keys.TAB);
						}
						
					//	driver.findElement(By.xpath("//div[@id='ellDiag10code"+String.valueOf(i+1)+"']//input")).sendKeys(diagArray[i]);
						logger.info("Diagnosis "+i+" entered "+diagArray[i]);
					}
				Thread.sleep(2000);
				
				
				allWindowHandles = driver.getWindowHandles();

				for (String windowHandle : allWindowHandles) {
					if (!windowHandle.equals(LoginWindow) && !windowHandle.equals(secondWindow)&& !windowHandle.equals(chargeWindow)) {
						// Switch to the new window
						driver.switchTo().window(windowHandle);
						String title = driver.getTitle();
						driver.close();
						logger.info(title+ " window closed");
						
					}
				}
				
				driver.switchTo().window(chargeWindow);
			logger.info("Switched to charge window");

wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//label[text()='ICD-10 Diagnosis Codes']")));
			driver.findElement(By.xpath("//label[text()='ICD-10 Diagnosis Codes']")).click();
				logger.info("Clicked on text: ICD-10 Diagnosis Codes");
			
				
				driver.findElement(By.id("btnAddSave")).click();
				logger.info("Clicked on Add button");
				Thread.sleep(2000);
				
				
				
				
				try {
					
					String alertText = driver.switchTo().alert().getText();
					logger.info(alertText);
					driver.switchTo().alert().dismiss();
					excel.setCellData(sheetName, "Bot Status", rowNum, alertText);
					
					driver.switchTo().window(chargeWindow);
					wait.until(ExpectedConditions.elementToBeClickable(By.id("btnCancel"))).click();
					driver.switchTo().window(secondWindow);
					logger.info("Skipping this record, Charge already in database");
				
					throw new SkipException("Skipping this record, Charge already in database");
					
					
				}catch(Exception e) {}
				
			
				try {
					
					Thread.sleep(3000);
					allWindowHandles = driver.getWindowHandles();
				
					
					for (String windowHandle : allWindowHandles) {
						if (!windowHandle.equals(LoginWindow) && !windowHandle.equals(secondWindow) && !windowHandle.equals(chargeWindow)) {
							// Switch to the new window
							driver.switchTo().window(windowHandle);
							String title = driver.getTitle();
							driver.close();
							logger.info(title+ " window closed");
							
						}
					}
					
					
				}catch(Exception e) {}
				
				driver.switchTo().window(chargeWindow);
				
				
				
			driver.findElement(By.id("btnOK")).click();
			logger.info("OK Button clicked");
			excel.setCellData(sheetName, "Bot Status", rowNum, "Pass");
			driver.switchTo().window(secondWindow);
			
			
		
			
			
			
			
		//	driver.switchTo().alert();
			// System.out.println("Alert text: " +  driver.switchTo().alert().getText());
			 //driver.switchTo().alert().accept();
			
			
		//	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='ellRefprovider']//span"))).click();
			
		/*	allWindowHandles = driver.getWindowHandles();

			for (String windowHandle : allWindowHandles) {
				if (!windowHandle.equals(LoginWindow) && !windowHandle.equals(secondWindow)&& !windowHandle.equals(providerWindow)) {
					// Switch to the new window
					driver.switchTo().window(windowHandle);
					logger.info("Switched to main window");
					break;
				}
			}
			
			System.out.println(driver.getTitle());
			driver.manage().window().maximize();

		
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//button[@id='btnSearch']/preceding-sibling::input[1]")));
			driver.findElement(By.xpath("//button[@id='btnSearch']/preceding-sibling::input[1]")).clear();
			driver.findElement(By.xpath("//button[@id='btnSearch']/preceding-sibling::input[1]")).sendKeys(lastNameProvider.split(" ")[lastNameProvider.split(" ").length-1]);

			logger.info("Provider entered as "+lastNameProvider.split(" ")[lastNameProvider.split(" ").length-1]);
			driver.findElement(By.id("btnSearch")).click();
			driver.findElement(By.id("btnSearch")).click();
			driver.findElement(By.id("btnSearch")).click();
			driver.findElement(By.id("btnSearch")).click();
			logger.info("Search button clicked");
Thread.sleep(2000);
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//nobr[contains(text(),'"+lastNameProvider.split(" ")[lastNameProvider.split(" ").length-1]+"')]")));
			
			driver.findElement(By.xpath("//nobr[contains(text(),'"+lastNameProvider.split(" ")[lastNameProvider.split(" ").length-1]+"')]")).click();
			logger.info("Provider selected");
			Thread.sleep(2000);
			driver.findElement(By.id("btnOK")).click();
			
			logger.info("OK Button clicked");
			*/
			
			
			/*
			Thread.sleep(2000);
			driver.findElement(By.xpath("//div[@id='ellRefprovider']//input")).clear();
			driver.findElement(By.xpath("//div[@id='ellRefprovider']//input")).sendKeys("VAR13" +Keys.TAB);
			logger.info("Success");  */
		//	driver.close();
		//	logger.info("Closed");
	/*		
			driver.switchTo().window(providerWindow);
			logger.info("Switched to main window");
			Thread.sleep(5000);
			
			 driver.switchTo().alert();
			 System.out.println("Alert text: " +  driver.switchTo().alert().getText());
			 driver.switchTo().alert().dismiss();
		        // Get and print the alert text
		       
			 driver.switchTo().window(providerWindow);
				logger.info("Switched to main window");
			 
			

*/
		}





	}


	public boolean waitFunction(WebElement webEle) throws InterruptedException {
		try {
			Thread.sleep(3000);
			webEle.isDisplayed();
			logger.info("Element found:"+ webEle);

		}catch(Exception e) {

			for(int i=0; i<5; i++) {
				Thread.sleep(4000);
				try{ 
					webEle.isDisplayed();


					logger.info("Element found:"+ webEle);

					break;

				}catch(Exception e1) {}
			}
		}
		return webEle.isDisplayed();
	}
	@AfterMethod()
	public void afterMethod(ITestResult result) throws IOException {

		if(!result.isSuccess()) {
			// Test Failed
			String error = result.getThrowable().getLocalizedMessage();
			logger.info(error);
			//result.getThrowable().printStackTrace();
			try {
				TakesScreenshot ts = (TakesScreenshot) driver;
				File ss = ts.getScreenshotAs(OutputType.FILE);
				String ssPath = "./Screenshots/" + result.getName() + " - " + rowNum + ".png";
				FileUtils.copyFile(ss, new File(ssPath));
			} catch (Exception e) {
				System.out.println("Error taking screenshot");
			}

		}
		else {
			logger.info("Test completed successfully");
		}}

	@DataProvider
	public static Object[][] getData(){


		if(excel == null){


			excel = new ExcelReader(System.getProperty("user.dir")+"\\"+excelFileName);


		}


		int rows = excel.getRowCount(sheetName);
		int cols = excel.getColumnCount(sheetName);

		Object[][] data = new Object[rows-1][1];

		Hashtable<String,String> table = null;

		for(int rowNum=2; rowNum<=rows; rowNum++){

			table = new Hashtable<String,String>();

			for(int colNum=0; colNum<cols; colNum++){

				//	data[rowNum-2][colNum]=	excel.getCellData(sheetName, colNum, rowNum);

				table.put(excel.getCellData(sheetName, colNum, 1), excel.getCellData(sheetName, colNum, rowNum));	
				data[rowNum-2][0]=table;	

			}
		}

		return data;

	}
	
	 public static String removeSingleLetters(String input) {
	        if (input == null || input.isEmpty()) {
	            return input;
	        }
	        // Split by spaces
	        input=input.trim().replaceAll(",", "");
			 input=input.trim().replaceAll("MD", "");
			 input=input.trim().replaceAll("\\s+", " ");
	        String[] words = input.split(" ");
	        StringBuilder result = new StringBuilder();

	        for (String word : words) {
	            // Include words that are not single letters
	            if (word.length() > 1 || !word.matches("[a-zA-Z]")) {
	                result.append(word).append(" ");
	            }
	        }

	        // Remove trailing space
	        return result.toString().trim();
	    }



}
